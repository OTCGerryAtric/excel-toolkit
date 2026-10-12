"""Offline checks for the M2 add-in source (no Excel is started).

VBA cannot run outside Excel, so these tests check what can be proven statically: the ribbon XML
and its callbacks are wired to real VBA procedures with the right signatures, the command registry
matches the ribbon, the VBA sources import cleanly (CRLF, ASCII, Option Explicit, names), and the
convert/undo code uses no clipboard and launches nothing. The .xlam packaging step is exercised on
a real Office package built with openpyxl. Behaviour in Excel is proven by the live harness.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook

import build_addin

ROOT = Path(__file__).resolve().parents[2]
ADDIN_VBA = sorted((ROOT / "addin" / "vba").iterdir())
ALL_VBA = build_addin.VBA_SOURCES
UI_NS = "{http://schemas.microsoft.com/office/2009/07/customui}"
CALLBACK_ARITY = {"onLoad": 1, "onAction": 1, "getLabel": 2, "getScreentip": 2, "getSupertip": 2, "getEnabled": 2,
                  "getPressed": 2}


def _text(path: Path) -> str:
    return path.read_bytes().decode("ascii")


def _code(path: Path) -> str:
    """Source without comments (good enough for VBA: no ' inside our string literals is followed by code)."""
    lines = []
    for line in _text(path).splitlines():
        out, in_string = [], False
        for ch in line:
            if ch == '"':
                in_string = not in_string
            if ch == "'" and not in_string:
                break
            out.append(ch)
        lines.append("".join(out))
    return "\n".join(lines)


STANDARD_MODULES = [p for p in ALL_VBA if p.suffix == ".bas"]  # class members are scoped to their class


def _public_procedures() -> dict[str, tuple[str, int, str]]:
    """name -> (kind, parameter count, file) for every Public Sub/Function in the project's standard
    modules - the global namespace that callbacks, OnKey/OnUndo and unqualified calls resolve against."""
    found = {}
    for path in STANDARD_MODULES:
        for m in re.finditer(r"^Public (Sub|Function) (\w+)\(([^)]*)\)", _code(path), re.M):
            kind, name, params = m.groups()
            assert name not in found, f"{name} is defined twice ({found.get(name)} and {path.name}) - ambiguous name"
            count = 0 if not params.strip() else params.count(",") + 1
            found[name] = (kind, count, path.name)
    return found


def _ui():
    return ET.fromstring(build_addin.CUSTOM_UI.read_bytes())


# ---------- Ribbon ----------

def test_ribbon_xml_uses_the_2009_namespace_and_one_tab():
    root = _ui()
    assert root.tag == f"{UI_NS}customUI"
    tabs = root.findall(f".//{UI_NS}tab")
    assert [t.get("label") for t in tabs] == ["Excel Toolkit"]


def test_every_ribbon_callback_is_a_public_sub_with_the_right_signature():
    procedures = _public_procedures()
    used = set()
    for element in _ui().iter():
        for attribute, arity in CALLBACK_ARITY.items():
            if attribute == "onAction" and element.tag == f"{UI_NS}toggleButton":
                arity = 2  # toggleButton onAction(control, pressed)
            name = element.get(attribute)
            if name:
                used.add(name)
                assert name in procedures, f"{attribute}={name} has no Public procedure"
                kind, count, _ = procedures[name]
                assert kind == "Sub" and count == arity, f"{name}: expected Sub with {arity} parameter(s), got {kind}/{count}"
    assert {"Tk_RibbonLoaded", "Tk_OnAction", "Tk_GetLabel", "Tk_GetEnabled"} <= used


def _controls(tags=("button", "toggleButton")):
    return [e for e in _ui().iter() if e.tag in {f"{UI_NS}{tag}" for tag in tags}]


def test_ribbon_controls_match_the_command_registry():
    main = _code(ROOT / "addin" / "vba" / "TkMain.bas")
    controls = [c.get("id") for c in _controls()]
    buttons = {c.get("id") for c in _controls(("button",))}
    registry_cases = set(re.findall(r'Case "(tk\w+)"\s*\n\s*Registry = Array', main))
    listed_ids = re.findall(r'"(\w+)"', re.search(r"Tk_CommandIds = Array\(([^)]*)\)", main).group(1))
    on_action_cases = set(re.findall(r'Case "(tk\w+)": Tk_', main))
    assert sorted(controls) == sorted(listed_ids)
    assert set(controls) == registry_cases
    assert buttons == on_action_cases  # the toggle goes through Tk_OnToggle instead
    assert "tkShortcuts" in controls and "tkShortcuts" not in buttons


def test_every_control_has_a_keytip_and_top_level_controls_a_built_in_icon():
    for group in _ui().iter(f"{UI_NS}group"):
        for control in group:
            assert control.get("image") is None  # built-in Office icons only
            if control.tag != f"{UI_NS}separator":
                assert control.get("imageMso") and control.get("keytip"), control.get("id")
    for control in _controls():
        assert control.get("keytip"), control.get("id")


def test_shortcuts_are_off_by_default_and_bound_keys_have_handlers():
    shortcuts = _code(ROOT / "addin" / "vba" / "TkShortcuts.bas")
    assert re.search(r"^Private enabled As Boolean", shortcuts, re.M)  # VBA initialises it to False
    assert "Tk_SetShortcuts True" not in "".join(_code(p) for p in ALL_VBA)  # never switched on by code
    procedures = _public_procedures()
    bindings = re.findall(r'Array\("(\^\+[A-Z])", "(\w+)"\)', shortcuts)
    assert [key for key, _ in bindings] == ["^+I", "^+O", "^+P", "^+M"]
    for _, handler in bindings:
        assert procedures[handler][:2] == ("Sub", 0)
    assert re.search(r"Public Sub Auto_Close\(\)\s*\n\s*If enabled Then Tk_SetShortcuts False", shortcuts)


# ---------- Legacy equivalence (M3a) ----------

LEGACY = ROOT / "baseline" / "vba" / "PERSONAL_XLSB_olevba.txt"


def _legacy_transitions(macro: str) -> tuple[dict, str]:
    """From the legacy If/Else chain: {condition format: next format} and the default (Else) format."""
    text = LEGACY.read_text(encoding="utf-8")
    body = text.split(f"Sub {macro}()", 1)[1].split("End Sub", 1)[0]
    conditions = re.findall(r'If Selection\.NumberFormat = "([^"]*)" Then\s*\n\s*Selection\.NumberFormat = "([^"]*)"', body)
    assigned = re.findall(r'Selection\.NumberFormat = "([^"]*)"\s*\n', body)
    default = assigned[-1]  # the innermost Else
    return dict(conditions), default


def _toolkit_cycle(name: str) -> list[str]:
    code = _code(ROOT / "addin" / "vba" / "TkFormat.bas")
    block = code.split(f'Case "{name}"', 1)[1].split("Case ", 1)[0]
    return [s.replace("{GBP}", "£") for s in re.findall(r'"([^"]*)"', block)]


@pytest.mark.parametrize("macro,cycle", [("Number_Format_Toggle", "number"), ("Currency_Format_Toggle", "currency"),
                                         ("Percentage_Format_Toggle", "percent"), ("Month_Date_Format", "date")])
def test_format_cycles_reproduce_the_legacy_macros(macro, cycle):
    """The toolkit rule (cycle[i] -> cycle[i+1] for i < last, anything else -> cycle[0]) applied to the
    toolkit's data gives exactly the legacy macro's transitions and default."""
    transitions, default = _legacy_transitions(macro)
    formats = _toolkit_cycle(cycle)
    assert formats[0] == default
    assert transitions == {formats[i]: formats[i + 1] for i in range(len(formats) - 1)}


UNDO_RECORD_INTERFACE = {"Description": 0, "CheckIdentity": 0, "FindConflict": 0, "Restore": 1}


@pytest.mark.parametrize("name", ["TkCellSnapshot", "TkFormatSnapshot", "TkStyleSnapshot", "TkFormulaSnapshot"])
def test_snapshot_classes_provide_the_undo_record_interface(name):
    """TkUndo holds records late-bound (As Object), so the compiler cannot check this - the test does."""
    code = _code(ROOT / "addin" / "vba" / f"{name}.cls")
    for method, arity in UNDO_RECORD_INTERFACE.items():
        m = re.search(rf"^Public Function {method}\(([^)]*)\) As String", code, re.M)
        assert m, f"{name} lacks Public Function {method}(...) As String"
        assert (0 if not m.group(1).strip() else m.group(1).count(",") + 1) == arity, method


# ---------- Legacy equivalence (M3b) ----------

def _legacy_body(macro: str) -> str:
    return LEGACY.read_text(encoding="utf-8").split(f"Sub {macro}()", 1)[1].split("End Sub", 1)[0]


def _style_code() -> str:
    return _code(ROOT / "addin" / "vba" / "TkStyle.bas")


def _vba_function(code: str, name: str) -> str:
    return code.split(f"Public Function {name}(", 1)[1].split("End Function", 1)[0]


FILLS = {"Yellow_Cell_Toggle": "yellow", "Green_Cell_Toggle": "green", "Light_Blue_Cell_Toggle": "lightblue",
         "Grey_Cell_Toggle": "grey"}


@pytest.mark.parametrize("macro,name", FILLS.items())
def test_fill_toggles_use_the_legacy_colours_and_switch_off_to_no_fill(macro, name):
    """Same colour and decision (Interior.Color equal to the colour -> off) as the legacy macro; the one
    documented fix is that "off" is No Fill rather than white."""
    body = _legacy_body(macro)
    checked = re.findall(r"If Selection\.Interior\.Color = RGB\((\d+), (\d+), (\d+)\) Then", body)
    assigned = re.findall(r"Selection\.Interior\.Color = RGB\((\d+), (\d+), (\d+)\)\s*\n", body)
    assert len(checked) == 1 and assigned == [("255", "255", "255"), checked[0]]  # off = white, on = colour
    toolkit = re.search(rf'Case "{name}": Tk_FillColour = RGB\((\d+), (\d+), (\d+)\)', _style_code())
    assert toolkit and toolkit.groups() == checked[0]
    decide = _style_code().split("Private Function Decide(", 1)[1]
    assert "If current = colour Then" in decide and "newValue = NO_FILL" in decide
    assert "area.Interior.Pattern = xlNone" in _style_code()


def test_toolkit_never_fills_white():
    for path in ADDIN_VBA:
        assert "RGB(255, 255, 255)" not in _code(path), path.name


def test_outline_reproduces_the_legacy_borders():
    body = _legacy_body("Grey_Ouline")
    legacy_on = re.findall(r"With Selection\.Borders\((\w+)\)\s*\n\s*\.LineStyle = (\w+)\s*\n\s*\.ThemeColor = (\d+)"
                           r"\s*\n\s*\.TintAndShade = (-?[\d.]+)\s*\n\s*\.Weight = (\w+)", body)
    legacy_off = re.findall(r"Selection\.Borders\((\w+)\)\.LineStyle = xlNone", body.split("Else", 1)[1])
    edges = [e for e, *_ in legacy_on]
    assert edges == legacy_off == ["xlEdgeLeft", "xlEdgeTop", "xlEdgeBottom", "xlEdgeRight", "xlInsideVertical",
                                   "xlInsideHorizontal"]
    assert {tuple(rest) for _, *rest in legacy_on} == {("xlContinuous", "1", "-0.249946592608417", "xlThin")}
    assert re.search(r"If Selection\.Borders\(xlEdgeTop\)\.LineStyle = xlNone Then", body)

    code = _style_code()
    apply = code.split("Private Sub ApplyChange(", 1)[1].split("End Sub", 1)[0]
    assert re.search(r"For Each edge In Array\(" + r", ".join(edges) + r"\)", apply)
    assert re.search(r"\.LineStyle = xlContinuous\s*\n\s*\.ThemeColor = OUTLINE_THEME\s*\n\s*"
                     r"\.TintAndShade = OUTLINE_TINT\s*\n\s*\.Weight = xlThin", apply)  # legacy order
    assert "Private Const OUTLINE_THEME As Long = 1" in code
    assert "Private Const OUTLINE_TINT As Double = -0.249946592608417" in code
    assert "current = target.Borders(xlEdgeTop).LineStyle" in code and "newValue = (current = xlNone)" in code


def test_row_height_and_column_width_toggles_follow_the_legacy_rules_with_the_sheet_standard_size():
    """Legacy: 15 <-> 4.5 and 2 -> 8.43 -> 12 -> 2, where 15 and 8.43 are Excel's usual standard row height
    and column width. Documented change (owner, 2026-10-12): the sheet's own StandardHeight/StandardWidth
    takes their place, matched within half a pixel; the other values are unchanged."""
    row = _legacy_body("Row_Height_Toggle")
    assert re.findall(r"Selection\.RowHeight = ([\d.]+)", row) == ["15", "4.5", "15"]
    col = _legacy_body("Column_Width_Toggle")
    assert re.findall(r"Selection\.ColumnWidth = ([\d.]+)", col) == ["2", "8.43", "8.43", "12", "2"]
    code = _style_code()
    rows = _vba_function(code, "Tk_NextRowHeight")
    assert "If Abs(current - standardHeight) <= ROW_MATCH Then Tk_NextRowHeight = 4.5" in rows
    assert re.search(r"Tk_NextRowHeight = standardHeight\s*$", rows)
    cols = _vba_function(code, "Tk_NextColumnWidth")
    assert re.findall(r"If Abs\(current - (\w+)\) <= COLUMN_MATCH Then Tk_NextColumnWidth = (\w+)", cols) == \
        [("2", "standardWidth"), ("standardWidth", "12")]
    assert re.search(r"Tk_NextColumnWidth = 2\s*$", cols)
    assert "Tk_NextRowHeight(target.RowHeight, target.Worksheet.StandardHeight)" in code
    assert "Tk_NextColumnWidth(target.ColumnWidth, target.Worksheet.StandardWidth)" in code


def test_best_fit_uses_the_legacy_expression_and_integer_rounding():
    body = _legacy_body("Best_Fit_Column_Width")
    assert "Dim CColWidth As Integer" in body
    assert "CColWidth = Round(ActiveCell.ColumnWidth + 5) / 5" in body
    assert "Selection.ColumnWidth = CColWidth * 5" in body
    fn = _vba_function(_style_code(), "Tk_BestFitWidth")
    assert "Dim steps As Integer" in fn
    assert "steps = Round(activeWidth + 5) / 5" in fn and "Tk_BestFitWidth = steps * 5" in fn
    assert "Tk_BestFitRange(Selection, ActiveCell)" in _style_code()


def test_size_undo_captures_whole_rows_and_columns_and_refuses_hidden_ones():
    snap = _code(ROOT / "addin" / "vba" / "TkStyleSnapshot.cls")
    assert "ws.Range(ws.Cells(area.Row, 1), ws.Cells(area.Row + area.Rows.Count - 1, 1))" in snap
    assert "ws.Range(ws.Cells(1, area.Column), ws.Cells(1, area.Column + area.Columns.Count - 1))" in snap
    assert re.search(r"If size = 0 Then", snap)  # hidden rows/columns read as 0 and are refused


def test_border_undo_reads_every_cell_and_each_shared_edge_once():
    """For several cells Excel reports the first cell's edge, never "mixed" (M3b live run), so borders are
    read cell by cell: each cell's top and left, then the last row's bottoms and the last column's rights
    (the facing edges of the cells around the selection)."""
    snap = _code(ROOT / "addin" / "vba" / "TkStyleSnapshot.cls")
    rows = snap.split("Private Function AddBorderBlocks(", 1)[1].split("End Function", 1)[0]
    assert rows.index("Array(xlEdgeTop, xlEdgeLeft)") < rows.index("Array(xlEdgeBottom)") < rows.index("Array(xlEdgeRight)")
    state = snap.split("Private Function BorderState(", 1)[1].split("End Function", 1)[0]
    assert "For Each cell In rng.Cells" in state and "CellBorderState(cell, cellEdges)" in state
    assert "CellBorderState(strip.Cells(i), cellEdges)" in snap
    assert "MAX_BORDER_CELLS" in snap


def test_style_snapshot_splits_ranges_by_corner_cells():
    """Rows(n)/Columns(n) return a range that counts rows/columns, not cells (M3b live run: a mixed
    selection with a merged cell was refused); Offset runs off the sheet on whole columns."""
    snap = _code(ROOT / "addin" / "vba" / "TkStyleSnapshot.cls")
    assert not any(token in snap for token in (".Offset(", ".Resize(", ".Rows(", ".Columns("))
    assert "Set SubRange = ws.Range(rng.Cells(r1, c1), rng.Cells(r2, c2))" in snap
    assert ".CountLarge = 1" not in snap.replace("rng.Cells.CountLarge = 1", "")


def test_style_undo_refuses_gradients_and_never_compares_a_theme_marker_with_equals():
    snap = _code(ROOT / "addin" / "vba" / "TkStyleSnapshot.cls")
    assert "xlPatternLinearGradient" in snap and "xlPatternRectangularGradient" in snap
    # theme is a number or the NO_THEME string: `theme = NO_THEME` would be a Type mismatch on numbers
    assert not re.search(r"\b\w*[Tt]heme\w*\(?\w*\)?\s*=\s*NO_THEME\b(?!\s*$)", snap.replace("ThemeOf = NO_THEME", ""))
    assert "IsRgb(theme)" in snap


# ---------- Clear Sheet References and Find Constants (M3c) ----------

sys.path.insert(0, str(ROOT / "addin" / "live"))
from m3c_cases import AMBIGUOUS_CASES, REWRITE_CASES  # noqa: E402
from m3c_rewrite_model import _is_cell_reference, remove_sheet_prefix, rewrite_sheet_prefix  # noqa: E402


def _formulas_code() -> str:
    return _code(ROOT / "addin" / "vba" / "TkFormulas.bas")


@pytest.mark.parametrize("sheet,before,after", REWRITE_CASES, ids=[f"{s}|{b}" for s, b, _ in REWRITE_CASES])
def test_clear_sheet_references_rewrite_rules(sheet, before, after):
    assert remove_sheet_prefix(before, sheet) == after
    assert remove_sheet_prefix(after, sheet) == after  # idempotent: what VerifyApplied relies on


@pytest.mark.parametrize("sheet,formula", AMBIGUOUS_CASES, ids=[f"{s}|{f}" for s, f in AMBIGUOUS_CASES])
def test_formulas_excel_could_read_two_ways_are_left_unchanged_and_reported(sheet, formula):
    assert rewrite_sheet_prefix(formula, sheet) == (formula, True)


def test_cell_reference_bounds_follow_excel():
    for token in ("A1", "$A$1", "A$1", "$A1", "XFD1", "XFD1048576", "b2", "AB12"):
        assert _is_cell_reference(token), token
    for token in ("A", "AB", "XFD", "1", "$1", "A0", "XFE1", "XFD1048577", "ABCD1", "A1B", "", "$", "R1C1"):
        assert not _is_cell_reference(token), token


def test_the_legacy_text_replace_damages_what_the_toolkit_keeps():
    """The legacy macro: Cells.Replace "<name>!" (or "'<name>'!" if the name has a space) anywhere in any
    cell, case-insensitive. On these cases it changes text the toolkit rightly leaves alone."""
    def legacy(formula, sheet):
        what = f"'{sheet}'!" if " " in sheet else f"{sheet}!"
        return re.sub(re.escape(what), "", formula, flags=re.I)
    damaged = [(s, b) for s, b, a in REWRITE_CASES if legacy(b, s) != a]
    for case in [("Sheet1", "=MySheet1!A1"), ("Sheet1", "=[Book.xlsx]Sheet1!A1"), ("Sheet1", '="Sheet1!A1"'),
                 ("Sheet3", "=SUM(Sheet1:Sheet3!A1)"), ("REF", "=REF!A1+#REF!"), ("A", "=SUM(AB:A!B2)"),
                 ("Sheet1", "=SUM(Sheet2!A1:Sheet1!B2)")]:
        assert case in damaged
    assert legacy("Sales note: see Sheet1!A1", "Sheet1") != "Sales note: see Sheet1!A1"  # and text cells


def test_rewrite_rules_in_vba_match_the_model():
    code = _formulas_code()
    fn = _vba_function(code, "Tk_RewriteSheetPrefix")
    assert "result = RemovePrefixes(formula, sheetName)" in fn
    assert "If HasAmbiguousColon(formula) Or HasAmbiguousColon(result) Then" in fn and "result = formula" in fn
    remove = code.split("Private Function RemovePrefixes(", 1)[1].split("End Function", 1)[0]
    assert 'j = QuoteEnd(formula, i, """")' in remove and "j = QuoteEnd(formula, i, \"'\")" in remove
    assert "OwnPrefix(formula, i, Replace(Mid$(formula, i + 1, j - i - 1), \"''\", \"'\"), sheetName)" in remove
    assert "OwnPrefix(formula, i, Mid$(formula, i, j - i + 1), sheetName)" in remove
    own = code.split("Private Function OwnPrefix(", 1)[1].split("End Function", 1)[0]
    assert "StrComp(nameText, sheetName, vbTextCompare) <> 0" in own
    assert 'If prevCh = "]" Or prevCh = "#" Then Exit Function' in own
    assert "If Not RangeHalfBefore(formula, i - 1, sheetName) Then Exit Function" in own
    half = code.split("Private Function RangeHalfBefore(", 1)[1].split("End Function", 1)[0]
    assert "If Not IsCellReference(Mid$(formula, k, colon - k)) Then Exit Function" in half
    assert "StrComp(PrefixNameBefore(formula, k - 1), sheetName, vbTextCompare) = 0" in half
    cell = code.split("Private Function IsCellReference(", 1)[1].split("End Function", 1)[0]
    assert "colNumber < 1 Or colNumber > 16384" in cell and "CLng(rowText) <= 1048576" in cell
    assert 'If Left$(rowText, 1) = "0" Then Exit Function' in cell
    assert 'IsNameChar = (ch Like "[A-Za-z0-9_.]") Or code > 127' in code
    snap = _code(ROOT / "addin" / "vba" / "TkFormulaSnapshot.cls")
    assert "newText = Tk_RewriteSheetPrefix(oldText, sheetName, ambiguous)" in snap
    assert "If snap.AmbiguousCount > 0 Then" in code and "Check those by hand." in code


def test_clear_sheet_references_reads_and_writes_formula_cells_only():
    snap = _code(ROOT / "addin" / "vba" / "TkFormulaSnapshot.cls")
    assert "scope.SpecialCells(xlCellTypeFormulas)" in snap
    assert "If Not area.Cells(r, c).HasFormula Then GoTo NextCell" in snap
    assert ".HasArray Then" in snap  # array formulas refused, not rewritten cell by cell
    for path in ADDIN_VBA:
        assert not re.search(r"\.Replace\b", _code(path)), f"{path.name}: no Range.Replace (the legacy text replace)"
    main = _formulas_code()
    assert "If target.CountLarge = 1 Then Set scope = target.Worksheet.UsedRange Else Set scope = target" in main


def test_find_constants_is_the_legacy_selection_and_changes_nothing_else():
    legacy = _legacy_body("Find_Constants")
    assert "Selection.SpecialCells(xlCellTypeConstants, 1).Select" in legacy  # 1 = xlNumbers
    fn = _vba_function(_formulas_code(), "Tk_FindConstants")
    assert "target.SpecialCells(xlCellTypeConstants, xlNumbers)" in fn and "found.Select" in fn
    assert "Tk_SetRecord" not in fn and "Formula" not in fn and ".Value" not in fn


# ---------- Clipboard Paste Values experiment (M3d, prepared only) ----------

M3D_PROBE = ROOT / "addin" / "live" / "m3d_paste_probe.bas"
M3D_HARNESS = ROOT / "addin" / "live" / "m3d_native_paste_experiment.py"


def test_m3d_probe_tests_the_native_command_and_stays_out_of_the_addin():
    raw = M3D_PROBE.read_bytes()
    assert raw.isascii() and raw.count(b"\r\n") == raw.count(b"\n")
    code = _code(M3D_PROBE)
    native = _vba_function(code, "Probe_NativePasteValues")
    assert 'Application.CommandBars.ExecuteMso "PasteValues"' in native and "PasteSpecial" not in native
    assert 'GetEnabledMso("PasteValues")' in native
    # the object-model paste appears only as the labelled control
    assert code.count(".PasteSpecial") == 1 and ".PasteSpecial" in _vba_function(code, "Probe_ControlPasteSpecial")
    assert M3D_PROBE.name not in (ROOT / "addin" / "build_addin.py").read_text(encoding="utf-8")
    assert not (ROOT / "addin" / "vba" / M3D_PROBE.name).exists()
    assert not any("PasteValues" in _code(path) for path in ADDIN_VBA)  # no Paste Values in the add-in yet


def test_m3d_harness_uses_a_disposable_excel_only():
    text = M3D_HARNESS.read_text(encoding="utf-8")
    assert 'DispatchEx("Excel.Application")' in text and "GetActiveObject" not in text
    assert "DEFAULT_XLAM" not in text and "PERSONAL" in text  # no add-in; PERSONAL.XLSB only hashed
    assert "SaveAs" not in text and "Close(SaveChanges=False)" in text
    assert "DO NOT RUN until authorised" in text


# ---------- VBA project ----------

def _module_level_names(path: Path) -> list[str]:
    """Names declared at module level: procedures, Consts, variables and Declares."""
    names, in_else = [], False
    for line in _code(path).splitlines():
        if line.startswith("#Else"):
            in_else = True  # the #Else branch of #If VBA7 re-declares the same names
        elif line.startswith("#End If"):
            in_else = False
        if in_else or line[:1].isspace():
            continue  # procedure bodies are indented in this project
        m = re.match(r"(?:(?:Public|Private|Friend) )?(?:Sub|Function|Property (?:Get|Let|Set)) (\w+)", line)
        if m:
            if m.group(0).split()[-2:-1] not in (["Let"], ["Set"]):  # Property Let/Set share the Get's name
                names.append(m.group(1))
            continue
        m = re.match(r"(?:Public|Private|Global) Const (\w+)|(?:Public|Private|Dim|Global) (?!Const|Declare|Sub|Function|Property)(\w+)", line)
        if m:
            names.append(m.group(1) or m.group(2))
        m = re.match(r"(?:(?:Public|Private) )?Declare (?:PtrSafe )?(?:Function|Sub) (\w+)", line)
        if m:
            names.append(m.group(1))
    return names


@pytest.mark.parametrize("path", ALL_VBA, ids=lambda p: p.name)
def test_module_level_names_are_unique_ignoring_case(path):
    """VBA is case-insensitive: Const TK_VERSION and Function Tk_Version are an 'Ambiguous name' compile error."""
    names = [n.lower() for n in _module_level_names(path)]
    duplicates = sorted({n for n in names if names.count(n) > 1})
    assert not duplicates, f"{path.name}: ambiguous names {duplicates}"


def test_public_names_are_unique_across_the_project_ignoring_case():
    seen = {}
    for path in STANDARD_MODULES:
        for line in _code(path).splitlines():
            m = re.match(r"Public (?:Sub|Function|Const) (\w+)", line)
            if m:
                key = m.group(1).lower()
                assert key not in seen, f"{m.group(1)} in {path.name} clashes with {seen[key]}"
                seen[key] = f"{m.group(1)} in {path.name}"


def test_cross_module_calls_resolve_to_public_procedures():
    procedures = _public_procedures()
    for path in ALL_VBA:
        for name in set(re.findall(r"\b((?:Tk_|PhaseB_)\w+|SafeHasSpill)\b", _code(path))):
            assert name in procedures, f"{path.name} calls {name}, which is not a Public procedure"


def test_ctrl_z_handler_is_registered_by_name_and_takes_no_arguments():
    undo = _code(ROOT / "addin" / "vba" / "TkUndo.bas")
    registered = re.search(r'Application\.OnUndo "Undo " & label, "(\w+)"', undo).group(1)
    kind, count, _ = _public_procedures()[registered]
    assert (kind, count) == ("Sub", 0)


@pytest.mark.parametrize("path", ADDIN_VBA, ids=lambda p: p.name)
def test_vba_source_imports_cleanly(path):
    raw = path.read_bytes()
    raw.decode("ascii")  # the VBA editor imports in the ANSI code page; keep sources ASCII
    assert raw.count(b"\n") == raw.count(b"\r\n"), "CRLF line endings required (LF .cls imports as a module)"
    text = _text(path)
    assert re.search(rf'^Attribute VB_Name = "{path.stem}"\r?$', text, re.M)
    assert re.search(r"^Option Explicit\r?$", text, re.M)
    if path.suffix == ".cls":
        assert text.startswith("VERSION 1.0 CLASS\r\nBEGIN\r\n")
    assert max(len(line) for line in text.splitlines()) <= 1023
    continuation = 0
    for line in text.splitlines():
        continuation = continuation + 1 if line.rstrip().endswith(" _") else 0
        assert continuation <= 24, "VBA allows at most 24 line continuations"


@pytest.mark.parametrize("path", ALL_VBA, ids=lambda p: p.name)
def test_module_declarations_precede_the_first_procedure(path):
    """VBA rejects module-level Dim/Const/Declare/variables after a procedure ("Only comments may appear...")."""
    seen_procedure = False
    for line in _code(path).splitlines():
        if re.match(r"(Public |Private |Friend )?(Sub|Function|Property (Get|Let|Set)) \w+", line):
            seen_procedure = True
        elif seen_procedure and re.match(r"(Public|Private|Dim|Const|Global)\b(?! (Sub|Function|Property)\b)|(Private |Public )?Declare\b", line):
            pytest.fail(f"{path.name}: module-level declaration after a procedure: {line.strip()}")


@pytest.mark.parametrize("path", ADDIN_VBA, ids=lambda p: p.name)
def test_type_checks_are_not_combined_with_and(path):
    """VBA's And does not short-circuit: `VarType(x) = vbString And Len(x) > 0` still calls Len on an
    error value and raises Type mismatch (found in the first M2 live run). Nest the test instead."""
    for number, line in enumerate(_code(path).splitlines(), start=1):
        assert not re.search(r"VarType\((?:[^()]|\([^()]*\))*\)\s*=\s*vb\w+\s+And\b", line), f"{path.name}:{number}: {line.strip()}"


@pytest.mark.parametrize("path", ADDIN_VBA, ids=lambda p: p.name)
def test_no_clipboard_and_nothing_launched(path):
    code = _code(path)
    for pattern in (r"\.Copy\b", r"PasteSpecial", r"CutCopyMode", r"DataObject", r"\bShell\b", r"WScript",
                    r"ShellExecute", r"CreateProcess", r"\bKill\b", r"For Output", r"SaveSetting", r"RegWrite"):
        assert not re.search(pattern, code), f"{path.name} uses {pattern}"


def test_build_imports_every_addin_module_and_phase_b():
    names = {p.stem for p in ALL_VBA}
    assert names == set(build_addin.EXPECTED_TYPES)
    assert {p.stem for p in ADDIN_VBA} <= names


def test_crlf_copies_normalise_line_endings(tmp_path, monkeypatch):
    monkeypatch.setattr(build_addin, "BUILD", tmp_path)
    for copy in build_addin.crlf_copies():
        raw = copy.read_bytes()
        assert raw.count(b"\n") == raw.count(b"\r\n") and b"\r\r" not in raw


# ---------- Packaging ----------

@pytest.fixture
def package(tmp_path):
    path = tmp_path / "book.xlsx"
    wb = Workbook()
    wb.active["A1"] = 1
    wb.save(path)
    return path


def _parts(path):
    with zipfile.ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}


def test_inject_adds_custom_ui_part_and_relationship(package):
    before = _parts(package)
    ui = build_addin.CUSTOM_UI.read_bytes()
    build_addin.inject_custom_ui(package, ui)
    after = _parts(package)
    assert build_addin.verify_package(package, ui, expect_vba=False) == []
    assert after[build_addin.UI_PART] == ui
    untouched = set(before) - {"_rels/.rels", "[Content_Types].xml"}
    assert all(after[name] == before[name] for name in untouched)
    rels = ET.fromstring(after["_rels/.rels"])
    originals = ET.fromstring(before["_rels/.rels"])
    assert len(rels) == len(originals) + 1  # existing relationships kept


def test_inject_is_idempotent(package):
    ui = build_addin.CUSTOM_UI.read_bytes()
    build_addin.inject_custom_ui(package, ui)
    build_addin.inject_custom_ui(package, ui)
    assert build_addin.verify_package(package, ui, expect_vba=False) == []


def test_verify_reports_a_package_without_vba_or_ribbon(package):
    problems = build_addin.verify_package(package, build_addin.CUSTOM_UI.read_bytes())
    assert any("customUI14.xml" in p for p in problems)
    assert any("vbaProject.bin" in p for p in problems)


def test_malformed_ribbon_xml_is_refused_without_touching_the_package(package):
    before = package.read_bytes()
    with pytest.raises(ET.ParseError):
        build_addin.inject_custom_ui(package, b"<customUI><broken></customUI>")
    assert package.read_bytes() == before
