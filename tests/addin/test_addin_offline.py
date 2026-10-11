"""Offline checks for the M2 add-in source (no Excel is started).

VBA cannot run outside Excel, so these tests check what can be proven statically: the ribbon XML
and its callbacks are wired to real VBA procedures with the right signatures, the command registry
matches the ribbon, the VBA sources import cleanly (CRLF, ASCII, Option Explicit, names), and the
convert/undo code uses no clipboard and launches nothing. The .xlam packaging step is exercised on
a real Office package built with openpyxl. Behaviour in Excel is proven by the live harness.
"""

from __future__ import annotations

import re
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
CALLBACK_ARITY = {"onLoad": 1, "onAction": 1, "getLabel": 2, "getScreentip": 2, "getSupertip": 2, "getEnabled": 2}


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


def _public_procedures() -> dict[str, tuple[str, int, str]]:
    """name -> (kind, parameter count, file) for every Public Sub/Function in the add-in project."""
    found = {}
    for path in ALL_VBA:
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
            name = element.get(attribute)
            if name:
                used.add(name)
                assert name in procedures, f"{attribute}={name} has no Public procedure"
                kind, count, _ = procedures[name]
                assert kind == "Sub" and count == arity, f"{name}: expected Sub with {arity} parameter(s), got {kind}/{count}"
    assert {"Tk_RibbonLoaded", "Tk_OnAction", "Tk_GetLabel", "Tk_GetEnabled"} <= used


def test_ribbon_buttons_match_the_command_registry():
    buttons = [b.get("id") for b in _ui().iter(f"{UI_NS}button")]
    main = _code(ROOT / "addin" / "vba" / "TkMain.bas")
    registry_cases = set(re.findall(r'Case "(tk\w+)"\s*\n\s*Registry = Array', main))
    listed = re.search(r"Tk_CommandIds = Array\(([^)]*)\)", main).group(1)
    listed_ids = re.findall(r'"(\w+)"', listed)
    on_action_cases = set(re.findall(r'Case "(tk\w+)": Tk_', main))
    assert buttons == listed_ids == ["tkConvertToValues", "tkUndo", "tkAudit"]
    assert set(buttons) == registry_cases == on_action_cases


def test_every_button_has_a_built_in_icon_and_a_keytip():
    for button in _ui().iter(f"{UI_NS}button"):
        assert button.get("imageMso") and button.get("keytip"), button.get("id")
        assert button.get("image") is None  # built-in Office icons only


# ---------- VBA project ----------

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
