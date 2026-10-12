"""M3b scripted live check - fills, Grey Outline and sizes vs the legacy macros. DO NOT RUN until authorised.

Same boundaries as m2/m3a: a SEPARATE disposable Excel process (DispatchEx: no XLSTART, no PERSONAL.XLSB,
no installed add-ins), workbooks it creates and closes without saving, the development add-in opened with
Workbooks.Open, nothing installed. The legacy macros are imported from baseline/vba into the disposable
workbook - never from the real PERSONAL.XLSB.

Each command runs side by side with its legacy macro on identical fixtures and is compared property by
property (fill: pattern, colour, theme colour, tint, pattern colour; borders: every edge's style, weight
and colour of every cell in the fixture, including the cells around the selection; sizes: every row
height / column width in the fixture). The one documented difference - a fill toggled off is No Fill,
not white - is checked explicitly. Every toolkit change is then undone and compared with the state
before it.

Diagnostics (recorded, not pass/fail) answer the open questions in docs/validation/m3b-fills-borders-sizes.md:
what ThemeColor reports for RGB / no-fill / theme / mixed ranges, whether adjacent cells share edges,
and how heights and widths read back.

  <scratch venv>/python addin/live/m3b_live_check.py [--xlam addin/build/ExcelToolkit.xlam] [--perf]
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

from m2_live_check import DEFAULT_XLAM, PERSONAL, sha256
from m3a_live_check import Check as M3aCheck

HERE = Path(__file__).resolve().parent
FIXTURE = "A1:G12"
SELECTIONS = ["B2", "B2:D6", "B2:B4,D6:E8"]
EDGES = (7, 8, 9, 10)  # xlEdgeLeft, Top, Bottom, Right
X_NONE, X_SOLID, X_CONTINUOUS, X_THIN, X_MEDIUM, X_DASH, X_GRAY50 = -4142, 1, 1, 2, -4138, -4115, -4125
TINT = -0.249946592608417


def rgb(r, g, b):
    return r + 256 * g + 65536 * b


FILLS = {"yellow": ("Yellow_Cell_Toggle", rgb(255, 255, 210)), "green": ("Green_Cell_Toggle", rgb(145, 245, 190)),
         "lightblue": ("Light_Blue_Cell_Toggle", rgb(221, 235, 247)), "grey": ("Grey_Cell_Toggle", rgb(242, 242, 242))}


def _try(obj, prop):
    try:
        return getattr(obj, prop)
    except Exception:  # noqa: BLE001 - recorded as data
        return "error"


def fill_state(cell):
    it = cell.Interior
    return (it.Pattern, it.Color, _try(it, "ThemeColor"), it.TintAndShade, it.PatternColorIndex, it.PatternColor,
            _try(it, "PatternThemeColor"), it.PatternTintAndShade)


def border_state(cell):
    out = []
    for edge in EDGES:
        b = cell.Borders(edge)
        out.append((b.LineStyle, b.Weight, b.ColorIndex, b.Color, _try(b, "ThemeColor"), b.TintAndShade))
    return tuple(out)


class Check(M3aCheck):
    # ----- snapshots and fixtures -----

    @staticmethod
    def fills(ws, address=FIXTURE):
        return {c.Address: fill_state(c) for c in ws.Range(address).Cells}

    @staticmethod
    def borders(ws, address=FIXTURE):
        return {c.Address: border_state(c) for c in ws.Range(address).Cells}

    @staticmethod
    def heights(ws, rows=range(1, 21)):
        return {r: ws.Rows(r).RowHeight for r in rows}

    @staticmethod
    def widths(ws, columns=range(1, 11)):
        return {c: ws.Columns(c).ColumnWidth for c in columns}

    @staticmethod
    def reset(ws):
        ws.Cells.ClearFormats()
        ws.Cells.UseStandardHeight = True
        ws.Cells.UseStandardWidth = True
        ws.Rows.Hidden = False
        ws.Columns.Hidden = False

    @staticmethod
    def setup_fill(ws, start, colour):
        Check.reset(ws)
        area = ws.Range("A1:F10")
        if start == "same":
            area.Interior.Color = colour
        elif start == "other rgb":
            area.Interior.Color = rgb(200, 100, 50)
        elif start == "theme":
            area.Interior.ThemeColor = 5
            area.Interior.TintAndShade = 0.4
        elif start == "pattern":
            area.Interior.Pattern = X_GRAY50
            area.Interior.Color = rgb(255, 0, 0)
            area.Interior.PatternColor = rgb(0, 0, 255)
        elif start == "mixed":
            for i, cell in enumerate(area.Cells):
                if i % 3 == 0:
                    cell.Interior.Color = colour
                elif i % 3 == 1:
                    cell.Interior.ThemeColor = 6
                    cell.Interior.TintAndShade = -0.25

    @staticmethod
    def setup_borders(ws, start):
        Check.reset(ws)
        if start == "outlined":
            for edge in (7, 8, 9, 10, 11, 12):
                b = ws.Range("B2:D6").Borders(edge)
                b.LineStyle, b.ThemeColor, b.TintAndShade, b.Weight = X_CONTINUOUS, 1, TINT, X_THIN
        elif start == "thick red box":
            for edge in EDGES:
                b = ws.Range("A1:E8").Borders(edge)
                b.LineStyle, b.Weight, b.Color = X_CONTINUOUS, X_MEDIUM, rgb(255, 0, 0)
        elif start == "neighbours":
            ws.Range("A1:G1").Borders(9).LineStyle = X_CONTINUOUS          # bottom of the row above
            b = ws.Range("F2:F9").Borders(7)                               # left of the column to the right
            b.LineStyle, b.Color = X_DASH, rgb(0, 0, 255)
            ws.Range("A7:G7").Borders(8).ThemeColor = 6                    # theme top edge below
        elif start == "mixed":
            ws.Range("B2").Borders(8).LineStyle = X_CONTINUOUS
            b = ws.Range("C2").Borders(8)
            b.LineStyle, b.ThemeColor, b.TintAndShade = X_CONTINUOUS, 7, 0.5
            ws.Range("D3:D5").Borders(10).Weight = X_MEDIUM

    def select(self, ws, selection):
        ws.Activate()
        ws.Range(selection).Select()

    def legacy_run(self, selection, macro):
        self.select(self.legacy, selection)
        self.run_legacy(macro)

    # ----- diagnostics (data, not pass/fail) -----

    def diagnostics(self):
        ws = self.toolkit
        self.reset(ws)
        ws.Range("A1").Interior.Color = rgb(255, 255, 210)
        ws.Range("A2:A3").Interior.Color = rgb(255, 255, 210)
        ws.Range("B1:B2").Interior.ThemeColor = 5
        ws.Range("C1").Interior.ThemeColor = 2          # white theme colour (light 1)
        ws.Range("C2").Interior.Color = rgb(255, 255, 255)
        ws.Range("E1").Borders(9).LineStyle = X_CONTINUOUS
        ws.Rows(5).RowHeight = 4.5
        ws.Columns("F").ColumnWidth = 8.43
        found = {
            "fill ThemeColor, RGB single": _try(ws.Range("A1").Interior, "ThemeColor"),
            "fill ThemeColor, RGB two cells": _try(ws.Range("A2:A3").Interior, "ThemeColor"),
            "fill ThemeColor, no fill": _try(ws.Range("D1").Interior, "ThemeColor"),
            "fill ThemeColor, theme two cells": _try(ws.Range("B1:B2").Interior, "ThemeColor"),
            "fill ThemeColor, theme white + RGB white": _try(ws.Range("C1:C2").Interior, "ThemeColor"),
            "fill Color, theme white + RGB white": _try(ws.Range("C1:C2").Interior, "Color"),
            "border ThemeColor, automatic": _try(ws.Range("E1").Borders(9), "ThemeColor"),
            "shared edge: E2 top after E1 bottom set": ws.Range("E2").Borders(8).LineStyle,
            "row height after 4.5": ws.Rows(5).RowHeight,
            "column width after 8.43": ws.Columns("F").ColumnWidth,
        }
        self.results.append({"check": "diagnostics", "passed": True, **{k: str(v) for k, v in found.items()}})
        for key, value in found.items():
            print(f"  diagnostic: {key}: {value}")

    # ----- equivalence and exact undo -----

    def fills_equivalence(self):
        mismatches, undo_failures, off_cases, cases = [], [], 0, 0
        for name, (macro, colour) in FILLS.items():
            for selection in SELECTIONS:
                for start in ["none", "same", "other rgb", "theme", "pattern", "mixed"]:
                    cases += 1
                    for ws in (self.legacy, self.toolkit):
                        self.setup_fill(ws, start, colour)
                    before = self.fills(self.toolkit)
                    self.legacy_run(selection, macro)
                    outcome = self.run("Tk_FillRange", self.toolkit.Range(selection), name)
                    legacy, toolkit = self.fills(self.legacy), self.fills(self.toolkit)
                    selected = {c.Address for c in self.toolkit.Range(selection).Cells}
                    if "removed" in outcome:  # the documented fix: No Fill where the legacy macro set white
                        off_cases += 1
                        expected_ok = all(toolkit[a][0] == X_NONE for a in selected) and \
                            all(legacy[a][:2] == (X_SOLID, rgb(255, 255, 255)) for a in selected)
                        same_elsewhere = all(toolkit[a] == legacy[a] for a in toolkit if a not in selected)
                        good = expected_ok and same_elsewhere
                    else:
                        good = outcome.startswith("ok:") and toolkit == legacy
                    if not good:
                        mismatches.append((name, selection, start, outcome))
                    undo = self.run("Tk_RestoreLast")
                    if not undo.startswith("ok:") or self.fills(self.toolkit) != before:
                        undo_failures.append((name, selection, start, undo))
        self.record(f"fills: toolkit matches the legacy macro property by property ({cases} cases; {off_cases} "
                    "switch-offs give No Fill instead of white)", not mismatches, mismatches=mismatches[:10])
        self.record("fills: every change undone exactly (RGB, theme + tint, patterns, mixed)", not undo_failures,
                    failures=undo_failures[:10])

    def outline_equivalence(self):
        mismatches, undo_failures, cases = [], [], 0
        for selection in SELECTIONS:
            for start in ["none", "outlined", "thick red box", "neighbours", "mixed"]:
                for presses in (1, 2):
                    cases += 1
                    for ws in (self.legacy, self.toolkit):
                        self.setup_borders(ws, start)
                    for _ in range(presses - 1):  # second press from the state after one press
                        self.legacy_run(selection, "Grey_Ouline")
                        self.run("Tk_OutlineRange", self.toolkit.Range(selection))
                    before = self.borders(self.toolkit)
                    self.legacy_run(selection, "Grey_Ouline")
                    outcome = self.run("Tk_OutlineRange", self.toolkit.Range(selection))
                    if not outcome.startswith("ok:") or self.borders(self.legacy) != self.borders(self.toolkit):
                        mismatches.append((selection, start, presses, outcome))
                    undo = self.run("Tk_RestoreLast")
                    if not undo.startswith("ok:") or self.borders(self.toolkit) != before:
                        undo_failures.append((selection, start, presses, undo))
        self.record(f"outline: toolkit matches the legacy macro edge by edge, incl. neighbouring cells ({cases} cases)",
                    not mismatches, mismatches=mismatches[:10])
        self.record("outline: every change undone exactly, incl. the neighbours' facing edges", not undo_failures,
                    failures=undo_failures[:10])

    def sizes_equivalence(self):
        mismatches, undo_failures, cases = [], [], 0
        row_cases = [(sel, start) for sel in SELECTIONS + ["3:5"] for start in ["default", "15", "4.5", "mixed"]]
        for selection, start in row_cases:
            cases += 1
            for ws in (self.legacy, self.toolkit):
                self.reset(ws)
                if start in ("15", "4.5"):
                    ws.Range("A1:A12").RowHeight = float(start)
                elif start == "mixed":
                    for r in range(2, 9):
                        ws.Rows(r).RowHeight = 15 if r % 2 else 30
            before = self.heights(self.toolkit)
            self.legacy_run(selection, "Row_Height_Toggle")
            outcome = self.run("Tk_RowHeightRange", self.toolkit.Range(selection))
            if not outcome.startswith("ok:") or self.heights(self.legacy) != self.heights(self.toolkit):
                mismatches.append(("row", selection, start, outcome))
            undo = self.run("Tk_RestoreLast")
            if not undo.startswith("ok:") or self.heights(self.toolkit) != before:
                undo_failures.append(("row", selection, start, undo))

        col_cases = [(sel, start) for sel in SELECTIONS + ["C:D"] for start in ["default", "2", "8.43", "12", "mixed"]]
        for selection, start in col_cases:
            cases += 1
            for ws in (self.legacy, self.toolkit):
                self.reset(ws)
                if start not in ("default", "mixed"):
                    ws.Range("A1:G1").ColumnWidth = float(start)
                elif start == "mixed":
                    for c in range(1, 8):
                        ws.Columns(c).ColumnWidth = 2 if c % 2 else 20
            before = self.widths(self.toolkit)
            self.legacy_run(selection, "Column_Width_Toggle")
            outcome = self.run("Tk_ColumnWidthRange", self.toolkit.Range(selection))
            if not outcome.startswith("ok:") or self.widths(self.legacy) != self.widths(self.toolkit):
                mismatches.append(("column", selection, start, outcome))
            undo = self.run("Tk_RestoreLast")
            if not undo.startswith("ok:") or self.widths(self.toolkit) != before:
                undo_failures.append(("column", selection, start, undo))

        for active_width in (2, 7.5, 8.43, 10, 12.5, 17.5, 30):
            for selection in ("B2:D6", "B2:B4,D6:E8"):
                cases += 1
                for ws in (self.legacy, self.toolkit):
                    self.reset(ws)
                    ws.Columns("B").ColumnWidth = active_width
                before = self.widths(self.toolkit)
                self.legacy_run(selection, "Best_Fit_Column_Width")   # active cell: B2
                outcome = self.run("Tk_BestFitRange", self.toolkit.Range(selection), self.toolkit.Range("B2"))
                if not outcome.startswith("ok:") or self.widths(self.legacy) != self.widths(self.toolkit):
                    mismatches.append(("best fit", selection, active_width, outcome))
                undo = self.run("Tk_RestoreLast")
                if not undo.startswith("ok:") or self.widths(self.toolkit) != before:
                    undo_failures.append(("best fit", selection, active_width, undo))
        self.record(f"sizes: row height, column width and best fit match the legacy macros ({cases} cases, incl. "
                    "whole rows/columns and multi-area)", not mismatches, mismatches=mismatches[:10])
        self.record("sizes: every change undone exactly (whole affected rows and columns)", not undo_failures,
                    failures=undo_failures[:10])

    # ----- refusals, conflicts, failures -----

    def refusals_and_conflicts(self):
        ws = self.toolkit
        self.reset(ws)
        ws.Rows(4).Hidden = True
        before = self.heights(ws)
        outcome = self.run("Tk_RowHeightRange", ws.Range("B2:B6"))
        self.record("hidden row in the selection: refused, row stays hidden, heights unchanged",
                    outcome.startswith("refused:") and ws.Rows(4).Hidden and self.heights(ws) == before, outcome=outcome)
        outcome = self.run("Tk_RowHeightRange", ws.Range("B2:B3,B5:B6"))
        self.record("visible rows around a hidden one: applied and undone",
                    outcome.startswith("ok:") and self.run("Tk_RestoreLast").startswith("ok:")
                    and self.heights(ws) == before and ws.Rows(4).Hidden, outcome=outcome)
        ws.Rows(4).Hidden = False
        ws.Columns("C").Hidden = True
        before = self.widths(ws)
        outcome = self.run("Tk_ColumnWidthRange", ws.Range("B2:D2"))
        self.record("hidden column in the selection: refused, widths unchanged",
                    outcome.startswith("refused:") and ws.Columns("C").Hidden and self.widths(ws) == before,
                    outcome=outcome)
        ws.Columns("C").Hidden = False

        self.reset(ws)
        ws.Range("C3").Interior.Pattern = 4000  # xlPatternLinearGradient
        before = self.fills(ws)
        outcome = self.run("Tk_FillRange", ws.Range("B2:D4"), "yellow")
        self.record("gradient fill in the selection: refused, nothing changed",
                    outcome.startswith("refused:") and self.fills(ws) == before, outcome=outcome)

        self.reset(ws)
        ws.Range("B9:D9").Merge()
        ws.Range("B9:D9").Interior.ThemeColor = 5
        before, borders_before = self.fills(ws), self.borders(ws)
        ok_fill = self.run("Tk_FillRange", ws.Range("A8:E10"), "green")
        undo_fill = self.run("Tk_RestoreLast")
        ok_outline = self.run("Tk_OutlineRange", ws.Range("A8:E10"))
        undo_outline = self.run("Tk_RestoreLast")
        self.record("merged cells in a mixed selection: fill and outline applied and undone exactly",
                    all(o.startswith("ok:") for o in (ok_fill, undo_fill, ok_outline, undo_outline))
                    and self.fills(ws) == before and self.borders(ws) == borders_before,
                    outcomes=[ok_fill, undo_fill, ok_outline, undo_outline])
        ws.Range("B9:D9").UnMerge()

        self.reset(ws)
        self.run("Tk_FillRange", ws.Range("B2:D4"), "yellow")
        ws.Range("C3").Interior.Color = rgb(0, 255, 0)
        after_edit = self.fills(ws)
        outcome = self.run("Tk_RestoreLast")
        self.record("conflict (fill edited): undo refused, nothing restored",
                    outcome.startswith("refused:") and self.fills(ws) == after_edit, outcome=outcome)
        self.run("Tk_OutlineRange", ws.Range("B2:D4"))
        ws.Range("C1").Borders(9).LineStyle = X_DASH      # shared with C2's top edge, which undo would restore
        after_edit = self.borders(ws)
        outcome = self.run("Tk_RestoreLast")
        self.record("conflict (neighbouring edge edited): undo refused, nothing restored",
                    outcome.startswith("refused:") and self.borders(ws) == after_edit, outcome=outcome)
        self.run("Tk_ColumnWidthRange", ws.Range("H1"))
        ws.Columns("H").ColumnWidth = 33
        outcome = self.run("Tk_RestoreLast")
        self.record("conflict (width edited): undo refused, width kept",
                    outcome.startswith("refused:") and ws.Columns("H").ColumnWidth == 33, outcome=outcome)
        self.run("Tk_FillRange", ws.Range("J1"), "grey")
        name = ws.Name
        ws.Name = "RenamedM3b"
        outcome = self.run("Tk_RestoreLast")
        self.record("identity: undo refused after the sheet is renamed", outcome.startswith("refused:"), outcome=outcome)
        ws.Name = name

    def protection(self):
        ws = self.toolkit
        self.reset(ws)
        cases = [
            # Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells,
            #         AllowFormattingColumns, AllowFormattingRows) - positional (keywords are ignored late-bound)
            ("fill", ("", True, True, True, False, False, True, True), ("Tk_FillRange", "yellow"), False),
            ("fill", ("", True, True, True, False, True, False, False), ("Tk_FillRange", "yellow"), True),
            ("row height", ("", True, True, True, False, True, True, False), ("Tk_RowHeightRange",), False),
            ("row height", ("", True, True, True, False, False, False, True), ("Tk_RowHeightRange",), True),
            ("column width", ("", True, True, True, False, True, False, True), ("Tk_ColumnWidthRange",), False),
            ("column width", ("", True, True, True, False, False, True, False), ("Tk_ColumnWidthRange",), True),
        ]
        for label, protect, (procedure, *args), allowed in cases:
            before = (self.fills(ws, "B2:C3"), self.heights(ws), self.widths(ws))
            ws.Protect(*protect)
            assert ws.ProtectContents, "fixture: sheet not protected"
            outcome = self.run(procedure, ws.Range("B2:C3"), *args)
            if allowed:
                good = outcome.startswith("ok:") and self.run("Tk_RestoreLast").startswith("ok:")
            else:
                good = outcome.startswith("refused:")
            ws.Unprotect()
            good = good and (self.fills(ws, "B2:C3"), self.heights(ws), self.widths(ws)) == before
            self.record(f"protected sheet, {label} {'allowed' if allowed else 'not allowed'}: "
                        f"{'applied and undone' if allowed else 'refused, nothing changed'}", good, outcome=outcome)

    def failures(self):
        ws = self.toolkit
        self.setup_borders(ws, "neighbours")
        before = self.borders(ws)
        self.run("Tk_TestSetFailAfter", 2, 0)
        outcome = self.run("Tk_OutlineRange", ws.Range("B2:D6"))
        self.run("Tk_TestSetFailAfter", 0, 0)
        self.record("forced failure after two outline edges: rolled back, every edge identical",
                    outcome.startswith("rolled back:") and self.borders(ws) == before, outcome=outcome)

        self.reset(ws)
        before = self.heights(ws)
        self.run("Tk_TestSetFailAfter", 1, 0)
        outcome = self.run("Tk_RowHeightRange", ws.Range("B2:B3,B6:B8"))
        self.run("Tk_TestSetFailAfter", 0, 0)
        self.record("forced failure after the first area (row height): rolled back, heights identical",
                    outcome.startswith("rolled back:") and self.heights(ws) == before, outcome=outcome)

        self.setup_fill(ws, "mixed", rgb(255, 255, 210))
        self.run("Tk_FillRange", ws.Range("A1:F10"), "yellow")
        self.run("Tk_TestSetFailAfter", 0, 2)
        outcome = self.run("Tk_RestoreLast")
        self.run("Tk_TestSetFailAfter", 0, 0)
        self.record("forced restore failure: reported as incomplete, record unusable",
                    outcome.startswith("incomplete:") and self.run("Tk_RestoreLast").startswith("refused:"),
                    outcome=outcome)

    def too_many_blocks(self):
        ws = self.wb.Worksheets.Add()
        ws.Range("A1").Interior.ThemeColor = 5
        ws.Range("A1:A2").Copy(ws.Range("A3:A60000"))  # alternating theme fill / no fill; Destination copy
        sample = [fill_state(ws.Range(a)) for a in ("A1", "A2", "A59999")]
        outcome = self.run("Tk_FillRange", ws.Range("A1:A60000"), "yellow")
        self.record("too many distinct fill blocks: refused, nothing changed",
                    outcome.startswith("refused:") and sample == [fill_state(ws.Range(a)) for a in ("A1", "A2", "A59999")],
                    outcome=outcome)

    # ----- performance -----

    def performance(self):
        ws = self.wb.Worksheets.Add()
        yellow = rgb(255, 255, 210)

        def alternating(address, first):
            ws.Range(first).Interior.ThemeColor = 5
            ws.Range(f"{first}:{first[0]}{int(first[1:]) + 1}").Copy(ws.Range(address))

        def column_exceptions():
            ws.Columns("K").Interior.Pattern = X_NONE
            for row in range(1, 1_000_001, 100_000):
                ws.Range(f"K{row}").Interior.ThemeColor = 6

        cases = [
            ("fill on, uniform 10,000 (no fill)", "Tk_FillRange", ("A1:A10000", "yellow"), lambda: None),
            ("fill off, uniform 10,000 RGB yellow", "Tk_FillRange", ("B1:B10000", "yellow"),
             lambda: ws.Range("B1:B10000").Interior.__setattr__("Color", yellow)),
            ("fill on, uniform 10,000 theme colour", "Tk_FillRange", ("C1:C10000", "yellow"),
             lambda: ws.Range("C1:C10000").Interior.__setattr__("ThemeColor", 5)),
            ("fill on, 10,000 alternating rows", "Tk_FillRange", ("E1:E10000", "yellow"),
             lambda: alternating("E1:E10000", "E1")),
            ("fill on, whole column uniform", "Tk_FillRange", ("I:I", "yellow"), lambda: None),
            ("fill on, whole column, 10 exceptions", "Tk_FillRange", ("K:K", "yellow"), column_exceptions),
            ("outline on, 10,000 cells (100 x 100)", "Tk_OutlineRange", ("M1:DH100",), lambda: None),
            ("outline off, 10,000 outlined cells (100 x 100)", "Tk_OutlineRange", ("M1:DH100",),
             lambda: self.run("Tk_OutlineRange", ws.Range("M1:DH100"))),
            ("outline on, 40,000 cells (100 x 400)", "Tk_OutlineRange", ("M201:DH600",), lambda: None),
            ("row height, whole column (every row)", "Tk_RowHeightRange", ("DM:DM",), lambda: None),
            ("column width, whole row (every column)", "Tk_ColumnWidthRange", ("5:5",), lambda: None),
        ]
        before = border_state(ws.Range("DK500"))
        outcome = self.run("Tk_OutlineRange", ws.Range("DK:DK"))
        self.record("outline on a whole column (over the 50,000-cell limit): refused, nothing changed",
                    outcome.startswith("refused:") and border_state(ws.Range("DK500")) == before, outcome=outcome)
        for label, procedure, (address, *args), prepare in cases:
            prepare()
            started = time.perf_counter()
            outcome = self.run(procedure, ws.Range(address), *args)
            wall = time.perf_counter() - started
            timings = self.run("Tk_LastStyleTimings") if outcome.startswith("ok:") else "(not applied)"
            started = time.perf_counter()
            undo = self.run("Tk_RestoreLast")
            undo_wall = time.perf_counter() - started
            self.record(f"performance case '{label}': applied and undone", outcome.startswith("ok:") and undo.startswith("ok:"),
                        outcome=outcome, undo=undo)
            entry = {"case": label, "outcome": outcome, "apply_wall_ms": round(wall * 1000, 1), "vba": timings,
                     "undo": undo, "undo_wall_ms": round(undo_wall * 1000, 1),
                     "undo_vba_ms": round(self.run("Tk_LastRestoreSeconds") * 1000, 1)}
            self.timings.append(entry)
            print(entry)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlam", type=Path, default=DEFAULT_XLAM)
    parser.add_argument("--perf", action="store_true")
    parser.add_argument("--out", type=Path, default=HERE / "results" / f"m3b_live_{time.strftime('%Y%m%d_%H%M%S')}.json")
    args = parser.parse_args()

    personal_before = sha256(PERSONAL)
    check = Check(args.xlam)
    try:
        print("add-in version:", check.run("Tk_Version"))
        check.diagnostics()
        check.fills_equivalence()
        check.outline_equivalence()
        check.sizes_equivalence()
        check.refusals_and_conflicts()
        check.protection()
        check.failures()
        check.too_many_blocks()
        if args.perf:
            check.performance()
    finally:
        fired = check.watchdog.fired
        check.close()
    check.record("PERSONAL.XLSB unchanged", sha256(PERSONAL) == personal_before)
    if fired:
        check.record("run completed without stalling", False, note="watchdog terminated the test Excel")
    passed = sum(r["passed"] for r in check.results)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps({"results": check.results, "timings": check.timings}, indent=2, default=str),
                        encoding="utf-8")
    print(f"\n{passed}/{len(check.results)} checks passed -> {args.out}")
    return 0 if passed == len(check.results) else 1


if __name__ == "__main__":
    sys.exit(main())
