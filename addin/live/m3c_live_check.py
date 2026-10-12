"""M3c scripted live check - Clear Sheet References and Find Constants. DO NOT RUN until authorised.

Same boundaries as m2/m3a/m3b: a SEPARATE disposable Excel process (DispatchEx: no XLSTART, no
PERSONAL.XLSB, no installed add-ins), workbooks it creates and closes without saving, the development
add-in opened with Workbooks.Open, nothing installed. The legacy macros come from baseline/vba.

  <scratch venv>/python addin/live/m3c_live_check.py [--xlam addin/build/ExcelToolkit.xlam] [--perf]
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

from m2_live_check import DEFAULT_XLAM, PERSONAL, sha256
from m3a_live_check import Check as M3aCheck
from m3c_cases import AMBIGUOUS_CASES, REWRITE_CASES
from m3c_rewrite_model import rewrite_sheet_prefix

HERE = Path(__file__).resolve().parent
OWN = "P&L"
FIXTURE = "A1:F12"

# (cell, formula or constant) on the sheet P&L; '=' entries are formulas.
CELLS = {
    "B2": "='P&L'!A1+1",
    "B3": "=SUM('P&L'!A1:'P&L'!A3)",
    "B4": "='P&L 2'!A1",
    "B5": "=\"'P&L'!A1\"&'P&L'!A1",
    "B6": "=Other!A1+'P&L'!A2",
    "B7": "=A1*2",
    "C2": "'P&L'!A1 is a note",           # text: never changed
    "C3": "Text mentioning P&L!A1",
    "C4": 42,
    "C5": "=SEQUENCE(3,1,'P&L'!A1)",       # a spill anchor; C6:C7 are spilled values
    "D2": "=INDIRECT(\"'P&L'!A1\")",
}


# Formulas entered (Formula2) on each sheet of the meaning workbook, sheets in order Sheet1, A, AB, Data,
# Sheet2, Other, each with numbers in A1:C3. Excel stores some of them differently (A:Sheet1!B2 as the 3D
# range Sheet1:A!B2); the check works on what Excel stored.
MEANING_FORMULAS = {
    "Sheet1": [
        "=Sheet1!B2+1",
        "=SUM(Sheet1!A1:Sheet1!B2)",
        "=SUM(Sheet1!C3:Sheet1!A1)",
        "=SUM(A1:Sheet1!B2)",
        "=SUM($A$1:Sheet1!$B$2)",
        "=SUM(INDEX(A:A,1):Sheet1!B2)",
        "=SUM(Sheet1!A1:INDEX(A:A,3))",
        "=SUM(Sheet2!A1:Sheet1!B2)",            # first half on another sheet: #VALUE!, kept
        "=SUM(Sheet2!A1:B2)+Sheet1!A1",
        "=SUM(A:Sheet1!B2)",                     # stored as the 3D range Sheet1:A!B2
        "=SUM(Data:Sheet1!B2)",
        "=A!B2+AB!B2+Sheet1!B2",
        "=SUM(Sheet1!A:A)+SUM(Sheet1!1:3)",
        "=INDIRECT(\"Sheet1!B2\")+Sheet1!C2",
        "=SUMPRODUCT((Sheet1!A1:A3>1)*Sheet1!B1:B3)",
        "=Sheet1!B2&\" Sheet1!B2\"",
        "=IFERROR(Sheet1!Z99/0,Sheet1!A1)",
        "=SUM(AB:Sheet1!B2)+Sheet1!C1",          # 3D range plus an own reference: left unchanged
        "=SUM(AB:'Sheet1'!B2)+Sheet1!C1",        # stored as AB:Sheet1!B2 (not 3D): left unchanged
    ],
    "A": [
        "=A!B2+1",
        "=SUM(A!A1:A!B2)",
        "=SUM(a!B2:A!C3)+AB!A1",
        "=SUM(Sheet1:A!B2)",                     # 3D range ending at this sheet
        "=SUM(A:AB!B2)+A!C1",                    # 3D range starting at this sheet: left unchanged
    ],
}


class Check(M3aCheck):
    def __init__(self, xlam):
        super().__init__(xlam)
        self.meaning_log: list[dict] = []

    def build_book(self):
        wb = self.excel.Workbooks.Add()
        while wb.Worksheets.Count < 4:
            wb.Worksheets.Add()
        for ws, name in zip(wb.Worksheets, (OWN, "P&L 2", "Other", "O'Brien")):
            ws.Name = name
        own = wb.Worksheets(OWN)
        own.Range("A1:A3").Value2 = ((1,), (2,), (3,))
        for name in ("P&L 2", "Other"):
            wb.Worksheets(name).Range("A1:A2").Value2 = ((10,), (20,))
        for address, content in CELLS.items():
            if isinstance(content, str) and content.startswith("="):
                own.Range(address).Formula2 = content
            else:
                own.Range(address).Value2 = content
        return wb, own

    @staticmethod
    def formulas(ws, address=FIXTURE):
        return {c.Address: (c.Formula2, c.HasFormula) for c in ws.Range(address).Cells}

    @staticmethod
    def values(ws, address=FIXTURE):
        return {c.Address: c.Value2 for c in ws.Range(address).Cells}

    # ----- scenarios -----

    def rewrite_rules(self):
        wrong = []
        for sheet, before, after in REWRITE_CASES:
            got = self.run("Tk_RemoveSheetPrefix", before, sheet)
            if got != after:
                wrong.append((sheet, before, after, got))
        self.record(f"rewrite rules: VBA Tk_RemoveSheetPrefix gives the expected result for all {len(REWRITE_CASES)} cases",
                    not wrong, wrong=wrong[:10])
        wrong = [(sheet, formula, got) for sheet, formula in AMBIGUOUS_CASES
                 if (got := self.run("Tk_RemoveSheetPrefix", formula, sheet)) != formula]
        self.record(f"rewrite rules: all {len(AMBIGUOUS_CASES)} formulas Excel could read two ways are left unchanged",
                    not wrong, wrong=wrong)

    def clear_references(self):
        wb, own = self.build_book()
        try:
            own.Activate()
            before_f, before_v = self.formulas(own), self.values(own)
            outcome = self.run("Tk_ClearSheetReferences", own.Range("B2:D7"))
            after = self.formulas(own)
            expected = {"$B$2": "=A1+1", "$B$3": "=SUM(A1:A3)", "$B$4": "='P&L 2'!A1",
                        "$B$5": "=\"'P&L'!A1\"&A1", "$B$6": "=Other!A1+A2", "$B$7": "=A1*2",
                        "$C$5": "=SEQUENCE(3,1,A1)", "$D$2": "=INDIRECT(\"'P&L'!A1\")"}
            wrong = {a: after[a][0] for a, f in expected.items() if after[a][0] != f}
            texts_same = all(after[a] == before_f[a] for a in ("$C$2", "$C$3", "$C$4"))
            self.record("selection: own-sheet prefixes removed; other sheets, similar names, string literals, "
                        "INDIRECT text and text cells untouched",
                        outcome.startswith("ok:") and not wrong and texts_same, outcome=outcome, wrong=wrong)
            self.record("values unchanged by the rewrite", self.values(own) == before_v)
            undo = self.run("Tk_RestoreLast")
            self.record("undo restores every formula exactly", undo.startswith("ok:") and self.formulas(own) == before_f,
                        undo=undo)

            outcome = self.run("Tk_ClearSheetReferences", own.Range("A1"))          # single cell: whole sheet
            self.record("single cell selected: works on the whole sheet, then undone",
                        outcome.startswith("ok:") and "the sheet" in outcome and self.formulas(own)["$B$2"][0] == "=A1+1"
                        and self.run("Tk_RestoreLast").startswith("ok:") and self.formulas(own) == before_f, outcome=outcome)

            outcome = self.run("Tk_ClearSheetReferences", own.Range("B4,B7,C2:C4"))  # mixed, nothing to do
            self.record("mixed selection with nothing to rewrite: 'nothing', no change",
                        outcome.startswith("nothing:") and self.formulas(own) == before_f, outcome=outcome)

            outcome = self.run("Tk_ClearSheetReferences", own.Range("B2,B5:B6,C2:C3"))  # multi-area
            got = self.formulas(own)
            self.record("multi-area selection: only the selected formulas rewritten, then undone",
                        outcome.startswith("ok:") and got["$B$2"][0] == "=A1+1" and got["$B$3"] == before_f["$B$3"]
                        and self.run("Tk_RestoreLast").startswith("ok:") and self.formulas(own) == before_f, outcome=outcome)

            # refusals, conflicts, failures
            own.Range("E2:E3").FormulaArray = "='P&L'!A1:A2*2"
            snapshot = self.formulas(own)
            outcome = self.run("Tk_ClearSheetReferences", own.Range("B2:E3"))
            self.record("legacy array formula (Ctrl+Shift+Enter) referring to the sheet: refused, nothing changed",
                        outcome.startswith("refused:") and self.formulas(own) == snapshot, outcome=outcome)
            own.Range("E2:E3").ClearContents()

            before_f = self.formulas(own)
            self.run("Tk_ClearSheetReferences", own.Range("B2:B6"))
            own.Range("B3").Formula2 = "=99"
            edited = self.formulas(own)
            outcome = self.run("Tk_RestoreLast")
            self.record("conflict: undo refused after a rewritten formula is edited, nothing restored",
                        outcome.startswith("refused:") and self.formulas(own) == edited, outcome=outcome)
            own.Range("B3").Formula2 = CELLS["B3"]

            self.run("Tk_ClearSheetReferences", own.Range("B2:B6"))
            own.Name = "Renamed"
            outcome = self.run("Tk_RestoreLast")
            self.record("identity: undo refused after the sheet is renamed", outcome.startswith("refused:"), outcome=outcome)
            own.Name = OWN
            for address in ("B2", "B3", "B5", "B6"):
                own.Range(address).Formula2 = CELLS[address]

            own.Protect()
            before_f = self.formulas(own)
            outcome = self.run("Tk_ClearSheetReferences", own.Range("B2:B6"))
            own.Unprotect()
            self.record("protected sheet: refused, nothing changed",
                        outcome.startswith("refused:") and self.formulas(own) == before_f, outcome=outcome)

            self.run("Tk_TestSetFailAfter", 2, 0)
            outcome = self.run("Tk_ClearSheetReferences", own.Range("B2:B6"))
            self.run("Tk_TestSetFailAfter", 0, 0)
            self.record("forced failure after two formulas: rolled back, formulas identical",
                        outcome.startswith("rolled back:") and self.formulas(own) == before_f, outcome=outcome)

            self.run("Tk_ClearSheetReferences", own.Range("B2:B6"))
            self.run("Tk_TestSetFailAfter", 0, 1)
            outcome = self.run("Tk_RestoreLast")
            self.run("Tk_TestSetFailAfter", 0, 0)
            self.record("forced restore failure: reported as incomplete, record unusable",
                        outcome.startswith("incomplete:") and self.run("Tk_RestoreLast").startswith("refused:"),
                        outcome=outcome)
        finally:
            wb.Close(SaveChanges=False)

    def legacy_comparison(self):
        """Plain own-sheet references on a sheet named Data: same result as the legacy macro. The legacy
        macro's text replace also changes a text cell and a string literal, which the toolkit leaves."""
        contents = {"B2": "=Data!A1+1", "B3": "=SUM(Data!A1:Data!A3)", "B4": '="Data!A1"', "C2": "Data!A1 is a note"}
        books, original = [], self.legacy.Name
        try:
            for ws in (self.legacy, self.excel.Workbooks.Add().Worksheets(1)):
                if ws is not self.legacy:
                    books.append(ws.Parent)
                ws.Cells.Clear()
                ws.Name = "Data"
                ws.Range("A1:A3").Value2 = ((1,), (2,), (3,))
                for address, content in contents.items():
                    if content.startswith("="):
                        ws.Range(address).Formula2 = content
                    else:
                        ws.Range(address).Value2 = content
            legacy, toolkit = self.legacy, books[0].Worksheets(1)
            legacy.Parent.Activate()
            legacy.Activate()
            legacy.Range("A1").Select()
            self.run_legacy("Clear_Sheet_Formula_References")
            outcome = self.run("Tk_ClearSheetReferences", toolkit.Range("A1"))
            got = {a: (legacy.Range(a).Formula2, toolkit.Range(a).Formula2) for a in contents}
            self.record("plain own-sheet references: same result as the legacy macro; the legacy macro also changed a "
                        "string literal and a text cell, the toolkit did not",
                        outcome.startswith("ok:") and got["B2"][0] == got["B2"][1] == "=A1+1"
                        and got["B3"][0] == got["B3"][1] == "=SUM(A1:A3)"
                        and got["B4"][1] == contents["B4"] and got["C2"][1] == contents["C2"]
                        and got["B4"][0] != contents["B4"] and got["C2"][0] != contents["C2"], got=got)
            self.run("Tk_RestoreLast")
            legacy.Name = original
        finally:
            for book in books:
                book.Close(SaveChanges=False)

    def meaning(self):
        """Calculated results are unchanged by the rewrite, including sheets named like columns (A, AB), 3D
        ranges, a range whose first half is on another sheet, and a formula whose stored text Excel would
        read back differently (AB:'Sheet1'!B2 is stored as AB:Sheet1!B2, which re-entered is a 3D range).
        Formulas Excel could read two ways are left exactly as they are and reported; undo is exact."""
        wb = self.excel.Workbooks.Add()
        try:
            names = ("Sheet1", "A", "AB", "Data", "Sheet2", "Other")
            while wb.Worksheets.Count < len(names):
                wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            for index in range(1, len(names) + 1):
                wb.Worksheets(index).Name = f"tmp{index}"
            for index, name in enumerate(names, start=1):
                ws = wb.Worksheets(index)
                ws.Name = name
                ws.Range("A1:C3").Value2 = tuple(tuple(index * 1000 + r * 10 + c for c in range(3)) for r in range(3))
            for own_name, entries in MEANING_FORMULAS.items():
                own = wb.Worksheets(own_name)
                own.Activate()
                cells = [own.Range(f"E{row}") for row in range(10, 10 + len(entries))]
                for cell, formula in zip(cells, entries):
                    cell.Formula2 = formula
                self.excel.Calculate()
                before = [(c.Formula2, c.Value2) for c in cells]
                expected = [rewrite_sheet_prefix(f, own_name) for f, _ in before]
                changes = sum(1 for (f, _), (new, _) in zip(before, expected) if new != f)
                skipped = [c.GetAddress(False, False) for c, (_, amb) in zip(cells, expected) if amb]
                outcome = self.run("Tk_ClearSheetReferences", own.Range("A1"))
                self.excel.Calculate()
                after = [(c.Formula2, c.Value2) for c in cells]
                value_changes = {c.GetAddress(False, False): (b, a) for c, (_, b), (_, a) in zip(cells, before, after) if a != b}
                text_wrong = {c.GetAddress(False, False): (b[0], e[0], a[0]) for c, b, e, a in zip(cells, before, expected, after)
                              if (a[0] != b[0]) != (e[0] != b[0])          # changed exactly where expected
                              or self.run("Tk_RemoveSheetPrefix", a[0], own_name) != a[0]}
                ambiguous_kept = all(a[0] == b[0] for c, b, a, (_, amb) in zip(cells, before, after, expected) if amb)
                expected_kind = "warning:" if skipped else "ok:"
                self.record(f"meaning on '{own_name}': {changes} formula(s) rewritten, every calculated result unchanged",
                            outcome.startswith(expected_kind) and not value_changes and not text_wrong, outcome=outcome,
                            value_changes=value_changes, text_wrong=text_wrong)
                self.record(f"meaning on '{own_name}': {len(skipped)} formula(s) Excel could read two ways left exactly "
                            "as they were and reported",
                            ambiguous_kept and len(skipped) > 0 and f"{len(skipped)} formula(s) naming this sheet were left "
                            f"unchanged (the first is in {skipped[0]})" in outcome, skipped=skipped, outcome=outcome)
                undo = self.run("Tk_RestoreLast")
                self.excel.Calculate()
                self.record(f"meaning on '{own_name}': undo restores every formula and result exactly",
                            undo.startswith("ok:") and [(c.Formula2, c.Value2) for c in cells] == before, undo=undo)
                self.meaning_log.append({"sheet": own_name, "outcome": outcome,
                                         "cells": [{"before": b[0], "after": a[0], "value": b[1]} for b, a in zip(before, after)]})
            outcome = self.run("Tk_ClearSheetReferences", wb.Worksheets("Sheet1").Range("E27,E28"))
            self.record("only ambiguous formulas selected: 'nothing', reported, no change",
                        outcome.startswith("nothing:") and "read two ways" in outcome, outcome=outcome)
        finally:
            wb.Close(SaveChanges=False)

    def find_constants(self):
        ws = self.legacy
        ws.Cells.Clear()
        ws.Range("A1").Value2 = 1
        ws.Range("A2").Formula2 = "=A1+1"
        ws.Range("B1").Value2 = "text"
        ws.Range("B2").Value2 = 2.5
        ws.Range("C3").Value2 = 7
        results = []
        for selection in ("A1:C3", "A1", "A1:A2,C3"):
            ws.Activate()
            ws.Range(selection).Select()
            self.run_legacy("Find_Constants")
            legacy = self.excel.Selection.Address
            ws.Range(selection).Select()
            outcome = self.run("Tk_FindConstants", ws.Range(selection))
            results.append((selection, legacy, self.excel.Selection.Address, outcome))
        self.record("find constants: selects the same cells as the legacy macro (block, single cell = whole sheet, "
                    "two areas)", all(r[1] == r[2] and r[3].startswith("ok:") for r in results), results=results)
        ws.Cells.Clear()
        ws.Range("A1").Formula2 = "=1+1"
        outcome = self.run("Tk_FindConstants", ws.Range("A1:B2"))
        self.record("find constants with no typed-in numbers: 'nothing'", outcome.startswith("nothing:"), outcome=outcome)
        self.run("Tk_FormatRange", ws.Range("D1"), "number")
        self.run("Tk_FindConstants", ws.Range("A1:B2"))
        self.record("find constants leaves the existing undo record alone", self.run("Tk_HasRecord") is True)
        self.run("Tk_RestoreLast")

    def performance(self):
        wb = self.excel.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = OWN
            ws.Range("A1:A10000").Value2 = 1
            ws.Range("B1:B10000").Formula2 = "='P&L'!A1*2"
            ws.Range("C1:C10000").Formula2 = "=A1*3"
            for label, address in (("10,000 formulas to rewrite", "B1:B10000"),
                                   ("10,000 formulas, nothing to rewrite", "C1:C10000")):
                started = time.perf_counter()
                outcome = self.run("Tk_ClearSheetReferences", ws.Range(address))
                wall = time.perf_counter() - started
                timings = self.run("Tk_LastFormulaTimings") if outcome.startswith("ok:") else ""
                undo, undo_wall = "", 0.0
                if outcome.startswith("ok:"):
                    started = time.perf_counter()
                    undo = self.run("Tk_RestoreLast")
                    undo_wall = time.perf_counter() - started
                self.record(f"performance case '{label}'", outcome.startswith(("ok:", "nothing:"))
                            and (not undo or undo.startswith("ok:")), outcome=outcome, undo=undo)
                entry = {"case": label, "outcome": outcome, "apply_wall_ms": round(wall * 1000, 1), "vba": timings,
                         "undo": undo, "undo_wall_ms": round(undo_wall * 1000, 1)}
                self.timings.append(entry)
                print(entry)
        finally:
            wb.Close(SaveChanges=False)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlam", type=Path, default=DEFAULT_XLAM)
    parser.add_argument("--perf", action="store_true")
    parser.add_argument("--out", type=Path, default=HERE / "results" / f"m3c_live_{time.strftime('%Y%m%d_%H%M%S')}.json")
    args = parser.parse_args()

    personal_before = sha256(PERSONAL)
    check = Check(args.xlam)
    try:
        print("add-in version:", check.run("Tk_Version"))
        check.rewrite_rules()
        check.clear_references()
        check.legacy_comparison()
        check.meaning()
        check.find_constants()
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
    args.out.write_text(json.dumps({"results": check.results, "timings": check.timings, "meaning": check.meaning_log},
                                   indent=2, default=str),
                        encoding="utf-8")
    print(f"\n{passed}/{len(check.results)} checks passed -> {args.out}")
    return 0 if passed == len(check.results) else 1


if __name__ == "__main__":
    sys.exit(main())
