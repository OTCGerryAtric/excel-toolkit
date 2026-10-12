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
from m3c_cases import REWRITE_CASES

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


class Check(M3aCheck):
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
        books = []
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
            legacy.Name = "Sheet1"
        finally:
            for book in books:
                book.Close(SaveChanges=False)

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
    args.out.write_text(json.dumps({"results": check.results, "timings": check.timings}, indent=2, default=str),
                        encoding="utf-8")
    print(f"\n{passed}/{len(check.results)} checks passed -> {args.out}")
    return 0 if passed == len(check.results) else 1


if __name__ == "__main__":
    sys.exit(main())
