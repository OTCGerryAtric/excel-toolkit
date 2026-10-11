"""M3a scripted live check - number-format commands vs the legacy macros. DO NOT RUN until authorised.

Same boundaries as m2_live_check.py: a SEPARATE disposable Excel process (DispatchEx: no XLSTART, no
PERSONAL.XLSB, no installed add-ins), workbooks it creates and closes without saving, the development
add-in opened with Workbooks.Open, nothing installed. The legacy macros are imported from the baseline
source (baseline/vba) into the disposable workbook - never from the real PERSONAL.XLSB.

Not covered here (manual checklist, docs/validation/m3a-number-formats.md): ribbon clicks, real Ctrl+Z,
and the shortcut-precedence test, which needs the owner's Excel with PERSONAL.XLSB loaded.

  <scratch venv>/python addin/live/m3a_live_check.py [--xlam addin/build/ExcelToolkit.xlam] [--perf]
"""

from __future__ import annotations

import argparse
import json
import random
import sys
import tempfile
import time
from pathlib import Path

import win32com.client
import win32process

from m2_live_check import DEFAULT_XLAM, PERSONAL, Watchdog, _terminate, sha256

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
LEGACY_DUMP = ROOT / "baseline" / "vba" / "PERSONAL_XLSB_olevba.txt"
AREA = "A1:F12"
POUND = "£"

CYCLES = {
    "number": ("Number_Format_Toggle", ["#,##0_-;[Red](#,##0)_-;-_-", "#,##0.0_-;[Red](#,##0.0)_-;-_-",
                                        "#,##0.00_-;[Red](#,##0.00)_-;-_-"]),
    "currency": ("Currency_Format_Toggle", [f"{POUND}#,##0_-;[Red]({POUND}#,##0)_-;-_-",
                                            f"{POUND}#,##0.0_-;[Red]({POUND}#,##0.0)_-;-_-",
                                            f"{POUND}#,##0.00_-;[Red]({POUND}#,##0.00)_-;-_-"]),
    "percent": ("Percentage_Format_Toggle", ["0%_-;[Red](0%)_-;-_-", "0.0%_-;[Red](0.0%)_-;-_-",
                                             "0.00%_-;[Red](0.00%)_-;-_-", "0.000%_-;[Red](0.000%)_-;-_-"]),
    "date": ("Month_Date_Format", ["dd-mmm-yy", "mmm-yy"]),
}
SELECTIONS = ["B2", "B2:D10", "B2:B5,D7:E9"]
MIXED = ["0.00", "General", "0%", "dd-mmm-yy", "@"]


def legacy_module_file() -> Path:
    """Module1 from the baseline olevba dump, as an importable .bas (shortcut attributes removed)."""
    text = LEGACY_DUMP.read_text(encoding="utf-8")
    body = text.split("VBA MACRO Module1.bas", 1)[1].split("\n- - - ", 1)[1].split("\n", 1)[1]
    body = body.split("\n-------", 1)[0]
    lines = [line for line in body.splitlines() if "VB_ProcData.VB_Invoke_Func" not in line]
    source = 'Attribute VB_Name = "LegacyModule1"\r\n' + "\r\n".join(lines) + "\r\n"
    path = Path(tempfile.gettempdir()) / "tk_m3a_LegacyModule1.bas"
    path.write_bytes(source.encode("cp1252"))  # VBA imports in the ANSI code page (the pound sign)
    return path


class Check:
    def __init__(self, xlam: Path):
        self.excel = win32com.client.DispatchEx("Excel.Application")
        _, self.pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
        self.watchdog = Watchdog(self.pid)
        self.results: list[dict] = []
        self.timings: list[dict] = []
        try:
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.addin = self.excel.Workbooks.Open(str(xlam))
            self.name = self.addin.Name
            self.run("Tk_SetQuiet", True)
            self.wb = self.excel.Workbooks.Add()
            self.wb.VBProject.VBComponents.Import(str(legacy_module_file()))
            while self.wb.Worksheets.Count < 2:
                self.wb.Worksheets.Add()
            self.legacy, self.toolkit = self.wb.Worksheets(1), self.wb.Worksheets(2)
        except Exception:
            self.watchdog.stop()
            _terminate(self.pid)
            raise

    def run(self, procedure, *args):
        self.watchdog.tick()
        value = self.excel.Run(f"'{self.name}'!{procedure}", *args)
        self.watchdog.tick()
        return value

    def run_legacy(self, macro):
        self.watchdog.tick()
        self.excel.Run(f"'{self.wb.Name}'!{macro}")
        self.watchdog.tick()

    def record(self, name, passed, **details):
        self.results.append({"check": name, "passed": bool(passed), **details})
        print(("PASS " if passed else "FAIL ") + name + ("" if passed else f"  {details}"))

    @staticmethod
    def formats(ws, address=AREA):
        return {c.Address: c.NumberFormat for c in ws.Range(address).Cells}

    @staticmethod
    def setup(ws, start, address=AREA):
        ws.Range(address).Value2 = 1234.5
        if start == "mixed":
            for i, cell in enumerate(ws.Range(address).Cells):
                cell.NumberFormat = MIXED[i % len(MIXED)]
        else:
            ws.Range(address).NumberFormat = start

    def toolkit_press(self, selection, cycle):
        return self.run("Tk_FormatRange", self.toolkit.Range(selection), cycle)

    def legacy_press(self, selection, macro):
        self.legacy.Activate()
        self.legacy.Range(selection).Select()
        self.run_legacy(macro)

    # ----- scenarios -----

    def equivalence(self):
        mismatches, undo_failures, cases = [], [], 0
        for cycle, (macro, formats) in CYCLES.items():
            for selection in SELECTIONS:
                for start in formats + ["General", "0.00", "mixed"]:
                    cases += 1
                    for ws in (self.legacy, self.toolkit):
                        self.setup(ws, start)
                    before = self.formats(self.toolkit)
                    self.legacy_press(selection, macro)
                    outcome = self.toolkit_press(selection, cycle)
                    if not outcome.startswith("ok:") or self.formats(self.legacy) != self.formats(self.toolkit):
                        mismatches.append((cycle, selection, start, outcome))
                    if not self.run("Tk_RestoreLast").startswith("ok:") or self.formats(self.toolkit) != before:
                        undo_failures.append((cycle, selection, start))
                # five presses in a row from General
                for ws in (self.legacy, self.toolkit):
                    self.setup(ws, "General")
                for press in range(5):
                    self.legacy_press(selection, macro)
                    self.toolkit_press(selection, cycle)
                    if self.formats(self.legacy) != self.formats(self.toolkit):
                        mismatches.append((cycle, selection, f"press {press + 1}", ""))
        self.record(f"equivalence: toolkit matches the legacy macro cell by cell ({cases} start states + 5-press cycles)",
                    not mismatches, mismatches=mismatches[:10])
        self.record("equivalence: every toolkit change undone exactly (incl. mixed formats)", not undo_failures,
                    failures=undo_failures[:10])
        values = {c.Value2 for c in self.toolkit.Range(AREA).Cells}
        self.record("values untouched by format commands", values == {1234.5}, values=sorted(map(str, values))[:5])

    def undo_conflict_identity(self):
        ws = self.toolkit
        rng = "A1:E100"
        ws.Range(rng).Value2 = 1
        rnd = random.Random(7)
        for cell in ws.Range(rng).Cells:
            cell.NumberFormat = rnd.choice(MIXED + ["#,##0", "0.000"])
        before = self.formats(ws, rng)
        outcome = self.run("Tk_FormatRange", ws.Range(rng), "number")
        self.record("mixed 500 cells: applied", outcome.startswith("ok:"), outcome=outcome)
        self.record("mixed 500 cells: undo restores every cell's format exactly",
                    self.run("Tk_RestoreLast").startswith("ok:") and self.formats(ws, rng) == before)

        self.run("Tk_FormatRange", ws.Range(rng), "number")
        ws.Range("C50").NumberFormat = "0.0000"
        after_edit = self.formats(ws, rng)
        outcome = self.run("Tk_RestoreLast")
        self.record("conflict: undo refused after a format edit, nothing restored",
                    outcome.startswith("refused:") and self.formats(ws, rng) == after_edit, outcome=outcome)

        self.run("Tk_FormatRange", ws.Range("H1"), "date")
        name = ws.Name
        ws.Name = "RenamedM3a"
        outcome = self.run("Tk_RestoreLast")
        self.record("identity: undo refused after the sheet is renamed", outcome.startswith("refused:"), outcome=outcome)
        ws.Name = name

    def protection_and_failures(self):
        ws = self.toolkit
        ws.Range("A1:B2").NumberFormat = "General"
        ws.Protect()
        before = self.formats(ws, "A1:B2")
        outcome = self.run("Tk_FormatRange", ws.Range("A1:B2"), "number")
        self.record("protected (formatting not allowed): refused, nothing changed",
                    outcome.startswith("refused:") and self.formats(ws, "A1:B2") == before, outcome=outcome)
        ws.Unprotect()
        ws.Protect(AllowFormattingCells=True)
        outcome = self.run("Tk_FormatRange", ws.Range("A1:B2"), "number")
        self.record("protected (formatting allowed): applied and undone",
                    outcome.startswith("ok:") and self.run("Tk_RestoreLast").startswith("ok:")
                    and self.formats(ws, "A1:B2") == before, outcome=outcome)
        ws.Unprotect()

        ws.Range("A1:A3").NumberFormat = "0.00"
        ws.Range("C1:C3").NumberFormat = "General"
        before = self.formats(ws, "A1:C3")
        self.run("Tk_TestSetFailAfter", 1, 0)
        outcome = self.run("Tk_FormatRange", ws.Range("A1:A3,C1:C3"), "percent")
        self.run("Tk_TestSetFailAfter", 0, 0)
        self.record("forced failure after the first area: rolled back, formats identical",
                    outcome.startswith("rolled back:") and self.formats(ws, "A1:C3") == before, outcome=outcome)

        self.setup(ws, "mixed", "A1:C6")
        self.run("Tk_FormatRange", ws.Range("A1:C6"), "percent")
        self.run("Tk_TestSetFailAfter", 0, 2)
        outcome = self.run("Tk_RestoreLast")
        self.run("Tk_TestSetFailAfter", 0, 0)
        self.record("forced restore failure: reported as incomplete, record unusable",
                    outcome.startswith("incomplete:") and self.run("Tk_RestoreLast").startswith("refused:"), outcome=outcome)

    def too_many_blocks(self):
        ws = self.wb.Worksheets.Add()
        ws.Range("A1").NumberFormat = "0.00"
        ws.Range("A2").NumberFormat = "General"
        ws.Range("A1:A2").Copy(ws.Range("A3:A60000"))  # alternating formats; Destination copy, no clipboard
        before = (ws.Range("A1").NumberFormat, ws.Range("A2").NumberFormat, ws.Range("A59999").NumberFormat)
        outcome = self.run("Tk_FormatRange", ws.Range("A1:A60000"), "number")
        after = (ws.Range("A1").NumberFormat, ws.Range("A2").NumberFormat, ws.Range("A59999").NumberFormat)
        self.record("too many distinct format blocks: refused, nothing changed",
                    outcome.startswith("refused:") and before == after, outcome=outcome)

    def shortcuts_toggle(self):
        self.run("Tk_SetShortcuts", True)
        on = self.run("Tk_ShortcutsEnabled")
        self.run("Tk_SetShortcuts", False)
        off = self.run("Tk_ShortcutsEnabled")
        self.record("shortcut toggle switches on and off without error (key behaviour: manual test)",
                    on is True and off is False)

    def performance(self):
        cases = []
        ws = self.wb.Worksheets.Add()
        cases.append(("uniform 10,000", "A1:A10000", lambda: ws.Range("A1:A10000").__setattr__("NumberFormat", "General")))

        def alternating(address, first, last):
            ws.Range(first).NumberFormat = "0.00"
            ws.Range(last).NumberFormat = "General"
            ws.Range(f"{first}:{last}").Copy(ws.Range(address))
        cases.append(("mixed 10,000 (alternating rows)", "C1:C10000", lambda: alternating("C1:C10000", "C1", "C2")))

        def by_column():
            for column, fmt in zip("EFG", ["0.00", "General", "0%"]):
                ws.Range(f"{column}1:{column}10000").NumberFormat = fmt
        cases.append(("mixed 30,000 (3 columns, each uniform)", "E1:G10000", by_column))
        cases.append(("whole column, uniform", "I:I", lambda: ws.Columns("I").__setattr__("NumberFormat", "General")))

        def column_exceptions():
            ws.Columns("K").NumberFormat = "General"
            for row in range(1, 1_000_001, 100_000):
                ws.Range(f"K{row}").NumberFormat = "0.00"
        cases.append(("whole column, 10 exceptions", "K:K", column_exceptions))

        for label, address, prepare in cases:
            prepare()
            started = time.perf_counter()
            outcome = self.run("Tk_FormatRange", ws.Range(address), "number")
            wall = time.perf_counter() - started
            timings = self.run("Tk_LastFormatTimings")
            started = time.perf_counter()
            undo = self.run("Tk_RestoreLast")
            undo_wall = time.perf_counter() - started
            entry = {"case": label, "outcome": outcome, "apply_wall_ms": round(wall * 1000, 1), "vba": timings,
                     "undo": undo, "undo_wall_ms": round(undo_wall * 1000, 1),
                     "undo_vba_ms": round(self.run("Tk_LastRestoreSeconds") * 1000, 1)}
            self.timings.append(entry)
            print(entry)

    def close(self):
        try:
            self.run("Tk_SetShortcuts", False)
            self.wb.Close(SaveChanges=False)
            self.addin.Close(SaveChanges=False)
            self.excel.Quit()
        finally:
            self.watchdog.stop()
            time.sleep(1)
            _terminate(self.pid)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlam", type=Path, default=DEFAULT_XLAM)
    parser.add_argument("--perf", action="store_true")
    parser.add_argument("--out", type=Path, default=HERE / "results" / f"m3a_live_{time.strftime('%Y%m%d_%H%M%S')}.json")
    args = parser.parse_args()

    personal_before = sha256(PERSONAL)
    check = Check(args.xlam)
    try:
        print("add-in version:", check.run("Tk_Version"))
        check.equivalence()
        check.undo_conflict_identity()
        check.protection_and_failures()
        check.too_many_blocks()
        check.shortcuts_toggle()
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
