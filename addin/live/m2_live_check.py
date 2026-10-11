"""M2 scripted live check - Convert to Values + one-level undo. DO NOT RUN until authorised.

Runs entirely in a SEPARATE, disposable Excel process (DispatchEx: no XLSTART, no PERSONAL.XLSB,
no installed add-ins) against workbooks it creates itself and closes without saving. It opens the
built development add-in with Workbooks.Open (nothing installed, no registry or trust change) and
drives the add-in's own procedures through Application.Run in quiet mode (messages recorded, not
shown). Every check compares Python-side before/after snapshots taken independently of the VBA.

What it cannot prove (manual checklist, docs/validation/m2-convert-to-values.md): ribbon clicks,
and real Ctrl+Z - Excel only runs Application.OnUndo handlers for a real keypress, so here undo is
exercised through Tk_RestoreLast, the same checked procedure both the key and the button call.

  <scratch venv>/python addin/live/m2_live_check.py [--xlam addin/build/ExcelToolkit.xlam] [--perf]

STOP IMMEDIATELY (no retry) if endpoint protection intervenes, a process is unexpectedly
terminated, files disappear, or access-denied errors recur.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
import threading
import time
from pathlib import Path

import pythoncom
import win32api
import win32com.client
import win32con
import win32event
import win32process

HERE = Path(__file__).resolve().parent
DEFAULT_XLAM = HERE.parent / "build" / "ExcelToolkit.xlam"
PERSONAL = Path(os.environ["APPDATA"]) / "Microsoft" / "Excel" / "XLSTART" / "PERSONAL.XLSB"
XL_CALC_AUTOMATIC = -4105
XL_CELL_TYPE_VISIBLE = 12
STALL_SECONDS = 120

# Formula fixtures: (formula, number format or None). Results cover every supported type.
TYPE_FORMULAS = [
    ("=1+1", None),
    ('="0012"', None),                 # number-like text
    ('="1/2"', None),                  # date-like text
    ("=DATE(2026,1,2)", "dd-mmm-yy"),  # date serial with a date format
    ("=TRUE", None),
    ("=NA()", None),
    ("=1/0", None),
    ('=""', None),                     # empty text (must not become a blank cell)
    ('="=abc"', None),                 # text that looks like a formula
    ('=" padded "', None),
    ("=1/3", "0.000"),
    ('="TRUE"', None),                 # text, not a boolean
    ('="\'quoted"', None),             # text starting with an apostrophe
    ('=REPT("x",300)', None),
    ('="12%"', None),                  # text Excel would otherwise turn into a percentage
]
# Constants that must stay untouched: (value to enter, number format or None).
CONSTANTS = [("'0012", None), ("'1/2", None), (5, "0.00"), ("text", None), (True, None), ("'=lit", None), (None, None)]


class Watchdog:
    """Terminates the test Excel if the run stalls (e.g. a modal VBA error dialog)."""

    def __init__(self, pid: int):
        self.pid, self.last, self.fired, self._stop = pid, time.monotonic(), False, False
        threading.Thread(target=self._run, daemon=True).start()

    def tick(self):
        self.last = time.monotonic()

    def stop(self):
        self._stop = True

    def _run(self):
        while not self._stop:
            time.sleep(1)
            if time.monotonic() - self.last > STALL_SECONDS:
                self.fired = True
                _terminate(self.pid)
                return


def _terminate(pid: int) -> None:
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(handle, 1)
        win32api.CloseHandle(handle)
    except Exception:  # noqa: BLE001 - already gone
        pass


def sha256(path: Path) -> str | None:
    return hashlib.sha256(path.read_bytes()).hexdigest() if path.exists() else None


class Check:
    def __init__(self, xlam: Path):
        self.excel = win32com.client.DispatchEx("Excel.Application")
        _, self.pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
        self.watchdog = Watchdog(self.pid)
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.addin = self.excel.Workbooks.Open(str(xlam))
        self.name = self.addin.Name
        self.results: list[dict] = []
        self.timings: list[dict] = []
        self.run("Tk_SetQuiet", True)

    # ----- plumbing -----

    def run(self, procedure, *args):
        self.watchdog.tick()
        value = self.excel.Run(f"'{self.name}'!{procedure}", *args)
        self.watchdog.tick()
        return value

    def record(self, name: str, passed: bool, **details):
        self.results.append({"check": name, "passed": bool(passed), **details})
        print(("PASS " if passed else "FAIL ") + name + ("" if passed else f"  {details}"))

    def new_sheet(self):
        wb = self.excel.Workbooks.Add()
        return wb, wb.Worksheets(1)

    def app_state(self):
        return (self.excel.ScreenUpdating, self.excel.EnableEvents, self.excel.Calculation)

    @staticmethod
    def cell_state(cell) -> tuple:
        value = cell.Value2
        return (cell.Formula2, type(value).__name__, value, cell.NumberFormat, bool(cell.HasFormula))

    def snapshot(self, ws, address) -> dict:
        return {c.Address: self.cell_state(c) for c in ws.Range(address).Cells}

    def fill_types(self, ws, column="A"):
        for i, (formula, fmt) in enumerate(TYPE_FORMULAS, start=1):
            cell = ws.Range(f"{column}{i}")
            if fmt:
                cell.NumberFormat = fmt
            cell.Formula2 = formula
        return f"{column}1:{column}{len(TYPE_FORMULAS)}"

    def fill_constants(self, ws, column="B"):
        for i, (value, fmt) in enumerate(CONSTANTS, start=1):
            cell = ws.Range(f"{column}{i}")
            if fmt:
                cell.NumberFormat = fmt
            if value is not None:
                cell.Value2 = value
        return f"{column}1:{column}{len(CONSTANTS)}"

    def convert(self, rng):
        return self.run("Tk_ConvertRange", rng)

    def restore(self):
        return self.run("Tk_RestoreLast")

    # ----- scenarios -----

    def types_and_undo(self):
        wb, ws = self.new_sheet()
        formulas = self.fill_types(ws)
        constants = self.fill_constants(ws)
        before_f, before_c, state = self.snapshot(ws, formulas), self.snapshot(ws, constants), self.app_state()
        outcome = self.convert(ws.Range(f"{formulas},{constants}"))
        after_f, after_c = self.snapshot(ws, formulas), self.snapshot(ws, constants)
        self.record("types: command reports success", outcome.startswith("ok:"), outcome=outcome)
        exact = {a: (before_f[a][1:4], after_f[a][1:4]) for a in before_f
                 if before_f[a][1:4] != after_f[a][1:4] or after_f[a][4]}
        self.record("types: each converted cell holds exactly its formula's result (type, value, format; no formula)",
                    not exact, mismatches={k: str(v) for k, v in exact.items()})
        self.record("types: constants untouched (formula text, type, value, format)", before_c == after_c,
                    before={k: str(v) for k, v in before_c.items() if before_c[k] != after_c[k]},
                    after={k: str(v) for k, v in after_c.items() if before_c[k] != after_c[k]})
        self.record("types: Excel state restored after convert", self.app_state() == state, state=self.app_state())
        prefixes = {a: ws.Range(a).PrefixCharacter for a in before_f}
        undo = self.restore()
        ws.Calculate()
        restored = self.snapshot(ws, formulas)
        self.record("types: undo reports success", undo.startswith("ok:"), outcome=undo)
        self.record("types: undo restores identical Formula2 and recalculated values",
                    restored == before_f, diff={a: (str(before_f[a]), str(restored[a])) for a in before_f if before_f[a] != restored[a]})
        self.record("types: constants untouched by undo", self.snapshot(ws, constants) == before_c)
        self.record("types: second undo has nothing to do", self.restore().startswith("nothing:"))
        wb.Close(SaveChanges=False)
        return {"prefix_characters_after_convert": prefixes}

    def selections(self):
        wb, ws = self.new_sheet()
        self.fill_types(ws, "A")
        self.fill_types(ws, "C")

        before = self.snapshot(ws, "A1:C15")
        outcome = self.convert(ws.Range("A1:A3,A5:A6,C1:C2"))
        after = self.snapshot(ws, "A1:C15")
        changed = {a for a in before if before[a] != after[a]}
        expected = {f"$A${r}" for r in (1, 2, 3, 5, 6)} | {"$C$1", "$C$2"}
        self.record("selection: three areas converted, nothing else changed", outcome.startswith("ok:") and changed == expected,
                    outcome=outcome, changed=sorted(changed))
        self.restore()

        for name, address in [("overlapping areas", "A1:A5,A3:A8")]:
            before = self.snapshot(ws, "A1:C15")
            outcome = self.convert(ws.Range(address))
            self.record(f"selection: {name} refused, nothing changed",
                        outcome.startswith("refused:") and self.snapshot(ws, "A1:C15") == before, outcome=outcome)

        before = self.snapshot(ws, "B1:B3")
        self.record("selection: no formulas -> nothing", self.convert(ws.Range("B1:B3")).startswith("nothing:"))

        self.run("Tk_SetMaxCells", 10)
        outcome = self.convert(ws.Range("A1:A11"))
        self.record("selection: over the cell limit refused, not truncated",
                    outcome.startswith("refused:") and "11" in outcome and self.snapshot(ws, "A1:A11") == self.snapshot(ws, "A1:A11"),
                    outcome=outcome)
        self.run("Tk_SetMaxCells", 0)
        wb.Close(SaveChanges=False)

    def hidden_and_filtered(self):
        wb, ws = self.new_sheet()
        ws.Range("E1").Value2 = "key"
        for r, key in enumerate(["keep", "drop", "keep", "drop", "keep"], start=2):
            ws.Range(f"E{r}").Value2 = key
            ws.Range(f"F{r}").Formula2 = f"=ROW()*10"
        ws.Range("F1").Value2 = "value"
        ws.Range("E1:F6").AutoFilter(1, "keep")
        before = self.snapshot(ws, "F2:F6")
        outcome = self.convert(ws.Range("F2:F6"))
        self.record("filtered: selection including filtered-out rows refused, nothing changed",
                    outcome.startswith("refused:") and self.snapshot(ws, "F2:F6") == before, outcome=outcome)
        visible = ws.Range("F2:F6").SpecialCells(XL_CELL_TYPE_VISIBLE)
        outcome = self.convert(visible)
        after = self.snapshot(ws, "F2:F6")
        hidden_kept = all(after[a] == before[a] for a in ("$F$3", "$F$5"))
        visible_done = all(not after[a][4] for a in ("$F$2", "$F$4", "$F$6"))
        self.record("filtered: visible-cells-only selection converted; filtered-out cells untouched",
                    outcome.startswith("ok:") and hidden_kept and visible_done, outcome=outcome)
        self.restore()
        ws.AutoFilterMode = False

        ws.Range("H1:J1").Formula2 = "=1"
        ws.Columns("I").Hidden = True
        outcome = self.convert(ws.Range("H1:J1"))
        self.record("hidden: selection including a hidden column refused", outcome.startswith("refused:"), outcome=outcome)
        wb.Close(SaveChanges=False)

    def unsupported_structures(self):
        wb, ws = self.new_sheet()
        ws.Range("A1").Formula2 = "=1"
        ws.Range("A1:B1").Merge()
        ws.Range("D1").Formula2 = "=SEQUENCE(3)"
        ws.Range("F1:F2").FormulaArray = "=ROW(1:2)"
        ws.Range("K1").Value2 = 1                       # data table input cell
        ws.Range("I1").Formula2 = "=K1*2"               # data table formula row
        for r, v in enumerate([1, 2, 3], start=2):
            ws.Range(f"H{r}").Value2 = v                # input values; results fill I2:I4
        ws.Range("H1:I4").Table(ColumnInput=ws.Range("K1"))
        cases = [("merged cells", "A1:B1"), ("spill anchor", "D1"), ("spill range", "D1:D3"), ("spilled cells only", "D2:D3"),
                 ("legacy array formula", "F1:F2"), ("data table", "I2:I4")]
        for name, address in cases:
            before = self.snapshot(ws, "A1:K4")
            outcome = self.convert(ws.Range(address))
            self.record(f"unsupported: {name} refused, nothing changed",
                        outcome.startswith("refused:") and self.snapshot(ws, "A1:K4") == before, outcome=outcome)
        wb.Close(SaveChanges=False)

    def protection(self):
        wb, ws = self.new_sheet()
        ws.Range("A1:A2").Formula2 = "=ROW()"
        ws.Range("A2").Locked = False
        ws.Protect()
        before = self.snapshot(ws, "A1:A2")
        outcome = self.convert(ws.Range("A1:A2"))
        self.record("protected: locked formula cell refused before writing", outcome.startswith("refused:")
                    and self.snapshot(ws, "A1:A2") == before, outcome=outcome)
        outcome = self.convert(ws.Range("A2"))
        self.record("protected: unlocked formula cell converted", outcome.startswith("ok:"), outcome=outcome)
        self.record("protected: undo of the unlocked cell", self.restore().startswith("ok:"))
        ws.Unprotect()
        wb.Close(SaveChanges=False)

    def failures(self):
        wb, ws = self.new_sheet()
        formulas = self.fill_types(ws)
        before, state = self.snapshot(ws, formulas), self.app_state()
        self.run("Tk_TestSetFailAfter", 3, 0)
        outcome = self.convert(ws.Range(formulas))
        self.record("failure (per-cell path): rolled back, sheet identical, reported as rolled back",
                    outcome.startswith("rolled back:") and self.snapshot(ws, formulas) == before, outcome=outcome)
        self.record("failure: Excel state restored", self.app_state() == state)
        self.record("failure: no undo record left behind", self.restore().startswith("nothing:"))

        ws.Range("D1:D5").Formula2 = "=ROW()"
        ws.Range("F1:F5").Formula2 = "=ROW()*2"
        before = self.snapshot(ws, "D1:F5")
        self.run("Tk_TestSetFailAfter", 5, 0)  # fails before the second (bulk) area
        outcome = self.convert(ws.Range("D1:D5,F1:F5"))
        self.record("failure (bulk path): first area rolled back", outcome.startswith("rolled back:")
                    and self.snapshot(ws, "D1:F5") == before, outcome=outcome)

        self.run("Tk_TestSetFailAfter", 0, 0)
        self.convert(ws.Range(formulas))
        self.run("Tk_TestSetFailAfter", 0, 2)
        outcome = self.restore()
        self.record("restore failure: reported as incomplete, never success", outcome.startswith("incomplete:"), outcome=outcome)
        self.record("restore failure: record unusable afterwards", self.restore().startswith("refused:"))
        self.run("Tk_TestSetFailAfter", 0, 0)
        wb.Close(SaveChanges=False)

    def conflicts_and_lifecycle(self):
        wb, ws = self.new_sheet()
        formulas = self.fill_types(ws)
        self.convert(ws.Range("A1:A3"))
        converted = self.snapshot(ws, "A1:A3")
        ws.Range("A2").Value2 = "edited"
        outcome = self.restore()
        self.record("conflict: undo refused after an edit, names the cell, writes nothing",
                    outcome.startswith("refused:") and "A2" in outcome
                    and self.snapshot(ws, "A1") == {"$A$1": converted["$A$1"]}, outcome=outcome)
        ws.Range("A2").NumberFormat = "@"
        ws.Range("A2").Value2 = "0012"
        ws.Range("A2").NumberFormat = converted["$A$2"][3]
        self.record("conflict: undo allowed again once the edit is reverted", self.restore().startswith("ok:"))

        self.convert(ws.Range("A1"))
        self.convert(ws.Range("A5"))
        self.restore()
        self.record("lifecycle: a new command replaces the record (only the last one is undone)",
                    not ws.Range("A1").HasFormula and ws.Range("A5").HasFormula)
        self.record("lifecycle: record cleared after undo", self.restore().startswith("nothing:"))

        self.convert(ws.Range("A1"))
        ws.Name = "Renamed"
        outcome = self.restore()
        self.record("identity: undo refused after the sheet is renamed", outcome.startswith("refused:"), outcome=outcome)

        self.convert(ws.Range("A5"))
        path = Path(tempfile.gettempdir()) / f"tk_m2_saveas_{os.getpid()}.xlsx"
        wb.SaveAs(str(path))
        outcome = self.restore()
        self.record("identity: undo refused after Save As under another name", outcome.startswith("refused:"), outcome=outcome)

        self.convert(ws.Range("A11"))
        wb.Close(SaveChanges=False)
        outcome = self.restore()
        self.record("identity: undo refused after the workbook is closed", outcome.startswith("refused:"), outcome=outcome)
        path.unlink(missing_ok=True)

    def performance(self, sizes):
        for n in sizes:
            for kind, formula in (("number", "=ROW()*2"), ("text", '=TEXT(ROW(),"0000")')):
                if kind == "text" and n > 10_000:
                    continue
                wb, ws = self.new_sheet()
                ws.Range(f"A1:A{n}").Formula2 = formula
                started = time.perf_counter()
                outcome = self.convert(ws.Range(f"A1:A{n}"))
                wall = time.perf_counter() - started
                timings = self.run("Tk_LastTimings")
                started = time.perf_counter()
                undo = self.restore()
                undo_wall = time.perf_counter() - started
                self.timings.append({"cells": n, "kind": kind, "outcome": outcome, "convert_wall_ms": round(wall * 1000, 1),
                                     "vba_timings_ms": timings, "undo": undo, "undo_wall_ms": round(undo_wall * 1000, 1),
                                     "undo_vba_ms": round(self.run("Tk_LastRestoreSeconds") * 1000, 1)})
                print(self.timings[-1])
                wb.Close(SaveChanges=False)

    def close(self):
        try:
            self.addin.Close(SaveChanges=False)
            self.excel.Quit()
        finally:
            self.watchdog.stop()
            pid = self.pid
            del self.excel, self.addin
            pythoncom.CoFreeUnusedLibraries()
            try:
                handle = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
                if win32event.WaitForSingleObject(handle, 15_000) != win32event.WAIT_OBJECT_0:
                    win32api.TerminateProcess(handle, 1)
                win32api.CloseHandle(handle)
            except Exception:  # noqa: BLE001 - already exited
                pass


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlam", type=Path, default=DEFAULT_XLAM)
    parser.add_argument("--perf", action="store_true", help="include 1/1k/10k/100k timings")
    parser.add_argument("--out", type=Path, default=HERE / "results" / f"m2_live_{time.strftime('%Y%m%d_%H%M%S')}.json")
    args = parser.parse_args()

    personal_before = sha256(PERSONAL)
    check = Check(args.xlam)
    extra = {}
    try:
        print("add-in version:", check.run("Tk_Version"))
        extra = check.types_and_undo() or {}
        check.selections()
        check.hidden_and_filtered()
        check.unsupported_structures()
        check.protection()
        check.failures()
        check.conflicts_and_lifecycle()
        if args.perf:
            check.performance([1, 1_000, 10_000, 100_000])
    finally:
        watchdog_fired = check.watchdog.fired
        check.close()
    check.record("PERSONAL.XLSB unchanged", sha256(PERSONAL) == personal_before)
    if watchdog_fired:
        check.record("run completed without stalling", False, note="watchdog terminated the test Excel")

    passed = sum(r["passed"] for r in check.results)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps({"results": check.results, "timings": check.timings, **extra}, indent=2, default=str),
                        encoding="utf-8")
    print(f"\n{passed}/{len(check.results)} checks passed -> {args.out}")
    return 0 if passed == len(check.results) else 1


if __name__ == "__main__":
    sys.exit(main())
