"""M3c syntax probe: how Excel itself writes and reads references whose sheet names look like cell, column
or row references, and how it treats ranges whose halves name a sheet. Evidence for the parser rules in
TkFormulas (RangeBeforeColon and the ambiguity check).

Runs in a SEPARATE disposable Excel process (DispatchEx: no XLSTART, no PERSONAL.XLSB, no add-ins) on a
workbook it creates and closes without saving. No add-in, no macros.

  <scratch venv>/python addin/live/m3c_syntax_probe.py [--out results/m3c_probe_<time>.json]
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import win32com.client
import win32process

from m2_live_check import PERSONAL, Watchdog, _terminate, sha256

HERE = Path(__file__).resolve().parent

# Sheet names that look like (or nearly like) references, plus ordinary ones. Each gets B2 = its position.
NAMES = ["Sheet1", "Sheet2", "A", "AB", "XFD", "XFE", "A1", "XFD1", "XFE1", "1", "R", "C", "RC", "R1C1", "Data", "Other"]

# Formulas entered on Sheet1 (the sheet Clear Sheet References would run on), to see what Excel stores and
# what they calculate to.
PROBES = [
    "=SUM(A:Sheet1!B2)",          # sheets A..Sheet1, or column A to Sheet1!B2?
    "=SUM(AB:Sheet1!B2)",
    "=SUM(XFE:Sheet1!B2)",
    "=SUM(A1:Sheet1!B2)",
    "=SUM(XFE1:Sheet1!B2)",
    "=SUM(Data:Sheet1!B2)",
    "=SUM(Sheet1!A:Sheet1!C)",
    "=SUM(Sheet1!1:Sheet1!3)",
    "=SUM(Sheet1!A1:Sheet1!B2)",
    "=SUM(Sheet1!A1:B2)",
    "=SUM(Sheet2!A1:Sheet1!B2)",
    "=SUM(Sheet2!A1:B2)",
    "=SUM(A:B2)",
    "=SUM(1:Sheet1!B2)",
    "=SUM(INDEX(Sheet2!A:A,1):Sheet1!B2)",
    "=SUM(INDEX(Sheet2!A:A,1):B2)",
    "=SUM(INDEX(A:A,1):Sheet1!B2)",
    "=SUM(INDEX(A:A,1):B2)",
    "=SUM($A:Sheet1!B2)",
    "=SUM($A$1:Sheet1!B2)",
    "=SUM(A$1:Sheet1!B2)",
    "=SUM($1:Sheet1!B2)",
    "=SUM(XFD1048576:Sheet1!B2)",
    "=SUM(XFD1048577:Sheet1!B2)",
    "=SUM(A0:Sheet1!B2)",
    "=SUM(Sheet2!A1:'Sheet1'!B2)",
    "=SUM(A1:'Sheet1'!B2)",
    "=SUM(AB:'Sheet1'!B2)",
    "=SUM(Sheet1!A1:Sheet2!B2)",
    "=SUM(A1:Sheet2!B2)",
    "=SUM(Sheet1!A:A)",
    "=SUM(Sheet1!A:C)",
    "=SUM(Sheet1!1:3)",
    "=SUM(Sheet1!A1:A!B2)",
    "=SUM(A1:A!B2)",
    "=SUM(Sheet1!A1:INDEX(A:A,3))",
    "=SUM(Sheet2!A1:INDEX(A:A,3))",
    "=SUM(Sheet2!A1:INDEX(Sheet1!A:A,3))",
    "=SUM(XFD1:Sheet1!B2)",
    "=SUM(B1048576:Sheet1!B2)",
    "=SUM(XFD1048576:B2)",
    "=SUM(C5:Sheet1!B2)",
    "=SUM(C5:B2)",
    "=SUM(Sheet1!C5:Sheet1!B2)",
]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", type=Path, default=HERE / "results" / f"m3c_probe_{time.strftime('%Y%m%d_%H%M%S')}.json")
    args = parser.parse_args()
    personal_before = sha256(PERSONAL)
    excel = win32com.client.DispatchEx("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
    watchdog = Watchdog(pid)
    out: dict = {"names": {}, "probes": []}
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        while wb.Worksheets.Count < len(NAMES):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for index in range(1, len(NAMES) + 1):
            wb.Worksheets(index).Name = f"tmp{index}"
        for index, name in enumerate(NAMES, start=1):
            ws = wb.Worksheets(index)
            ws.Name = name
            ws.Range("A1:C3").Value2 = tuple(tuple(index * 1000 + r * 10 + c for c in range(3)) for r in range(3))
        own = wb.Worksheets("Sheet1")
        watchdog.tick()

        # 1. How Excel writes a reference to each sheet: quoted or not.
        for row, name in enumerate(NAMES, start=10):
            cell = own.Range(f"E{row}")
            cell.Formula2 = "='" + name.replace("'", "''") + "'!B2"
            out["names"][name] = {"stored": cell.Formula2, "value": cell.Value2}
            print(f"{name!r:10} -> {cell.Formula2!r:16} = {cell.Value2}")
            watchdog.tick()

        # 2. What the probe formulas store and calculate to.
        for row, formula in enumerate(PROBES, start=10):
            cell = own.Range(f"G{row}")
            try:
                cell.Formula2 = formula
                entry = {"entered": formula, "stored": cell.Formula2, "value": cell.Value2}
            except Exception as exc:  # noqa: BLE001 - Excel rejected the formula
                entry = {"entered": formula, "error": str(exc).splitlines()[0][:200]}
            out["probes"].append(entry)
            print(entry)
            watchdog.tick()
        wb.Close(SaveChanges=False)
        excel.Quit()
    finally:
        watchdog.stop()
        del excel
        time.sleep(1)
        _terminate(pid)
    out["personal_unchanged"] = sha256(PERSONAL) == personal_before
    print("PERSONAL.XLSB unchanged:", out["personal_unchanged"], "| watchdog fired:", watchdog.fired)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(out, indent=2, default=str), encoding="utf-8")
    print("->", args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
