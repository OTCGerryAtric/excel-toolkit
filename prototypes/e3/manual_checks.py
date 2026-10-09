"""E3 manual checks - run in your own PowerShell window (it asks you to press keys).

Opens a NEW, separate Excel with a blank scratch workbook and the E3 test add-in.
Your own Excel and PERSONAL.XLSB are not touched. Nothing is saved.

  uv run python prototypes/e3/manual_checks.py
"""

import json
import sys
import time
from pathlib import Path

import win32com.client
import win32process

sys.path.insert(0, str(Path(__file__).resolve().parent))
from excel_link import ExcelBusy, Identity, attach  # noqa: E402
from journal import read_block  # noqa: E402

HERE = Path(__file__).resolve().parent
XLAM = HERE / "work" / "E3Commands.xlam"
OUT = HERE / "results" / "e3_manual.json"


def ask(prompt: str) -> str:
    return input(f"\n{prompt}\n> ").strip()


def main() -> None:
    if not XLAM.exists():
        sys.exit("Run prototypes/e3/run_e3.py first (it builds the test add-in).")
    results = {}
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    try:
        book = excel.Workbooks.Add()
        sheet = book.Worksheets(1)
        sheet.Range("B2:D6").Value = 1234.5
        sheet.Range("B2:D6").NumberFormat = "0.00"
        excel.Workbooks.Open(str(XLAM))
        excel.Run("E3Commands.xlam!E3_Init")
        book.Activate()
        sheet.Range("B2:D6").Select()
        pid = win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1]
        identity = Identity.from_handoff(f"{int(excel.Hwnd)}|{pid}|{book.Name}|{book.FullName}")

        print("A new Excel window has opened with a scratch workbook (cells B2:D6 selected, format 0.00).")
        ask("CHECK 1 - Ctrl+Z after a toggle.\n"
            "In that NEW Excel window: press Ctrl+Shift+J once (format changes to #,##0),\n"
            "then press Ctrl+Z.\nThen come back here and press Enter.")
        results["ctrl_z_after_toggle"] = {
            "format_now": sheet.Range("B2").NumberFormat,
            "restored_to_0.00": sheet.Range("B2").NumberFormat == "0.00",
            "user_says": ask("Did Excel's Undo menu / Ctrl+Z say 'Undo number format' and put the format back? (y/n)"),
        }

        ask("CHECK 2 - Excel busy (cell edit mode).\n"
            "In the NEW Excel window: click cell F2, type 123 but DO NOT press Enter (stay in the cell).\n"
            "Then come back here and press Enter. The script will try to read for 5 seconds.")
        started = time.monotonic()
        try:
            app, workbook = attach(identity, timeout=5)
            value = read_block(app, workbook, sheet.Name, "B2")
            results["edit_mode"] = {"outcome": "read succeeded", "seconds": round(time.monotonic() - started, 1), "value": value}
        except ExcelBusy as error:
            results["edit_mode"] = {"outcome": f"ExcelBusy: {error}", "seconds": round(time.monotonic() - started, 1)}
        print(f"Result: {results['edit_mode']}")

        ask("Now press Escape in the Excel cell, come back here and press Enter. The script retries once more.")
        try:
            app, workbook = attach(identity, timeout=5)
            results["after_escape"] = {"outcome": "read succeeded", "value": read_block(app, workbook, sheet.Name, "B2")}
        except ExcelBusy as error:
            results["after_escape"] = {"outcome": f"ExcelBusy: {error}"}
        print(f"Result: {results['after_escape']}")
    finally:
        OUT.write_text(json.dumps(results, indent=1, default=str), encoding="utf-8")
        print(f"\nSaved {OUT}. Closing the test Excel without saving.")
        try:
            for workbook in list(excel.Workbooks):
                workbook.Close(SaveChanges=False)
            excel.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    main()
