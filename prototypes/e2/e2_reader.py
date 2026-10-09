"""E2 reader: attach to one specific Excel instance and workbook, read it, write JSON.

Read-only. Identity comes from the caller (VBA): main window handle, process id and
workbook name. Backends:
  pywin32  - attach through the workbook window (WM_GETOBJECT / OBJID_NATIVEOM)
  xlwings  - free open-source core: xw.apps[pid].books[name]
  naive    - GetActiveObject("Excel.Application"), for comparison only
"""

import time

T0 = time.perf_counter()

import faulthandler  # noqa: E402
import argparse  # noqa: E402
import json  # noqa: E402
import os  # noqa: E402
import sys  # noqa: E402
from pathlib import Path  # noqa: E402

ERROR_NAMES = {
    2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?",
    2036: "#NUM!", 2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!", 2046: "#CONNECT!",
    2047: "#BLOCKED!", 2048: "#UNKNOWN!", 2049: "#FIELD!", 2050: "#CALC!",
}
COM_ERROR_BASE = -2146828288  # 0x800A0000 as a signed int; COM error value = base + Excel code
SMALL_SHEETS = ("Inputs", "Calc", "Errors", "Dyn", "Hidden", "VeryHidden", "Protected")
BULK_RANGES = {"1k": "A1:J100", "10k": "A1:J1000"}
WM_GETOBJECT = 0x003D
OBJID_NATIVEOM = -16
SMTO_ABORTIFHUNG = 0x0002


class HandoffError(Exception):
    pass


def as_2d(value) -> list[list]:
    if isinstance(value, (tuple, list)):
        return [list(row) if isinstance(row, (tuple, list)) else [row] for row in value]
    return [[value]]


def clean_com_value(value):
    if isinstance(value, int) and not isinstance(value, bool):
        code = value - COM_ERROR_BASE
        if 2000 <= code <= 2100:
            return {"error": ERROR_NAMES.get(code, f"#ERR{code}")}
    return value


def clean_xlwings_value(value):
    # err_to_str returns errors as strings, indistinguishable from text that looks like an error.
    if isinstance(value, str) and value in ERROR_NAMES.values():
        return {"error": value}
    return value


FORMULA_UNREADABLE = "<formula not readable>"


def read_formulas(read) -> list[list]:
    # Excel refuses Formula2 on protected sheets with hidden formulas.
    try:
        return as_2d(read())
    except Exception:  # noqa: BLE001 - pywintypes.com_error, surfaced through xlwings too
        return [[FORMULA_UNREADABLE]]


# ---------- Attach ----------

def check_window_identity(hwnd: int, pid: int) -> None:
    import win32gui
    import win32process

    if not win32gui.IsWindow(hwnd):
        raise HandoffError(f"window {hwnd} does not exist")
    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
    if window_pid != pid:
        raise HandoffError(f"window {hwnd} belongs to process {window_pid}, not {pid}")


def check_app_pid(app_hwnd: int, pid: int) -> None:
    # Since Excel 2013 each workbook has its own top-level window and Application.Hwnd
    # follows the active one, so the process id is the stable identity, not the handle.
    import win32process

    app_pid = win32process.GetWindowThreadProcessId(app_hwnd)[1]
    if app_pid != pid:
        raise HandoffError(f"attached to Excel process {app_pid}, expected {pid}")


def attach_pywin32(hwnd: int, pid: int):
    import pythoncom
    import win32com.client
    import win32gui

    check_window_identity(hwnd, pid)
    desk = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
    workbook_window = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
    if not workbook_window:
        raise HandoffError("no workbook window found under the Excel main window")
    _, lresult = win32gui.SendMessageTimeout(
        workbook_window, WM_GETOBJECT, 0, OBJID_NATIVEOM, SMTO_ABORTIFHUNG, 5000
    )
    if not lresult:
        raise HandoffError("Excel did not return its object model (busy or hung?)")
    window = win32com.client.Dispatch(
        pythoncom.ObjectFromLresult(lresult, pythoncom.IID_IDispatch, 0)
    )
    app = window.Application
    check_app_pid(int(app.Hwnd), pid)
    return app


def find_workbook_pywin32(app, name: str):
    for workbook in app.Workbooks:
        if workbook.Name == name:
            return workbook
    raise HandoffError(f"workbook {name!r} is not open in the target Excel instance")


def attach_xlwings(hwnd: int, pid: int):
    import xlwings as xw

    check_window_identity(hwnd, pid)
    if pid not in xw.apps.keys():
        raise HandoffError(f"xlwings found no Excel instance with process id {pid}")
    app = xw.apps[pid]
    check_app_pid(int(app.hwnd), pid)
    return app


def find_workbook_xlwings(app, name: str):
    names = [book.name for book in app.books]
    if name not in names:
        raise HandoffError(f"workbook {name!r} is not open in the target Excel instance")
    return app.books[name]


# ---------- Read ----------

def read_pywin32(workbook) -> dict:
    selection = workbook.Windows(1).RangeSelection
    areas = selection.Areas
    out = {
        "name": workbook.Name,
        "full_name": workbook.FullName,
        "saved": bool(workbook.Saved),
        "selection": {
            "sheet": selection.Worksheet.Name,
            "areas": [areas.Item(i).Address for i in range(1, areas.Count + 1)],
        },
        "sheets": {},
        "api_fallbacks": [],
    }
    for name in SMALL_SHEETS:
        sheet = workbook.Worksheets(name)
        used = sheet.UsedRange
        out["sheets"][name] = {
            "visible": int(sheet.Visible),
            "protected": bool(sheet.ProtectContents),
            "used_range": used.Address,
            "formula2": read_formulas(lambda: used.Formula2),
            "value": [[clean_com_value(v) for v in row] for row in as_2d(used.Value2)],
        }
    dyn = workbook.Worksheets("Dyn").Range("A1")
    out["spill"] = {"Dyn!A1_has_spill": bool(dyn.HasSpill), "Dyn!A1_spill": dyn.SpillingToRange.Address}
    return out


def read_xlwings(book) -> dict:
    selection = book.selection
    out = {
        "name": book.name,
        "full_name": book.fullname,
        "saved": bool(book.api.Saved),
        "selection": {
            "sheet": selection.sheet.name,
            "areas": [area.Address for area in selection.api.Areas] if selection is not None else None,
        },
        "sheets": {},
        "api_fallbacks": [
            "Workbook.Saved",
            "Range.Areas (multi-area selection)",
            "Worksheet.Visible (xlwings .visible is bool; cannot tell hidden from very hidden)",
            "Worksheet.ProtectContents",
            "Range.HasSpill / SpillingToRange",
        ],
    }
    for name in SMALL_SHEETS:
        sheet = book.sheets[name]
        used = sheet.used_range
        out["sheets"][name] = {
            "visible": int(sheet.api.Visible),
            "protected": bool(sheet.api.ProtectContents),
            "used_range": used.address,
            "formula2": read_formulas(lambda: used.formula2),
            "value": [
                [clean_xlwings_value(v) for v in row]
                for row in used.options(ndim=2, err_to_str=True).value
            ],
        }
    dyn = book.sheets["Dyn"].range("A1").api
    out["spill"] = {"Dyn!A1_has_spill": bool(dyn.HasSpill), "Dyn!A1_spill": dyn.SpillingToRange.Address}
    return out


def time_bulk_pywin32(workbook) -> dict:
    sheet = workbook.Worksheets("Bulk")
    timings = {}
    for label, address in BULK_RANGES.items():
        started = time.perf_counter()
        formulas = sheet.Range(address).Formula2
        values = sheet.Range(address).Value2
        timings[label] = {"ms": (time.perf_counter() - started) * 1000, "cells": len(formulas) * len(formulas[0]),
                          "first": [formulas[0][0], values[0][0]], "last": [formulas[-1][-1], values[-1][-1]]}
    return timings


def time_bulk_xlwings(book) -> dict:
    sheet = book.sheets["Bulk"]
    timings = {}
    for label, address in BULK_RANGES.items():
        started = time.perf_counter()
        formulas = sheet.range(address).formula2
        values = sheet.range(address).options(ndim=2).value
        timings[label] = {"ms": (time.perf_counter() - started) * 1000, "cells": len(formulas) * len(formulas[0]),
                          "first": [formulas[0][0], values[0][0]], "last": [formulas[-1][-1], values[-1][-1]]}
    return timings


def run_naive(name: str) -> dict:
    import win32com.client
    import win32process

    app = win32com.client.GetActiveObject("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(int(app.Hwnd))
    names = [workbook.Name for workbook in app.Workbooks]
    result = {"attached_pid": pid, "attached_hwnd": int(app.Hwnd), "workbook_count": len(names),
              "target_present": name in names}
    if name in names:
        result["Inputs!A1"] = app.Workbooks(name).Worksheets("Inputs").Range("A1").Value2
        result["full_name"] = app.Workbooks(name).FullName
    return result


# ---------- Main ----------

def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--backend", choices=["pywin32", "xlwings", "naive"], required=True)
    parser.add_argument("--hwnd", type=int, required=True)
    parser.add_argument("--pid", type=int, required=True)
    parser.add_argument("--workbook", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    faulthandler.dump_traceback_later(30, exit=True)  # prototype: dump the stack and exit if hung

    timings = {}
    result: dict = {"backend": args.backend, "args": {"hwnd": args.hwnd, "pid": args.pid, "workbook": args.workbook}}
    exit_code = 0
    try:
        started = time.perf_counter()
        if args.backend == "naive":
            result["naive"] = run_naive(args.workbook)
        elif args.backend == "pywin32":
            app = attach_pywin32(args.hwnd, args.pid)
            timings["attach_ms"] = (time.perf_counter() - started) * 1000
            workbook = find_workbook_pywin32(app, args.workbook)
            started = time.perf_counter()
            result["data"] = read_pywin32(workbook)
            timings["read_small_ms"] = (time.perf_counter() - started) * 1000
            timings["bulk"] = time_bulk_pywin32(workbook)
        else:
            app = attach_xlwings(args.hwnd, args.pid)
            timings["attach_ms"] = (time.perf_counter() - started) * 1000
            book = find_workbook_xlwings(app, args.workbook)
            started = time.perf_counter()
            result["data"] = read_xlwings(book)
            timings["read_small_ms"] = (time.perf_counter() - started) * 1000
            timings["bulk"] = time_bulk_xlwings(book)
        result["ok"] = True
    except HandoffError as error:
        result.update(ok=False, error_type="HandoffError", error=str(error))
        exit_code = 2
    except Exception as error:  # noqa: BLE001 - prototype: record anything else
        result.update(ok=False, error_type=type(error).__name__, error=str(error))
        exit_code = 3
    timings["process_ms"] = (time.perf_counter() - T0) * 1000
    result["timings"] = timings

    out = Path(args.out)
    tmp = out.with_suffix(".tmp")
    tmp.write_text(json.dumps(result, indent=1, default=str), encoding="utf-8")
    os.replace(tmp, out)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
