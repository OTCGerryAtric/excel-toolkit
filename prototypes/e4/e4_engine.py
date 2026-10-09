"""E4 shared engine: handle one read-only request and deliver the result to VBA.

Used by both e4_oneshot.py and e4_worker.py so the comparison is like for like.
Short-lived COM connections: attach per request, release everything afterwards.
"""

import gc
import re
import sys
import time
from pathlib import Path

import pythoncom

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "e3"))
from excel_link import ExcelBusy, ExcelGone, Identity, IdentityError, attach, call  # noqa: E402
from journal import as_2d, clean  # noqa: E402

CALLBACK = "E4Bridge.xlam!E4_OnResult"
NUMBER_LITERAL = re.compile(r'(?<![A-Za-z_$\d.])\d+(\.\d+)?(?![\d(:A-Za-z_])')


def column_letters(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def split_address(address: str) -> tuple[int, int]:
    match = re.match(r"\$?([A-Z]+)\$?(\d+)", address.split(":")[0])
    column = 0
    for char in match.group(1):
        column = column * 26 + ord(char) - 64
    return int(match.group(2)), column


def has_number_literal(formula: str) -> bool:
    without_strings = re.sub(r'"[^"]*"', "", formula)
    return bool(NUMBER_LITERAL.search(without_strings))


def classify(sheet: str, address: str, formulas: list[list], values: list[list]) -> list[tuple]:
    top, left = split_address(address)
    findings = []
    for r, row in enumerate(formulas):
        for c, formula in enumerate(row):
            value = values[r][c]
            if isinstance(formula, str) and formula.startswith("="):
                if isinstance(value, dict):
                    kind = "formula_error"
                elif has_number_literal(formula):
                    kind = "formula_with_constant"
                else:
                    kind = "formula"
                detail = formula
            elif value is None or value == "":
                continue
            elif isinstance(value, dict):
                kind, detail = "error_constant", value["error"]
            elif isinstance(value, bool):
                kind, detail = "boolean", str(value)
            elif isinstance(value, (int, float)):
                kind, detail = "number", str(value)
            else:
                kind, detail = "text", str(value)
            findings.append((sheet, f"{column_letters(left + c)}{top + r}", kind, detail))
    return findings


def handle(request: dict) -> dict:
    """Read-only. Returns a response dict; never raises for Excel-side problems."""
    timings = {}
    response = {"id": request["id"], "pid": request["pid"], "workbook": request["workbook"],
                "full_name": request["full_name"], "status": "ok", "error": "", "probe": None,
                "count": 0, "findings": []}
    started = time.perf_counter()
    try:
        identity = Identity.from_handoff(f"{request['hwnd']}|{request['pid']}|{request['workbook']}|{request['full_name']}")
        app, workbook = attach(identity)
        timings["attach_ms"] = (time.perf_counter() - started) * 1000
        started = time.perf_counter()
        sheet = call(app, lambda: workbook.Worksheets(request["sheet"]))
        rng = call(app, lambda: sheet.Range(request["address"]))
        formulas = as_2d(call(app, lambda: rng.Formula2))
        values = [[clean(v) for v in row] for row in as_2d(call(app, lambda: rng.Value2))]
        response["probe"] = call(app, lambda: workbook.Worksheets("Inputs").Range("A1").Value2)
        if request["op"] == "classify":
            response["findings"] = classify(request["sheet"], request["address"], formulas, values)
            response["count"] = len(response["findings"])
        else:  # "read": size benchmark, no findings
            response["count"] = sum(len(row) for row in formulas)
        timings["work_ms"] = (time.perf_counter() - started) * 1000
    except (IdentityError, ExcelBusy, ExcelGone) as error:
        response.update(status="failed", error=f"{type(error).__name__}: {error}")
    except Exception as error:  # noqa: BLE001 - report, don't crash the worker
        response.update(status="failed", error=f"{type(error).__name__}: {error}")
    finally:
        app = workbook = sheet = rng = None  # noqa: F841 - drop COM references
        gc.collect()
        pythoncom.CoFreeUnusedLibraries()
    response["timings"] = timings
    return response


def deliver(response: dict) -> str:
    """Call back into the requesting Excel process. Returns VBA's verdict or the failure."""
    started = time.perf_counter()
    try:
        identity = Identity.from_handoff(
            f"{_hwnd_for(response)}|{response['pid']}|{response['workbook']}|{response['full_name']}"
        )
        app, _ = attach(identity)
        findings = tuple(response["findings"]) if response["findings"] else ""
        verdict = call(app, lambda: app.Run(
            CALLBACK, response["id"], response["pid"], response["workbook"], response["full_name"],
            response["status"], response["error"], response["probe"], response["count"], findings,
        ))
    except (IdentityError, ExcelBusy, ExcelGone) as error:
        verdict = f"not delivered: {type(error).__name__}: {error}"
    except Exception as error:  # noqa: BLE001
        verdict = f"not delivered: {type(error).__name__}: {error}"
    finally:
        app = None  # noqa: F841
        gc.collect()
        pythoncom.CoFreeUnusedLibraries()
    response.setdefault("timings", {})["deliver_ms"] = (time.perf_counter() - started) * 1000
    return verdict


def _hwnd_for(response: dict) -> int:
    return response["hwnd"]
