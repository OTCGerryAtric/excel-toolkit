"""Phase B audit engine (OFFLINE-PREPARED — NOT run against live Excel here).

Handle ONE read-only AUDIT request and deliver structured findings to the Excel front end
through E4's existing COM callback. Reuses ``prototypes/e6/engine/excel_link`` (identity /
attach / call — byte-identical to E4's) and the project's detector + JSON transport.

Contract:
  * read-only: attach by identity, confirm the worksheet by name, snapshot the used range and
    run ``scan_worksheet``; the audited workbook is never written to;
  * never raises for an Excel-side problem — returns a response carrying ``status`` / ``error``
    plus the request ``id`` and correlation ``token``, so a failure can never be mistaken for
    an empty-but-successful audit (the front end trusts findings only when ``status == "ok"``);
  * findings travel as a FLAT JSON string (``display_json``) through the callback.

The real COM attach is injected (``attach_fn`` / ``call_fn``) so the whole engine is exercised
offline against a fake COM worksheet; the defaults use ``excel_link`` for the authorised live
run. The callback target ``PhaseBBridge.PhaseB_OnResult`` is a plain module in the OPEN
development workbook — no ``.xlam``, no install, no trust change.
"""

from __future__ import annotations

import gc
import sys
import time
from pathlib import Path

import pythoncom

ROOT = Path(__file__).resolve().parents[2]
for _p in (ROOT / "src", ROOT / "prototypes" / "e6" / "engine"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from excel_link import ExcelBusy, ExcelGone, Identity, IdentityError, attach, call  # noqa: E402
from excel_toolkit.formula_audit import display_json, scan_worksheet, snapshot_from_com  # noqa: E402

CALLBACK = "PhaseBBridge.PhaseB_OnResult"
ALLOWED_OPS = ("audit",)


def callback_name(macro_workbook: str | None) -> str:
    """Workbook-qualified callback target. An unqualified ``Application.Run`` name resolves
    only against the ACTIVE workbook (the audited one, which has no code), so the live run
    failed with "Cannot run the macro". The request supplies only the macro workbook's NAME;
    the procedure stays fixed here, so the pipe cannot be used to run arbitrary macros."""
    if not macro_workbook:
        return CALLBACK
    return "'" + macro_workbook.replace("'", "''") + "'!" + CALLBACK


def _default_attach(request):
    identity = Identity.from_handoff(
        f"{request['hwnd']}|{request['pid']}|{request['workbook']}|{request['full_name']}"
    )
    return attach(identity)


def _run_audit(request, attach_fn, call_fn):
    """Attach, confirm the worksheet, snapshot (read-only) and run the detector."""
    app, workbook = attach_fn(request)
    ws = None
    try:
        ws = call_fn(app, lambda: workbook.Worksheets(request["sheet"]))
        if call_fn(app, lambda: ws.Name) != request["sheet"]:
            raise IdentityError(f"worksheet {request['sheet']!r} is not the one that was requested")
        snapshot = call_fn(app, lambda: snapshot_from_com(ws))   # two bulk reads; no writes
        report = scan_worksheet(snapshot)
        return report, display_json(report, request["workbook"], request["full_name"])
    finally:
        app = workbook = ws = None  # noqa: F841 - drop COM references promptly
        gc.collect()
        pythoncom.CoFreeUnusedLibraries()


def handle_audit(request, attach_fn=_default_attach, call_fn=call) -> dict:
    """Read-only. Always returns a response dict; never raises for an Excel-side problem."""
    started = time.perf_counter()
    response = {
        "id": request.get("id"), "token": request.get("token"), "pid": request.get("pid"),
        "workbook": request.get("workbook"), "full_name": request.get("full_name"),
        "sheet": request.get("sheet"), "hwnd": request.get("hwnd"),
        "macro_workbook": request.get("macro_workbook"),
        "status": "ok", "error": "", "count": 0, "findings_json": "[]",
    }
    try:
        report, findings_json = _run_audit(request, attach_fn, call_fn)
        response["findings_json"] = findings_json
        response["count"] = len(report.findings)
    except (IdentityError, ExcelBusy, ExcelGone) as error:
        response.update(status="failed", error=f"{type(error).__name__}: {error}", findings_json="[]", count=0)
    except Exception as error:  # noqa: BLE001 - report, never crash the worker
        response.update(status="failed", error=f"{type(error).__name__}: {error}", findings_json="[]", count=0)
    response["elapsed_ms"] = (time.perf_counter() - started) * 1000
    return response


def deliver(response, attach_fn=None, call_fn=call, callback=None) -> str:
    """Call ``PhaseB_OnResult`` back in the requesting Excel. Passes BOTH the Excel identity
    (pid / workbook / full name) — so the front end can preserve E4's identity checks — and
    the correlation ``id`` + ``token``, then status / error / count / flat findings JSON.
    Returns VBA's verdict (or a 'not delivered: ...' note)."""
    if callback is None:
        callback = callback_name(response.get("macro_workbook"))
    try:
        if attach_fn is None:
            identity = Identity.from_handoff(
                f"{response['hwnd']}|{response['pid']}|{response['workbook']}|{response['full_name']}"
            )
            app, _ = attach(identity)
        else:
            app, _ = attach_fn(response)
        verdict = call_fn(app, lambda: app.Run(
            callback, response["id"], response["token"], response["pid"],
            response["workbook"], response["full_name"], response["status"],
            response["error"], response["count"], response["findings_json"],
        ))
    except (IdentityError, ExcelBusy, ExcelGone) as error:
        return f"not delivered: {type(error).__name__}: {error}"
    except Exception as error:  # noqa: BLE001
        return f"not delivered: {type(error).__name__}: {error}"
    finally:
        gc.collect()
        pythoncom.CoFreeUnusedLibraries()
    return verdict
