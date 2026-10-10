"""STAGE 2 - first live-Excel test. DO NOT RUN until authorised.

Preconditions (all required before running):
  * you have confirmed this kind of Python/Excel-COM automation is permitted on this laptop;
  * prototypes/integration/demo_workbook.xlsx is OPEN in Excel;
  * you have edited D3 on sheet 'Demo' from  =C3*(1+$A$10)  to  =C3*(1+$A$11)  and have
    NOT saved (this is the deliberate unsaved anomaly);
  * run with the scratch venv that has pywin32 + openpyxl.

First execution is READ-ONLY: it attaches, verifies identity, reads live formulas, runs the
detector, and checks integrity. It does NOT create a results workbook or navigate - that is
a separate step (show_and_navigate, defined but intentionally not called).

STOP IMMEDIATELY (no retry) if endpoint protection intervenes, a process is unexpectedly
terminated, files disappear, or access-denied errors recur; capture time + any threat id.
"""

from __future__ import annotations

import hashlib
import sys
from pathlib import Path

import win32gui
import win32process

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "prototypes" / "e6" / "engine"))

import excel_link  # E2/E6 tested attachment: Identity, attach, call, IdentityError, ExcelBusy
from openpyxl import load_workbook

from excel_toolkit.formula_audit import scan_worksheet, snapshot_from_com

TARGET_PATH = ROOT / "prototypes" / "integration" / "demo_workbook.xlsx"
TARGET_NAME = TARGET_PATH.name
SHEET = "Demo"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def find_and_attach():
    """Reuse E2's identity+attach to bind to the Excel instance that has the target workbook
    open. attach() verifies the window belongs to the pid, the workbook name is open and the
    file path matches - i.e. the intended process + workbook are confirmed before any read."""
    windows: list[int] = []
    win32gui.EnumWindows(lambda h, _: windows.append(h) or True, None)
    for hwnd in windows:
        if win32gui.GetClassName(hwnd) != "XLMAIN":
            continue
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        handoff = f"{hwnd}|{pid}|{TARGET_NAME}|{TARGET_PATH}"
        try:
            identity = excel_link.Identity.from_handoff(handoff)
            app, book = excel_link.attach(identity)
            return identity, app, book
        except (excel_link.IdentityError, excel_link.ExcelBusy):
            continue
    raise SystemExit(f"'{TARGET_NAME}' is not open in any Excel instance - open it and retry.")


def capture_state(app, ws):
    """Live in-memory formulas AND values of the used range (to prove the audit changed
    nothing in memory, not just on disk)."""
    formulas = excel_link.call(app, lambda: ws.UsedRange.Formula)
    values = excel_link.call(app, lambda: ws.UsedRange.Value)
    return formulas, values


def shape_report(app, ws):
    """Observe Excel's real .Formula return shapes to validate _as_grid's assumptions."""
    full = excel_link.call(app, lambda: ws.UsedRange.Formula)
    one_col = excel_link.call(app, lambda: ws.Range("K2:K6").Formula)
    one_cell = excel_link.call(app, lambda: ws.Range("K2").Formula)
    def shape(x):
        if isinstance(x, (list, tuple)):
            inner = x[0] if x else None
            return f"{type(x).__name__}[{len(x)}]" + (f" x {type(inner).__name__}[{len(inner)}]"
                                                      if isinstance(inner, (list, tuple)) else " (1-D)")
        return f"scalar({type(x).__name__})"
    return {"used_range": shape(full), "single_column": shape(one_col),
            "single_cell": shape(one_cell)}


def read_only_audit():
    saved_hash_before = _sha256(TARGET_PATH)
    identity, app, book = find_and_attach()

    # Verify worksheet identity before reading anything.
    ws = excel_link.call(app, lambda: book.Worksheets(SHEET))
    assert excel_link.call(app, lambda: ws.Name) == SHEET

    print("ATTACHED:")
    print(f"  pid={identity.pid} hwnd={identity.hwnd}")
    print(f"  workbook={excel_link.call(app, lambda: book.Name)!r}")
    print(f"  full_name={excel_link.call(app, lambda: book.FullName)!r}")
    print(f"  worksheet={SHEET!r}")
    print("COM .Formula shapes:", shape_report(app, ws))

    before = capture_state(app, ws)
    live = scan_worksheet(snapshot_from_com_via(app, ws))      # live, incl. unsaved edits
    saved = scan_worksheet(load_workbook(TARGET_PATH)[SHEET])  # saved file on disk
    after = capture_state(app, ws)

    live_f = sorted(f.coord for f in live.findings)
    saved_f = sorted(f.coord for f in saved.findings)
    delta = sorted(set(live_f) - set(saved_f))

    print(f"saved-file findings : {saved_f}")
    print(f"live findings       : {live_f}")
    print(f"unsaved-only (delta): {delta}   <- introduced by the unsaved edit")
    print(f"dependency coverage : {live.dependency_coverage:.0%}  pattern: {live.pattern_coverage:.0%}")
    print(f"in-memory integrity : {'UNCHANGED' if before == after else 'CHANGED (!)'}")
    print(f"saved-file integrity: {'UNCHANGED' if saved_hash_before == _sha256(TARGET_PATH) else 'CHANGED (!)'}")
    print("READ-ONLY run complete. Results workbook + navigation NOT performed (separate step).")
    return app, book


def snapshot_from_com_via(app, ws):
    """Call the tested bulk conversion under the busy-retry wrapper."""
    return excel_link.call(app, lambda: snapshot_from_com(ws))


def show_and_navigate(app, book, live_findings):  # pragma: no cover - SEPARATE step, not called
    """Step 2 (only after read-only passes): show findings in a NEW workbook and navigate to
    the first finding's cell in the ORIGINAL audited workbook. Left uncalled by design."""
    raise NotImplementedError("run only after the read-only checks pass and are reviewed")


if __name__ == "__main__":  # pragma: no cover - do not run until authorised
    read_only_audit()
