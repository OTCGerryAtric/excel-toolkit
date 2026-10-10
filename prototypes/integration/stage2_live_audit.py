"""STAGE 2 - first live-Excel test. DO NOT RUN until authorised.

Preconditions (all required before running):
  * you have confirmed this kind of Python/Excel-COM automation is permitted on this laptop;
  * prototypes/integration/demo_workbook.xlsx is OPEN in Excel;
  * you have edited D3 on sheet 'Demo' from  =C3*(1+$A$10)  to  =C3*(1+$A$11)  and have
    NOT saved (this is the deliberate unsaved anomaly);
  * run with the scratch venv that has pywin32 + openpyxl.

First execution is READ-ONLY: it attaches, verifies identity, reads live formulas, runs the
detector, validates single-column/single-cell COM conversion, and **asserts** the expected
evidence (saved {D4}; live {D3, D4}; unsaved-only delta {D3}; workbook unchanged in memory
and on disk) - failing loudly if any is wrong, so it can never announce success on a wrong
result. It does NOT create a results workbook or navigate - that is a separate step
(show_and_navigate, defined but intentionally not called).

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
from excel_toolkit.formula_audit.worksheet_source import _as_grid

TARGET_PATH = ROOT / "prototypes" / "integration" / "demo_workbook.xlsx"
TARGET_NAME = TARGET_PATH.name
SHEET = "Demo"

# Expected evidence (the run FAILS if any differs).
EXPECTED_SAVED = {"D4"}             # planted anomaly in the saved file
EXPECTED_LIVE = {"D3", "D4"}        # with the unsaved D3 edit applied
EXPECTED_DELTA = {"D3"}             # introduced only by the unsaved edit
# Column K roll-forward, to validate single-column COM conversion lands vertically in K3:K6.
EXPECTED_K = {3: "=K2+$A$10", 4: "=K3+$A$10", 5: "=K4+$A$10", 6: "=K5+$A$10"}


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

    # Single-column and single-cell COM reads -> validate _as_grid orientation + coordinates.
    col_grid = _as_grid(excel_link.call(app, lambda: ws.Range("K2:K6").Formula))
    cell_grid = _as_grid(excel_link.call(app, lambda: ws.Range("K2").Formula))

    snap = snapshot_from_com_via(app, ws)                      # live, incl. unsaved edits
    live = scan_worksheet(snap)
    saved = scan_worksheet(load_workbook(TARGET_PATH)[SHEET])  # saved file on disk
    after = capture_state(app, ws)

    live_f = {f.coord for f in live.findings}
    saved_f = {f.coord for f in saved.findings}
    delta = live_f - saved_f

    # Single-column K2:K6 must come back as 5 rows x 1 col (NOT transposed to one row),
    # with the roll-forward formulas landing in K3:K6 (rows 1..4 of the block).
    col_vertical = len(col_grid) == 5 and all(len(r) == 1 for r in col_grid)
    col_formulas_ok = col_vertical and all(col_grid[i][0] == EXPECTED_K[i + 2] for i in range(1, 5))
    cell_ok = len(cell_grid) == 1 and len(cell_grid[0]) == 1  # single cell -> 1x1
    # Used-range snapshot must also position column K (col 11) correctly.
    snap_k_ok = all(snap.value_at(r, 11) == f for r, f in EXPECTED_K.items())

    saved_cov = scan_worksheet(load_workbook(TARGET_PATH)[SHEET]).pattern_coverage
    checks = [
        ("saved-file findings == {D4}", saved_f == EXPECTED_SAVED, sorted(saved_f)),
        ("live findings == {D3,D4}", live_f == EXPECTED_LIVE, sorted(live_f)),
        ("unsaved-only delta == {D3}", delta == EXPECTED_DELTA, sorted(delta)),
        ("in-memory contents unchanged", before == after, "unchanged" if before == after else "CHANGED"),
        ("saved file unchanged", saved_hash_before == _sha256(TARGET_PATH), "unchanged"),
        ("single-column K2:K6 vertical (no transpose)", col_vertical, f"{len(col_grid)} rows"),
        ("single-column K3:K6 formulas correctly positioned", col_formulas_ok, col_grid),
        ("single-cell K2 -> 1x1", cell_ok, cell_grid),
        ("used-range snapshot positions column K", snap_k_ok,
         {r: snap.value_at(r, 11) for r in EXPECTED_K}),
        ("saved/live pattern-coverage parity", live.pattern_coverage == saved_cov,
         (round(saved_cov, 3), round(live.pattern_coverage, 3))),
    ]
    for name, ok, detail in checks:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}: {detail}")
    print(f"  dependency coverage : {live.dependency_coverage:.0%}  pattern: {live.pattern_coverage:.0%}")

    if not all(ok for _n, ok, _d in checks):
        raise SystemExit("STAGE 2 READ-ONLY: FAIL — expected live-Excel evidence not met "
                         "(check that D3 was edited to =C3*(1+$A$11) and NOT saved).")
    print("STAGE 2 READ-ONLY: PASS.")
    return app, book, ws, live, before


def snapshot_from_com_via(app, ws):
    """Call the tested bulk conversion under the busy-retry wrapper."""
    return excel_link.call(app, lambda: snapshot_from_com(ws))


def show_and_navigate(app, book, ws, live, before):
    """Display findings in a NEW results workbook (audited workbook never modified) and
    navigate to the first finding's cell in the ORIGINAL audited workbook."""
    audited_name = excel_link.call(app, lambda: book.Name)
    results = excel_link.call(app, lambda: app.Workbooks.Add())
    rws = excel_link.call(app, lambda: results.Worksheets(1))
    headers = ["Workbook", "Sheet", "Cell", "Confidence", "Potential inconsistency (review)"]
    for c, h in enumerate(headers, start=1):
        excel_link.call(app, lambda c=c, h=h: rws.Cells(1, c).__setattr__("Value", h))
    for i, f in enumerate(live.findings, start=2):
        row = [audited_name, live.sheet, f.coord, f.confidence, f.message]
        for c, val in enumerate(row, start=1):
            excel_link.call(app, lambda c=c, val=val, i=i: rws.Cells(i, c).__setattr__("Value", val))

    # Navigate to the first finding's cell IN THE ORIGINAL audited workbook (object model).
    first = live.findings[0]
    target = excel_link.call(app, lambda: book.Worksheets(live.sheet).Range(first.coord))
    excel_link.call(app, lambda: app.Goto(target, True))
    print(f"results workbook created ({len(live.findings)} findings); navigated to "
          f"{audited_name}!{live.sheet}!{first.coord}")

    # The audited workbook's formulas AND values must be unchanged by display + navigation.
    after_nav = capture_state(app, ws)
    if before != after_nav:
        raise SystemExit("STAGE 2: FAIL — audited workbook changed during display/navigation.")
    print("audited workbook unchanged after display + navigation (PASS). Results workbook left"
          " open and unsaved.")


def run(full: bool = False):
    app, book, ws, live, before = read_only_audit()
    if full:
        show_and_navigate(app, book, ws, live, before)


if __name__ == "__main__":  # pragma: no cover - do not run until authorised
    run(full=(len(sys.argv) > 1 and sys.argv[1] == "full"))
