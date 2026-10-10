"""SCAFFOLDING ONLY - the Python side of the end-to-end demo. NOT executed here.

Identity contract (handoff from VBA): ``hwnd|workbook_name|workbook_full_path|sheet_name``.
The process id is NOT supplied; Python derives it from the window handle, exactly as the
proven E2 attachment does. This script would:

  1. attach to THAT Excel instance and worksheet (reuse E2's tested handoff/attach -
     prototypes/e2); do NOT invent a separate instance-finding method,
  2. build a live snapshot via the TESTED ``snapshot_from_com`` (bulk read, unsaved edits),
  3. run the SAME detector used offline (``scan_worksheet``),
  4. emit findings as structured JSON (``findings_to_json``) - robust to tabs/newlines and
     unusual workbook/sheet names; each record keeps the audited workbook identity so the UI
     navigates back to the ORIGINAL workbook.

Only steps 2-4 are exercised offline (see tests). Step 1 (live COM attach) is NOT run on this
machine - it belongs in the authorised test environment.
"""

from __future__ import annotations

import sys

from excel_toolkit.formula_audit import findings_to_json, scan_worksheet, snapshot_from_com


def parse_handoff(handoff: str) -> dict:
    hwnd, name, path, sheet = (handoff.split("|", 3) + ["", "", "", ""])[:4]
    return {"hwnd": hwnd, "workbook": name, "workbook_path": path, "sheet": sheet}


def attach_active_worksheet(handoff: dict):
    """Attach to the Excel instance/worksheet named by the handoff and return the worksheet
    COM object. Wired on the authorised machine by reusing E2's proven handoff/attach (PID
    derived from the window handle); intentionally not implemented here so nothing launches
    COM in this environment."""
    raise NotImplementedError("live COM attach is wired in the authorised test environment (reuse E2)")


def audit(handoff: str) -> str:
    ident = parse_handoff(handoff)
    com_ws = attach_active_worksheet(ident)                 # live Excel (authorised env only)
    report = scan_worksheet(snapshot_from_com(com_ws))       # same detector as offline
    return findings_to_json(report, ident["workbook"], ident["workbook_path"])


if __name__ == "__main__":  # pragma: no cover - not executed here
    print(audit(sys.argv[1] if len(sys.argv) > 1 else ""))
