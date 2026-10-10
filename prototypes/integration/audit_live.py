"""SCAFFOLDING ONLY - the Python side of the end-to-end demo. NOT executed here.

Contract: VBA (ExcelToolkitAudit.bas) passes the explicit handoff
``hwnd|pid|workbook_name|workbook_path`` (the proven E2/E4 identity). This script would:

  1. attach to THAT Excel instance (reuse E2's handoff/attach - prototypes/e2),
  2. take the active worksheet COM object,
  3. build a live snapshot via the TESTED ``snapshot_from_com`` (reads unsaved formulas),
  4. run the SAME detector used offline (``scan_worksheet``),
  5. print one finding per line for VBA: sheet<TAB>cell<TAB>kind<TAB>confidence<TAB>message.

Only steps 3-5 are exercised offline (see tests/formula_audit/test_snapshot.py, which drives
``snapshot_from_com`` with a fake COM object). Step 1 (live COM attach) is NOT run on this
machine - it belongs in the authorised test environment.
"""

from __future__ import annotations

import sys

from excel_toolkit.formula_audit import scan_worksheet, snapshot_from_com


def attach_active_worksheet(handoff: str):
    """Attach to the Excel instance named by the handoff and return the active worksheet
    COM object. To be wired on the authorised machine by reusing E2's proven handoff/attach;
    intentionally not implemented here so nothing launches COM in this environment."""
    raise NotImplementedError("live COM attach is wired in the authorised test environment (reuse E2)")


def audit(handoff: str) -> str:
    com_ws = attach_active_worksheet(handoff)          # live Excel (authorised env only)
    report = scan_worksheet(snapshot_from_com(com_ws))  # same detector as offline
    lines = [f"{report.sheet}\t{f.coord}\t{f.kind}\t{f.confidence}\t{f.message}"
             for f in report.findings]
    return "\n".join(lines)


if __name__ == "__main__":  # pragma: no cover - not executed here
    print(audit(sys.argv[1] if len(sys.argv) > 1 else ""))
