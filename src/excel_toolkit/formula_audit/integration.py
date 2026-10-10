"""Integration helpers shared by the Excel front end: a robust, structured findings
transport (JSON) and navigation-target construction that preserves workbook/worksheet
identity. Pure functions — no Excel, no COM — so they are fully testable offline.
"""

from __future__ import annotations

import json


def a1_sheet_reference(sheet: str, cell: str) -> str:
    """A fully-qualified A1 reference with the sheet name always quoted and embedded
    apostrophes doubled, e.g. ``O'Brien`` → ``'O''Brien'!B2``. Always quoting is safe for
    names with spaces, ``&`` etc.; VBA navigation resolves identity via the object model, so
    this is for display/logging rather than parsing."""
    return f"'{sheet.replace(chr(39), chr(39) * 2)}'!{cell}"


def navigation_target(workbook: str, sheet: str, cell: str) -> dict:
    """Explicit navigation descriptor: the audited workbook, worksheet and cell are kept
    distinct so navigation targets the ORIGINAL workbook, never a results workbook."""
    return {"workbook": workbook, "sheet": sheet, "cell": cell,
            "reference": a1_sheet_reference(sheet, cell)}


def findings_payload(report, workbook: str, workbook_path: str) -> list[dict]:
    """Structured records for transport to the front end. Each finding keeps the audited
    workbook identity (name + path) plus the sheet/cell, so the UI can navigate back to the
    right workbook regardless of where results are displayed."""
    return [
        {
            "workbook": workbook,
            "workbook_path": workbook_path,
            "sheet": report.sheet,
            "cell": f.coord,
            "kind": f.kind,
            "confidence": f.confidence,
            "conflict": f.conflict,
            "expected_by_axis": [[axis, pattern] for axis, pattern in f.expected_by_axis],
            "message": f.message,
        }
        for f in report.findings
    ]


def findings_to_json(report, workbook: str, workbook_path: str) -> str:
    """Serialise findings as JSON — robust to tabs, newlines and unusual workbook/sheet
    names (everything is a JSON string value, nothing delimiter-separated)."""
    return json.dumps(findings_payload(report, workbook, workbook_path), ensure_ascii=False)
