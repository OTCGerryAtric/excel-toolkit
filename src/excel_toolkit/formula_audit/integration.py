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


# -- Display-oriented transport (flat) -----------------------------------------------------
# The full findings_payload carries a nested ``expected_by_axis``. The Excel front end only
# needs flat, scalar fields to render a results row and navigate back, and a *flat* array of
# string/bool values is far simpler (and safer) for the VBA JSON reader to parse than nested
# arrays. These are a thin projection of the same findings — no new detection, no rule change.

DISPLAY_FIELDS = ("workbook", "workbook_path", "sheet", "cell", "confidence", "kind",
                  "conflict", "message")


def display_records(report, workbook: str, workbook_path: str) -> list[dict]:
    """Flat per-finding records for the results table. Each keeps the audited workbook name
    and path plus sheet/cell, so a double-click can navigate to the ORIGINAL workbook."""
    return [
        {
            "workbook": workbook,
            "workbook_path": workbook_path,
            "sheet": report.sheet,
            "cell": f.coord,
            "confidence": f.confidence,
            "kind": f.kind,
            "conflict": f.conflict,
            "message": f.message,
        }
        for f in report.findings
    ]


def display_json(report, workbook: str, workbook_path: str) -> str:
    """Flat findings JSON for the COM callback: an array of objects whose values are strings
    or booleans only (no nested arrays), so tabs/newlines/quotes survive as JSON-escaped
    string values and the VBA reader stays small."""
    return json.dumps(display_records(report, workbook, workbook_path), ensure_ascii=False)


def navigation_from_record(record: dict) -> dict:
    """The navigation descriptor a results-row double-click produces: read the audited
    workbook/sheet/cell off the record and target the ORIGINAL workbook (never the results
    book). Mirrors what ``PhaseBBridge`` does from the hidden columns of a results row."""
    return navigation_target(record["workbook"], record["sheet"], record["cell"])
