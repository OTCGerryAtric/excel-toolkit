"""Offline tests for the integration transport and navigation-target helpers."""

import json

from excel_toolkit.formula_audit import (
    InMemoryWorksheet,
    a1_sheet_reference,
    findings_payload,
    findings_to_json,
    navigation_target,
    scan_worksheet,
)
from excel_toolkit.formula_audit.worksheet import WorksheetFinding, WorksheetReport


def test_a1_reference_quotes_and_escapes_apostrophes():
    assert a1_sheet_reference("Sheet1", "B2") == "'Sheet1'!B2"
    assert a1_sheet_reference("Q1 2026", "B2") == "'Q1 2026'!B2"
    assert a1_sheet_reference("P&L", "B2") == "'P&L'!B2"
    assert a1_sheet_reference("O'Brien", "B2") == "'O''Brien'!B2"


def test_navigation_target_preserves_identity():
    t = navigation_target("Model.xlsx", "P&L 'A'", "C5")
    assert t["workbook"] == "Model.xlsx"
    assert t["sheet"] == "P&L 'A'"
    assert t["cell"] == "C5"
    assert t["reference"] == "'P&L ''A'''!C5"


def test_findings_payload_keeps_workbook_identity():
    snap = InMemoryWorksheet(
        {"B1": 1, "C1": 1, "D1": 1, "E1": 1,
         "B2": "=B1*2", "C2": "=C1*2", "D2": "=D1*3", "E2": "=E1*2"}, "Sheet1")
    report = scan_worksheet(snap)
    payload = findings_payload(report, "Book.xlsx", r"C:\models\Book.xlsx")
    assert payload and all(p["workbook"] == "Book.xlsx" for p in payload)
    assert all(p["workbook_path"] == r"C:\models\Book.xlsx" for p in payload)
    assert all(p["sheet"] == "Sheet1" for p in payload)
    assert {p["cell"] for p in payload} == {"D2"}


def test_transport_json_robust_to_tabs_newlines_and_odd_names():
    # Hand-built report with hostile strings in the message and odd identity names.
    nasty_msg = "line1\tcol\nline2 with \"quotes\" and 'apostrophes'"
    finding = WorksheetFinding(
        coord="B2", row=2, col=2, axes=("row",), observed="RC[-1]*3", expected="RC[-1]*2",
        expected_by_axis=(("row", "RC[-1]*2"),), conflict=False, kind="singleton",
        block_by_axis=(), confidence="high", message=nasty_msg)
    report = WorksheetReport(sheet="Weird\tSheet\nName")
    report.findings = [finding]

    text = findings_to_json(report, 'Book "2026".xlsx', r"C:\a\b.xlsx")
    restored = json.loads(text)                     # must be valid JSON
    assert len(restored) == 1
    rec = restored[0]
    assert rec["message"] == nasty_msg              # tabs/newlines/quotes survive intact
    assert rec["sheet"] == "Weird\tSheet\nName"
    assert rec["workbook"] == 'Book "2026".xlsx'
    assert rec["cell"] == "B2" and rec["kind"] == "singleton"


def test_transport_json_valid_for_real_report():
    snap = InMemoryWorksheet(
        {"B1": 1, "C1": 1, "D1": 1, "E1": 1,
         "B2": "=B1*2", "C2": "=C1*2", "D2": "=D1*3", "E2": "=E1*2"}, "S")
    text = findings_to_json(scan_worksheet(snap), "S.xlsx", "S.xlsx")
    data = json.loads(text)
    assert isinstance(data, list) and data[0]["cell"] == "D2"
