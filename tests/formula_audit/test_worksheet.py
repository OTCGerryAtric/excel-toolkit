"""Worksheet auto-scan: per-sheet behaviour, axes, dedup/order, coverage, and metrics."""

import pytest
from openpyxl import load_workbook

from excel_toolkit.formula_audit import scan_worksheet

from .fixtures_ws import SHEET_LABELS, build_ws_fixtures


@pytest.fixture(scope="module")
def wb(tmp_path_factory):
    path = build_ws_fixtures(tmp_path_factory.mktemp("ws") / "corpus.xlsx")
    return load_workbook(path)


@pytest.mark.parametrize("sheet", list(SHEET_LABELS))
def test_flags_expected_and_no_false_positives(wb, sheet):
    expected_tp, known_fn, _axes = SHEET_LABELS[sheet]
    report = scan_worksheet(wb[sheet])
    flagged = {f.coord for f in report.findings}
    anomalies = expected_tp | known_fn
    assert expected_tp <= flagged, f"{sheet}: missed {expected_tp - flagged}"
    assert flagged & known_fn == set(), f"{sheet}: unexpectedly caught known FN {flagged & known_fn}"
    assert flagged - anomalies == set(), f"{sheet}: false positives {flagged - anomalies}"


def test_axes_recorded(wb):
    for sheet, (_tp, _fn, axes) in SHEET_LABELS.items():
        report = scan_worksheet(wb[sheet])
        by_coord = {f.coord: f for f in report.findings}
        for coord, expected_axes in axes.items():
            assert by_coord[coord].axes == expected_axes, (
                f"{sheet}!{coord}: axes {by_coord[coord].axes} != {expected_axes}")


def test_findings_deterministic_order_and_deduped(wb):
    report = scan_worksheet(wb["overlap"])
    coords = [f.coord for f in report.findings]
    assert coords == sorted(coords, key=lambda c: (int(c[1:]), c[0]))  # (row, col)
    assert len(coords) == len(set(coords))  # deduped across row + column scans


def test_findings_are_potential_not_certain(wb):
    report = scan_worksheet(wb["revenue_grid"])
    assert report.findings
    for f in report.findings:
        assert "Potential inconsistency" in f.message and "error" not in f.message.lower()
        assert f.confidence == "high"


def test_coverage_axes_reported_separately(wb):
    report = scan_worksheet(wb["revenue_grid"])
    # dependency coverage (references resolved) and pattern coverage (had series context)
    # are distinct numbers and both populated.
    assert 0.0 <= report.pattern_coverage <= report.dependency_coverage <= 1.0
    assert report.dependency_coverage == 1.0  # all references resolve in this sheet
    assert report.total_formulas > 0


def test_totals_headings_no_false_positives(wb):
    report = scan_worksheet(wb["totals_headings"])
    assert report.findings == []            # headings/totals/blanks cause no false positives


def test_isolated_formula_is_skipped_not_flagged():
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws["C3"] = "=A3+1"   # a lone formula: no series context in either direction
    report = scan_worksheet(ws)
    assert report.findings == []
    assert report.total_formulas == 1
    assert report.analysed_formulas == 0
    assert ("C3", "insufficient series context") in report.skipped_cells


def test_metrics_precision_recall(wb, capsys):
    tp = fp = fn = 0
    resolved = total = analysed = 0
    for sheet, (expected_tp, known_fn, _axes) in SHEET_LABELS.items():
        report = scan_worksheet(wb[sheet])
        flagged = {f.coord for f in report.findings}
        anomalies = expected_tp | known_fn
        tp += len(flagged & anomalies)
        fp += len(flagged - anomalies)
        fn += len(anomalies - flagged)
        resolved += report.resolved
        total += report.total_formulas
        analysed += report.analysed_formulas

    precision = tp / (tp + fp) if (tp + fp) else 1.0
    recall = tp / (tp + fn) if (tp + fn) else 1.0
    dep_cov = resolved / total if total else 0.0
    pat_cov = analysed / total if total else 0.0
    with capsys.disabled():
        print(f"\n  corpus formula cells  : {total}")
        print(f"  true positives        : {tp}")
        print(f"  false positives       : {fp}")
        print(f"  false negatives       : {fn}  (incl. known consecutive-error FN)")
        print(f"  precision             : {precision:.0%}")
        print(f"  recall                : {recall:.0%}")
        print(f"  dependency coverage   : {dep_cov:.0%}")
        print(f"  pattern coverage      : {pat_cov:.0%}")

    assert fp == 0            # the hard requirement: no legitimate formula flagged
    assert precision == 1.0
    assert total > 26         # corpus substantially larger than Gate 1
