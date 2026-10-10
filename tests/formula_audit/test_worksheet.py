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


def test_scanning_does_not_materialise_empty_cells():
    # Reliability: scanning a sparse, large-dimension sheet must not create empty Cell
    # objects (which would grow the in-memory workbook).
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    for r in range(1, 5):
        ws[f"B{r}"] = r
    ws["C2"], ws["C3"], ws["C4"] = "=B2+1", "=B3+1", "=B4+1"
    ws["Z500"] = 1  # inflate dimensions to 500 x 26 while staying sparse
    before = len(ws._cells)
    report = scan_worksheet(ws)
    after = len(ws._cells)
    assert after == before, f"scan created {after - before} empty cells"
    assert report.total_formulas == 3


def test_unsupported_formula_in_series_not_counted_as_analysed():
    # Coverage precision: an INDIRECT cell inside an otherwise-valid series is reported
    # unsupported and must NOT inflate pattern coverage.
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws["B2"], ws["C2"] = "=B1*2", "=C1*2"
    ws["D2"] = '=INDIRECT("A1")*2'
    ws["E2"] = "=E1*2"
    report = scan_worksheet(ws)
    assert report.total_formulas == 4
    assert report.resolved == 3                 # dependency coverage excludes INDIRECT
    assert report.analysed_formulas == 3        # D2 not counted as comparable
    assert report.pattern_coverage == 3 / 4
    assert ("D2", "unsupported") in report.unsupported_cells
    assert report.findings == []                # D2 (no pattern) is not flagged


def test_cross_axis_conflict_is_reported_not_silently_resolved():
    # D3 is an interior outlier in BOTH its row (expects pattern A) and its column
    # (expects a different pattern B). The disagreement must be surfaced, not hidden.
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    for col in ("B", "C", "E"):
        ws[f"{col}1"] = 10
    for r in (2, 4, 5):
        ws[f"A{r}"] = r
    ws["D1"] = 1
    ws["B3"], ws["C3"], ws["E3"] = "=B$1*2", "=C$1*2", "=E$1*2"   # row pattern A = R1C*2
    ws["D2"], ws["D4"], ws["D5"] = "=$A2*3", "=$A4*3", "=$A5*3"   # column pattern B = RC1*3
    ws["D3"] = "=D1+D2"                                           # differs from both
    report = scan_worksheet(ws)
    d3 = next(f for f in report.findings if f.coord == "D3")
    assert set(d3.axes) == {"row", "column"}
    assert d3.conflict is True
    by_axis = dict(d3.expected_by_axis)
    assert by_axis["row"] != by_axis["column"]
    assert "disagree" in d3.message


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
