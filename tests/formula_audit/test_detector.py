"""Detector tests over offline workbook fixtures.

Asserts the genuine mistakes are caught, the legitimate variations are NOT flagged
(false-positive rate measured and required to be zero), coverage is reported, and
unsupported constructs are surfaced rather than treated as clean.
"""

import pytest
from openpyxl import load_workbook

from excel_toolkit.formula_audit import UNSUPPORTED, scan_series, series_from_row

from .fixtures import SCENARIOS, build_fixtures


@pytest.fixture(scope="module")
def workbook(tmp_path_factory):
    path = build_fixtures(tmp_path_factory.mktemp("fixtures") / "inconsistency.xlsx")
    return load_workbook(path)  # data_only=False -> formula strings preserved


def _scan(workbook, scenario):
    sheet, row, (c0, c1), mistakes, legit, unsupported = SCENARIOS[scenario]
    ws = workbook[sheet]
    report = scan_series(series_from_row(ws, row, c0, c1))
    return report, mistakes, legit, unsupported


@pytest.mark.parametrize("scenario", list(SCENARIOS))
def test_scenario_flags_exactly_the_mistakes(workbook, scenario):
    report, mistakes, legit, _ = _scan(workbook, scenario)
    flagged = {f.coord for f in report.findings}
    assert flagged == mistakes, f"{scenario}: flagged {flagged}, expected {mistakes}"
    # No legitimate cell may be flagged.
    assert flagged.isdisjoint(legit), f"{scenario}: false positive on {flagged & legit}"


def test_findings_are_potential_not_certain(workbook):
    report, _, _, _ = _scan(workbook, "march")
    assert report.findings
    for f in report.findings:
        assert "Potential inconsistency" in f.message
        assert "error" not in f.message.lower()  # never asserted as a definite error
        assert f.confidence == "high"


def test_unsupported_reported_not_treated_as_clean(workbook):
    report, _, _, unsupported = _scan(workbook, "indirect")
    reported = {c for c, _status in report.unsupported_cells}
    assert reported == unsupported
    assert UNSUPPORTED in {status for _c, status in report.unsupported_cells}
    # the INDIRECT cell is neither flagged nor counted resolved
    assert report.resolved == report.total - len(unsupported)


def test_small_series_does_not_flag_the_majority():
    # Only two cells: no dominant pattern of >=3, so nothing is flagged.
    from excel_toolkit.formula_audit import CellFormula
    cells = [CellFormula("A1", "=B1"), CellFormula("A2", "=ZZ9")]
    assert scan_series(cells).findings == []


def test_metrics_zero_false_positives_full_recall(workbook, capsys):
    total_mistakes = total_found = 0
    total_legit = total_false_pos = 0
    resolved = total_cells = 0
    for scenario in SCENARIOS:
        report, mistakes, legit, _ = _scan(workbook, scenario)
        flagged = {f.coord for f in report.findings}
        total_mistakes += len(mistakes)
        total_found += len(flagged & mistakes)
        total_legit += len(legit)
        total_false_pos += len(flagged & legit)
        resolved += report.resolved
        total_cells += report.total

    recall = total_found / total_mistakes if total_mistakes else 1.0
    fp_rate = total_false_pos / total_legit if total_legit else 0.0
    coverage = resolved / total_cells if total_cells else 0.0

    with capsys.disabled():
        print(f"\n  detection recall      : {total_found}/{total_mistakes} = {recall:.0%}")
        print(f"  false-positive rate   : {total_false_pos}/{total_legit} = {fp_rate:.0%}")
        print(f"  formula coverage      : {resolved}/{total_cells} = {coverage:.0%}")

    assert recall == 1.0
    assert fp_rate == 0.0
