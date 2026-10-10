"""Series segmentation tested independently of anomaly detection."""

from excel_toolkit.formula_audit import segment_line


def test_splits_on_text_and_blank_boundaries():
    values = ["=A1", "=A2", "Heading", "=B1", None, 5, "=C1", 7, 8]
    assert segment_line(values) == [
        (0, 1, "formula"),
        (3, 3, "formula"),
        (5, 5, "const"),
        (6, 6, "formula"),
        (7, 8, "const"),
    ]


def test_splits_const_to_formula_transition():
    # actuals (numbers) meeting forecast (formulas) is a boundary.
    values = [100, 120, "=B1*2", "=C1*2"]
    assert segment_line(values) == [(0, 1, "const"), (2, 3, "formula")]


def test_all_blank_or_text_yields_nothing():
    assert segment_line([None, "", "Label", None]) == []


def test_dates_and_text_are_boundaries():
    import datetime
    values = ["=A1", datetime.datetime(2026, 1, 1), "=A3"]
    assert segment_line(values) == [(0, 0, "formula"), (2, 2, "formula")]


def test_single_trailing_formula_run():
    values = ["Label", "=A1", "=A2", "=A3"]
    assert segment_line(values) == [(1, 3, "formula")]
