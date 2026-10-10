"""Unit tests for R1C1 relative-reference normalisation, incl. absolute/mixed anchoring."""

from excel_toolkit.formula_audit import (
    RESOLVED,
    UNSUPPORTED,
    normalize_formula,
)


def test_filled_formula_has_same_pattern_across_columns():
    # =B2*(1+B3) at B4 and =C2*(1+C3) at C4 are "the same filled formula".
    a = normalize_formula("=B2*(1+B3)", "B4")
    b = normalize_formula("=C2*(1+C3)", "C4")
    assert a.coverage == RESOLVED and b.coverage == RESOLVED
    assert a.pattern == b.pattern == "R[-2]C*(1+R[-1]C)"


def test_wrong_row_breaks_the_pattern():
    good = normalize_formula("=D2*(1+D3)", "D4")
    bad = normalize_formula("=D2*(1+D4)", "D4")
    assert good.pattern != bad.pattern
    assert bad.pattern == "R[-2]C*(1+RC)"


def test_absolute_reference_is_fixed_not_offset():
    # $B$1 is the same cell wherever it is filled; two cells referencing it match.
    a = normalize_formula("=B3*(1+$B$1)", "B4")
    b = normalize_formula("=C3*(1+$B$1)", "C4")
    assert a.pattern == b.pattern == "R[-1]C*(1+R1C2)"
    # A different absolute target must produce a different pattern.
    c = normalize_formula("=D3*(1+$C$1)", "D4")
    assert c.pattern != a.pattern
    assert c.pattern == "R[-1]C*(1+R1C3)"


def test_mixed_reference_anchoring():
    # $B1 fixes the column, B$1 fixes the row.
    col_fixed = normalize_formula("=$B1", "C4")   # column absolute, row relative
    row_fixed = normalize_formula("=B$1", "C4")   # column relative, row absolute
    assert col_fixed.pattern == "R[-3]C2"
    assert row_fixed.pattern == "R1C[-1]"


def test_range_reference_normalised():
    r = normalize_formula("=SUM(B4:E4)", "F4")
    assert r.coverage == RESOLVED
    assert r.pattern == "SUM(RC[-4]:RC[-1])"


def test_indirect_is_unsupported_not_clean():
    r = normalize_formula('=INDIRECT("B2")*(1+C3)', "C4")
    assert r.coverage == UNSUPPORTED
    assert r.pattern is None
    assert any("INDIRECT" in n for n in r.notes)


def test_external_reference_unsupported():
    r = normalize_formula("=[Book2.xlsx]Sheet1!A1", "B2")
    assert r.coverage == UNSUPPORTED
    assert r.pattern is None


def test_spill_reference_unsupported_tokenizer_limit():
    # openpyxl's tokenizer cannot parse the spill operator '#', so such formulas are
    # reported unsupported (a known limitation), never silently treated as clean.
    r = normalize_formula("=SUM(D1#)", "A1")
    assert r.coverage == UNSUPPORTED
    assert r.pattern is None


def test_unparseable_formula_is_unsupported():
    r = normalize_formula("=this is not )( valid", "A1")
    assert r.coverage == UNSUPPORTED
    assert r.pattern is None
