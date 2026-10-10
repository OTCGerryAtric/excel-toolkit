"""Relative-reference (R1C1-style) normalisation of Excel formulas.

The audit idea: two cells that hold "the same formula filled across a range" have
*different* A1 text (``=B2*(1+B3)`` vs ``=C2*(1+C3)``) but the *same* structure once
each reference is expressed relative to the cell that holds it. Converting to an
R1C1-style form makes that structure explicit, so a formula that breaks its
neighbours' pattern gets a different normal form.

Absolute/mixed anchoring is preserved: ``$B$1`` normalises to a fixed ``R1C2`` (it does
not move when filled), while ``B1`` normalises to an offset like ``R[-3]C``. So a formula
that references an absolute cell is only "the same" as a neighbour that references the
*same* absolute cell — exactly Excel's fill semantics.

Coverage is reported honestly (mirroring E5): a formula is ``resolved`` only when every
reference was normalised. Run-time-dynamic constructs (``INDIRECT``/``OFFSET``), spill
ranges, defined names, tables and external links are reported as ``partial`` or
``unsupported`` and never silently treated as a clean, comparable pattern.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple

# Coverage statuses (same vocabulary as the E5 dependency work).
RESOLVED = "resolved"
PARTIAL = "partial"
UNSUPPORTED = "unsupported"

# A generic marker for a cell that holds a constant rather than a formula. All constants
# share one marker: structurally "a hard-coded value", regardless of the number itself.
CONST = "#CONST#"

# Functions whose real precedent is only known at run time, so the written formula is not
# a complete structural pattern (see docs/validation/phase0-e5.md).
DYNAMIC_FUNCS = {"INDIRECT", "OFFSET"}

_SINGLE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$")
_SHEET_SPLIT = re.compile(r"^((?:'[^']*'|[^'!]+)!)(.*)$")


@dataclass(frozen=True)
class NormalResult:
    """Outcome of normalising one cell's formula.

    pattern:  canonical R1C1 form, or ``None`` when coverage is not ``resolved``
              (``None`` means "do not use this as a comparable pattern").
    coverage: ``resolved`` | ``partial`` | ``unsupported``.
    notes:    human-readable reasons, e.g. why a formula was unsupported.
    """

    pattern: str | None
    coverage: str
    notes: tuple[str, ...] = ()


def _ref_to_r1c1(col_abs: str, col: str, row_abs: str, row: str,
                 anchor_row: int, anchor_col: int) -> str:
    col_idx = column_index_from_string(col)
    row_idx = int(row)
    if col_abs:
        col_part = f"C{col_idx}"
    else:
        delta = col_idx - anchor_col
        col_part = "C" if delta == 0 else f"C[{delta}]"
    if row_abs:
        row_part = f"R{row_idx}"
    else:
        delta = row_idx - anchor_row
        row_part = "R" if delta == 0 else f"R[{delta}]"
    return row_part + col_part


def _normalize_range_operand(value: str, anchor_row: int, anchor_col: int) -> tuple[str | None, str]:
    """Normalise one RANGE operand. Returns (normalised_or_None, coverage).

    Note: a spill reference (``D1#``) never reaches here — openpyxl's tokenizer cannot
    parse the ``#`` operator and raises, so such formulas are reported ``unsupported`` by
    ``normalize_formula``. (A future tokenizer could classify spill as ``partial``.)
    """
    sheet_prefix = ""
    rest = value
    m = _SHEET_SPLIT.match(value)
    if m:
        sheet_prefix, rest = m.group(1), m.group(2)
    if "[" in value or "]" in value:
        return None, UNSUPPORTED  # external workbook / structured table reference
    parts = rest.split(":")
    normalised = []
    for part in parts:
        sm = _SINGLE.match(part)
        if not sm:
            return None, UNSUPPORTED  # defined name or something we don't parse
        normalised.append(_ref_to_r1c1(*sm.groups(), anchor_row, anchor_col))
    return sheet_prefix + ":".join(normalised), RESOLVED


def normalize_formula(formula: str, anchor: str) -> NormalResult:
    """Normalise a formula string located at cell ``anchor`` (e.g. ``"D4"``)."""
    anchor_row, anchor_col = coordinate_to_tuple(anchor)
    try:
        tokens = Tokenizer(formula).items
    except Exception as err:  # noqa: BLE001 - tokenizer raises bare exceptions
        return NormalResult(None, UNSUPPORTED, (f"could not parse: {err}",))

    coverage = RESOLVED
    notes: list[str] = []
    parts: list[str] = []
    usable = True

    for tok in tokens:
        if tok.type == "WHITE-SPACE":
            continue
        if tok.type == "FUNC":
            name = tok.value.rstrip("(").upper()
            if name in DYNAMIC_FUNCS:
                coverage = UNSUPPORTED
                usable = False
                notes.append(f"{name} target is determined at run time")
            parts.append(name + ("(" if tok.value.endswith("(") else ""))
            continue
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            norm, status = _normalize_range_operand(tok.value, anchor_row, anchor_col)
            if norm is None:
                usable = False
                coverage = UNSUPPORTED if status == UNSUPPORTED else (
                    PARTIAL if coverage == RESOLVED else coverage)
                notes.append(f"unresolved reference {tok.value!r} ({status})")
                parts.append(tok.value)
            else:
                parts.append(norm)
            continue
        # numbers, text, logicals, errors, operators, parens, separators: keep verbatim
        parts.append(tok.value)

    pattern = "".join(parts) if usable else None
    return NormalResult(pattern, coverage, tuple(notes))


def normalize_cell_value(value, anchor: str) -> NormalResult:
    """Normalise whatever a cell holds: a formula, a constant, or nothing."""
    if isinstance(value, str) and value.startswith("="):
        return normalize_formula(value, anchor)
    if value is None or value == "":
        return NormalResult(None, UNSUPPORTED, ("empty cell",))
    return NormalResult(CONST, RESOLVED, ("constant",))
