"""Detect a formula that breaks the pattern established by its neighbours.

Conservative, high-confidence rule (unchanged since Gate 1): a cell is flagged only when
it is an **interior singleton** whose normalised pattern differs from a **dominant**
pattern that its immediate left and right neighbours both share. Edge cells and
block boundaries (e.g. actuals → forecast, year-end totals) are never flagged.

Two axes are reported and kept separate:
- **coverage** (per cell): could the formula's references be resolved? resolved / partial
  / unsupported — dependency-resolution coverage.
- **confidence** (per finding): only ``high`` is emitted here (flanked interior singleton).
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .normalize import PARTIAL, RESOLVED, UNSUPPORTED, NormalResult, normalize_cell_value

FLANKED_MESSAGE = (
    "Formula differs from the consistent pattern of the formulas on both sides. "
    "Potential inconsistency — review whether the difference is intentional."
)


@dataclass(frozen=True)
class CellFormula:
    coord: str
    raw: object  # formula string, constant, or None


@dataclass(frozen=True)
class Finding:
    coord: str
    observed: str
    expected: str
    confidence: str  # "high" in this increment
    message: str


@dataclass
class SeriesReport:
    total: int = 0
    resolved: int = 0
    partial: int = 0
    unsupported: int = 0
    dominant: str | None = None
    dominant_count: int = 0
    findings: list[Finding] = field(default_factory=list)
    unsupported_cells: list[tuple[str, str]] = field(default_factory=list)

    @property
    def coverage_rate(self) -> float:
        return self.resolved / self.total if self.total else 0.0


def dominant_pattern(patterns: list[str | None]) -> tuple[str | None, int]:
    """Most common resolved pattern and its count (ties broken by first occurrence)."""
    counts = Counter(p for p in patterns if p is not None)
    if not counts:
        return None, 0
    return counts.most_common(1)[0]


def flanked_outliers(coords: list[str], patterns: list[str | None], dominant: str) -> list[Finding]:
    """Interior cells whose pattern differs from ``dominant`` while both neighbours match it.

    Shared by single-series and worksheet scanning so the rule is identical everywhere.
    """
    findings = []
    for i in range(1, len(coords) - 1):  # interior only; never the first or last cell
        p = patterns[i]
        if p is None or p == dominant:
            continue
        if patterns[i - 1] == dominant and patterns[i + 1] == dominant:
            findings.append(Finding(coords[i], p, dominant, "high", FLANKED_MESSAGE))
    return findings


def scan_series(cells: list[CellFormula], *, min_group: int = 3) -> SeriesReport:
    """Scan one ordered 1-D run of cells (a row or column span) for a lone outlier.

    ``min_group``: the dominant pattern must appear at least this many times before any
    outlier is flagged, so we never "flag the majority" in a tiny or noisy series.
    """
    report = SeriesReport(total=len(cells))
    normals: list[NormalResult] = [normalize_cell_value(c.raw, c.coord) for c in cells]
    for cell, nr in zip(cells, normals):
        if nr.coverage == RESOLVED:
            report.resolved += 1
        elif nr.coverage == PARTIAL:
            report.partial += 1
            report.unsupported_cells.append((cell.coord, PARTIAL))
        else:
            report.unsupported += 1
            report.unsupported_cells.append((cell.coord, UNSUPPORTED))

    patterns = [nr.pattern for nr in normals]
    report.dominant, report.dominant_count = dominant_pattern(patterns)
    if report.dominant is not None and report.dominant_count >= min_group:
        coords = [c.coord for c in cells]
        report.findings = flanked_outliers(coords, patterns, report.dominant)
    return report


def series_from_row(ws, row: int, start_col: int, end_col: int) -> list[CellFormula]:
    """Build an ordered left-to-right series from one worksheet row span (inclusive)."""
    from openpyxl.utils.cell import get_column_letter
    return [CellFormula(f"{get_column_letter(col)}{row}", ws[f"{get_column_letter(col)}{row}"].value)
            for col in range(start_col, end_col + 1)]


def series_from_column(ws, col: int, start_row: int, end_row: int) -> list[CellFormula]:
    """Build an ordered top-to-bottom series from one worksheet column span (inclusive)."""
    from openpyxl.utils.cell import get_column_letter
    letter = get_column_letter(col)
    return [CellFormula(f"{letter}{row}", ws[f"{letter}{row}"].value)
            for row in range(start_row, end_row + 1)]
