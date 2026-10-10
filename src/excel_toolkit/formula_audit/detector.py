"""Detect a formula that breaks the pattern established by its neighbours.

First increment — deliberately conservative to keep false positives near zero:

A cell is flagged only when it is an **interior singleton** whose normalised pattern
differs from a **dominant** pattern that its immediate left and right neighbours both
share. This is the ``=D2*(1+D4)`` case (March) sitting between Jan/Feb and Apr.

It deliberately does **not** flag:
- the first or last cell in a series (period openings/closings, year-end totals), because
  an edge cell has no neighbour on one side;
- a cell at the boundary between two consistent blocks (e.g. actuals → forecast), because
  its neighbour on one side is not the dominant pattern;
- cells we could not normalise (INDIRECT/OFFSET, spill, names, tables, external) — these
  are reported as unsupported, never as "clean".

Two independent axes are reported, and kept separate on purpose:
- **coverage** (per cell): could we resolve the formula's references at all? resolved /
  partial / unsupported. This is dependency-resolution coverage.
- **confidence** (per finding): how sure are we the resolved-but-different formula is a
  genuine anomaly? Here only ``high`` is emitted (flanked interior singleton).
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .normalize import PARTIAL, RESOLVED, UNSUPPORTED, NormalResult, normalize_cell_value


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


def scan_series(cells: list[CellFormula], *, min_group: int = 3) -> SeriesReport:
    """Scan an ordered 1-D run of cells (a row or column span) for a lone outlier.

    ``min_group``: the dominant pattern must appear at least this many times before any
    outlier is flagged, so we never "flag the majority" in a tiny or noisy series.
    """
    report = SeriesReport(total=len(cells))
    normals: list[NormalResult] = []
    for cell in cells:
        nr = normalize_cell_value(cell.raw, cell.coord)
        normals.append(nr)
        if nr.coverage == RESOLVED:
            report.resolved += 1
        elif nr.coverage == PARTIAL:
            report.partial += 1
            report.unsupported_cells.append((cell.coord, PARTIAL))
        else:
            report.unsupported += 1
            report.unsupported_cells.append((cell.coord, UNSUPPORTED))

    patterns = [nr.pattern for nr in normals]
    counts = Counter(p for p in patterns if p is not None)
    if not counts:
        return report
    report.dominant, report.dominant_count = counts.most_common(1)[0]

    if report.dominant_count < min_group:
        return report  # no established pattern to deviate from

    for i in range(1, len(cells) - 1):  # interior only
        p = patterns[i]
        if p is None or p == report.dominant:
            continue
        if patterns[i - 1] == report.dominant and patterns[i + 1] == report.dominant:
            report.findings.append(Finding(
                coord=cells[i].coord,
                observed=p,
                expected=report.dominant,
                confidence="high",
                message=("Formula differs from the consistent pattern of the formulas on "
                         "both sides. Potential inconsistency — review whether the "
                         "difference is intentional."),
            ))
    return report


def series_from_row(ws, row: int, start_col: int, end_col: int) -> list[CellFormula]:
    """Build an ordered left-to-right series from one worksheet row span (inclusive)."""
    from openpyxl.utils.cell import get_column_letter
    out = []
    for col in range(start_col, end_col + 1):
        coord = f"{get_column_letter(col)}{row}"
        out.append(CellFormula(coord=coord, raw=ws[coord].value))
    return out


def series_from_column(ws, col: int, start_row: int, end_row: int) -> list[CellFormula]:
    """Build an ordered top-to-bottom series from one worksheet column span (inclusive)."""
    from openpyxl.utils.cell import get_column_letter
    letter = get_column_letter(col)
    return [CellFormula(coord=f"{letter}{row}", raw=ws[f"{letter}{row}"].value)
            for row in range(start_row, end_row + 1)]
