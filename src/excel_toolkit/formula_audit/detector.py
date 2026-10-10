"""Detect formulas that break the pattern established by their neighbours.

Two conservative, high-confidence rules (both emit only ``high`` confidence — there are no
medium/low tiers and no automatic correction):

1. **Singleton** (since Gate 1): a lone interior cell whose normalised pattern differs
   from a dominant pattern that both neighbours share — and whose pattern occurs **once**
   in the series (a recurring difference is structural, not an error).

2. **Block** (Gate 3): a **short** run (2..``max_block``) of consecutive cells that all
   share one non-dominant pattern, bounded on **both** sides by a **strongly established**
   dominant pattern, where that minority pattern occurs **only** in this run. Multi-cell
   anomalies require stronger evidence than singletons (``min_dominant`` ≥ 4 and the
   dominant must be at least twice the run length). Longer runs, recurring patterns, and
   two-regime boundaries (e.g. financial-year changes) are suppressed as ambiguous.

Coverage (resolved/partial/unsupported) is dependency-resolution coverage, reported
separately from pattern confidence.

Known limitations (deliberate, for a conservative high-confidence tool):
- The singleton rule requires the outlier's pattern to be **unique** in the series. This
  suppresses legitimate recurring exceptions (e.g. quarter-end formulas) but also means a
  **genuine error that happens to be repeated elsewhere in the same series is suppressed**
  (a false negative) — repetition is treated as structural, not accidental.
- Only short consecutive minority runs (length 2..MAX_BLOCK) are flagged; a genuine run of
  3+ identical consecutive errors is suppressed as ambiguous (a false negative).
- Only ``high`` confidence is emitted; there is no automatic correction.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .normalize import PARTIAL, RESOLVED, UNSUPPORTED, NormalResult, normalize_cell_value

FLANKED_MESSAGE = (
    "Formula differs from the consistent pattern of the formulas on both sides. "
    "Potential inconsistency — review whether the difference is intentional."
)
BLOCK_MESSAGE = (
    "Short run of consecutive formulas that differ from the strongly established pattern "
    "on both sides. Potential inconsistency — review whether the difference is intentional."
)

# Gate-3 block thresholds (conservative).
MAX_BLOCK = 2          # only short minority runs (pairs) are considered
MIN_BLOCK_DOMINANT = 4  # the flanking dominant must be strongly established


@dataclass(frozen=True)
class CellFormula:
    coord: str
    raw: object  # formula string, constant, or None


@dataclass(frozen=True)
class Finding:
    coord: str
    observed: str
    expected: str
    confidence: str          # always "high" in this detector
    message: str
    kind: str = "singleton"  # "singleton" | "block"
    block: tuple = ()        # for block findings: the coords of the whole run


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
    """Interior singletons: pattern differs from ``dominant``, both neighbours match it, and
    the pattern occurs exactly once in the series (so recurring variations are not flagged)."""
    counts = Counter(p for p in patterns if p is not None)
    findings = []
    for i in range(1, len(coords) - 1):  # interior only; never the first or last cell
        p = patterns[i]
        if p is None or p == dominant:
            continue
        if patterns[i - 1] == dominant and patterns[i + 1] == dominant and counts[p] == 1:
            findings.append(Finding(coords[i], p, dominant, "high", FLANKED_MESSAGE, kind="singleton"))
    return findings


def minority_blocks(coords: list[str], patterns: list[str | None], dominant: str,
                    *, max_block: int = MAX_BLOCK, min_dominant: int = MIN_BLOCK_DOMINANT) -> list[Finding]:
    """Short consecutive minority runs bounded on both sides by a strong dominant pattern."""
    counts = Counter(p for p in patterns if p is not None)
    if counts.get(dominant, 0) < min_dominant:
        return []
    n = len(patterns)
    findings: list[Finding] = []
    i = 0
    while i < n:
        p = patterns[i]
        if p is None or p == dominant:
            i += 1
            continue
        j = i
        while j < n and patterns[j] == p:  # maximal run of identical pattern p
            j += 1
        length = j - i
        bounded = i - 1 >= 0 and patterns[i - 1] == dominant and j < n and patterns[j] == dominant
        unique_run = counts[p] == length           # all occurrences of p are in this run
        strong = counts[dominant] >= 2 * length
        if 2 <= length <= max_block and bounded and unique_run and strong:
            run = tuple(coords[i:j])
            findings.extend(
                Finding(coords[k], p, dominant, "high", BLOCK_MESSAGE, kind="block", block=run)
                for k in range(i, j)
            )
        i = j
    return findings


def scan_series(cells: list[CellFormula], *, min_group: int = 3) -> SeriesReport:
    """Scan one ordered 1-D run of cells (a row or column span) for singleton and block outliers."""
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
        report.findings = (flanked_outliers(coords, patterns, report.dominant)
                           + minority_blocks(coords, patterns, report.dominant))
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
