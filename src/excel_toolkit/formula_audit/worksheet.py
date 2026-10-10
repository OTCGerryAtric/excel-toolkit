"""Automatic worksheet scanning: discover formula series across rows and columns,
segment them conservatively, and apply the high-confidence flanked-singleton rule.

Series discovery
----------------
A *line* is one worksheet row or column. Each line is split by ``segment_line`` into
maximal runs of adjacent cells of the **same kind**:
- ``formula`` — the cell holds a formula (``=...``);
- ``const``   — the cell holds a plain number.
Blanks, text labels/headings, dates and anything else are **boundaries** that end a run,
and a run also ends where the kind changes (so an actuals→forecast transition, i.e. a
numbers run meeting a formulas run, is split). Only ``formula`` segments are analysed.

This keeps segmentation conservative: totals at the end of a run are edge cells (never
flagged), boundaries between number blocks and formula blocks are split apart, and a
formula that interrupts blanks/labels simply never gains enough neighbours to be judged.

Coverage is reported on two independent axes, kept separate on purpose:
- **dependency coverage** = formulas whose references normalised / total formulas;
- **pattern coverage**    = formulas that sat in an analysable series (had neighbour
  context) / total formulas.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from .detector import dominant_pattern, flanked_outliers
from .normalize import RESOLVED, UNSUPPORTED, normalize_cell_value


def _member_kind(value) -> str | None:
    """'formula', 'const' (plain number), or None (boundary: blank/text/date/other)."""
    if isinstance(value, str):
        return "formula" if value.startswith("=") else None
    if isinstance(value, bool):
        return "const"
    if isinstance(value, (int, float)):
        return "const"
    return None


def segment_line(values: list) -> list[tuple[int, int, str]]:
    """Split a line's raw values into homogeneous runs.

    Returns inclusive ``(start, end, kind)`` spans where ``kind`` is 'formula' or 'const'.
    Runs end at any boundary cell and wherever the kind changes.
    """
    segments: list[tuple[int, int, str]] = []
    start: int | None = None
    kind: str | None = None
    for i, value in enumerate(values):
        k = _member_kind(value)
        if k is None:
            if start is not None:
                segments.append((start, i - 1, kind))
                start, kind = None, None
        elif start is None:
            start, kind = i, k
        elif k != kind:
            segments.append((start, i - 1, kind))
            start, kind = i, k
    if start is not None:
        segments.append((start, len(values) - 1, kind))
    return segments


@dataclass(frozen=True)
class WorksheetFinding:
    coord: str
    row: int
    col: int
    axes: tuple[str, ...]      # ("row",), ("column",) or ("row", "column")
    observed: str
    expected: str
    confidence: str
    message: str


@dataclass
class WorksheetReport:
    total_formulas: int = 0
    resolved: int = 0
    unsupported: int = 0
    analysed_formulas: int = 0  # formulas that sat in an analysable series
    findings: list[WorksheetFinding] = field(default_factory=list)
    unsupported_cells: list[tuple[str, str]] = field(default_factory=list)
    skipped_cells: list[tuple[str, str]] = field(default_factory=list)

    @property
    def dependency_coverage(self) -> float:
        return self.resolved / self.total_formulas if self.total_formulas else 0.0

    @property
    def pattern_coverage(self) -> float:
        return self.analysed_formulas / self.total_formulas if self.total_formulas else 0.0


def _lines(ws):
    """Yield ('row'|'column', [(coord, value), ...]) for every row then every column."""
    max_row, max_col = ws.max_row, ws.max_column
    for r in range(1, max_row + 1):
        yield "row", [(f"{get_column_letter(c)}{r}", ws.cell(row=r, column=c).value)
                      for c in range(1, max_col + 1)]
    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        yield "column", [(f"{letter}{r}", ws.cell(row=r, column=c).value)
                         for r in range(1, max_row + 1)]


def scan_worksheet(ws, *, min_group: int = 3) -> WorksheetReport:
    """Scan a worksheet for high-confidence flanked-singleton formula inconsistencies.

    Reads only (never writes). Findings are deduplicated by cell and returned in
    deterministic top-to-bottom, left-to-right order, with the axes that flagged each cell.
    """
    report = WorksheetReport()
    normal_cache: dict[str, object] = {}
    analysed: set[str] = set()
    formula_coords: set[str] = set()
    # coord -> {"axes": set, "observed": str, "expected": str}
    merged: dict[str, dict] = {}

    def normal_of(coord, value):
        nr = normal_cache.get(coord)
        if nr is None:
            nr = normalize_cell_value(value, coord)
            normal_cache[coord] = nr
        return nr

    for axis, line in _lines(ws):
        values = [v for _c, v in line]
        coords = [c for c, _v in line]
        for start, end, kind in segment_line(values):
            if kind != "formula":
                continue
            seg_coords = coords[start:end + 1]
            seg_patterns = [normal_of(c, values[idx]).pattern
                            for idx, c in zip(range(start, end + 1), seg_coords)]
            dom, count = dominant_pattern(seg_patterns)
            if dom is None or count < min_group:
                continue
            analysed.update(seg_coords)  # these formulas had neighbour context
            for f in flanked_outliers(seg_coords, seg_patterns, dom):
                entry = merged.setdefault(f.coord, {"axes": set(), "observed": f.observed,
                                                    "expected": f.expected})
                entry["axes"].add(axis)

    # One normalisation pass over every formula cell for coverage counting.
    for axis, line in _lines(ws):
        if axis != "row":
            continue
        for coord, value in line:
            if _member_kind(value) == "formula":
                formula_coords.add(coord)
                nr = normal_of(coord, value)
                if nr.coverage == RESOLVED:
                    report.resolved += 1
                else:
                    report.unsupported += 1
                    report.unsupported_cells.append((coord, nr.coverage))

    report.total_formulas = len(formula_coords)
    report.analysed_formulas = len(analysed & formula_coords)

    # Skipped = resolved formulas that never gained series context (isolated / short runs).
    resolved_coords = {c for c in formula_coords
                       if normal_cache[c].coverage == RESOLVED}
    for coord in sorted(resolved_coords - analysed, key=coordinate_to_tuple):
        report.skipped_cells.append((coord, "insufficient series context"))
    report.unsupported_cells.sort(key=lambda t: coordinate_to_tuple(t[0]))

    for coord, entry in merged.items():
        row, col = coordinate_to_tuple(coord)
        axes = tuple(sorted(entry["axes"]))
        where = " and ".join(axes)
        report.findings.append(WorksheetFinding(
            coord=coord, row=row, col=col, axes=axes,
            observed=entry["observed"], expected=entry["expected"], confidence="high",
            message=(f"Formula differs from the consistent pattern of its neighbours "
                     f"({where}). Potential inconsistency — review whether the difference "
                     f"is intentional."),
        ))
    report.findings.sort(key=lambda f: (f.row, f.col))
    return report
