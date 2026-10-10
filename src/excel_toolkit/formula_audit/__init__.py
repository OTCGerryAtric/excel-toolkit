"""Formula-audit engine (framework-independent, no Excel COM).

First capability: detect a formula that breaks the pattern of its neighbours, using
relative-reference (R1C1) normalisation. Pure Python over openpyxl's tokenizer; operates
on formula strings read from a workbook, never on a live Excel instance.
"""

from .detector import (
    CellFormula,
    Finding,
    SeriesReport,
    dominant_pattern,
    flanked_outliers,
    scan_series,
    series_from_column,
    series_from_row,
)
from .normalize import (
    CONST,
    PARTIAL,
    RESOLVED,
    UNSUPPORTED,
    NormalResult,
    normalize_cell_value,
    normalize_formula,
)
from .worksheet import (
    WorksheetFinding,
    WorksheetReport,
    scan_worksheet,
    segment_line,
)

__all__ = [
    "CellFormula", "Finding", "SeriesReport", "scan_series",
    "series_from_row", "series_from_column", "dominant_pattern", "flanked_outliers",
    "NormalResult", "normalize_formula", "normalize_cell_value",
    "RESOLVED", "PARTIAL", "UNSUPPORTED", "CONST",
    "WorksheetFinding", "WorksheetReport", "scan_worksheet", "segment_line",
]
