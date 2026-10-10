"""Formula-audit engine (framework-independent, no Excel COM).

First capability: detect a formula that breaks the pattern of its neighbours, using
relative-reference (R1C1) normalisation. Pure Python over openpyxl's tokenizer; operates
on formula strings read from a workbook, never on a live Excel instance.
"""

from .detector import (
    CellFormula,
    Finding,
    SeriesReport,
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

__all__ = [
    "CellFormula", "Finding", "SeriesReport", "scan_series",
    "series_from_row", "series_from_column",
    "NormalResult", "normalize_formula", "normalize_cell_value",
    "RESOLVED", "PARTIAL", "UNSUPPORTED", "CONST",
]
