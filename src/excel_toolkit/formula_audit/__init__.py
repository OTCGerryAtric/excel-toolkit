"""Formula-audit engine (framework-independent, no Excel COM).

First capability: detect a formula that breaks the pattern of its neighbours, using
relative-reference (R1C1) normalisation. Pure Python over openpyxl's tokenizer; operates
on formula strings read from a workbook, never on a live Excel instance.
"""

import openpyxl

from .detector import (
    CellFormula,
    Finding,
    MAX_BLOCK,
    MIN_BLOCK_DOMINANT,
    SeriesReport,
    dominant_pattern,
    flanked_outliers,
    minority_blocks,
    scan_series,
    series_from_column,
    series_from_row,
)

# The scanner relies on openpyxl 3.1.x behaviour (the ws._cells store used for
# non-mutating reads, and the formula tokenizer). This is a runtime check with a regression
# test, not a packaged constraint: the main project's dependency files are intentionally
# unchanged during the isolated-development phase, so ENFORCING a supported version in
# production (pinning openpyxl in pyproject/uv.lock) remains OUTSTANDING.
SUPPORTED_OPENPYXL = ("3.1",)


def openpyxl_supported() -> bool:
    return openpyxl.__version__.startswith(SUPPORTED_OPENPYXL)
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
from .worksheet_source import (
    InMemoryWorksheet,
    OpenpyxlWorksheet,
    WorksheetSource,
    as_source,
    snapshot_from_com,
)
from .integration import (
    a1_sheet_reference,
    findings_payload,
    findings_to_json,
    navigation_target,
)

__all__ = [
    "CellFormula", "Finding", "SeriesReport", "scan_series",
    "series_from_row", "series_from_column", "dominant_pattern", "flanked_outliers",
    "minority_blocks", "MAX_BLOCK", "MIN_BLOCK_DOMINANT",
    "NormalResult", "normalize_formula", "normalize_cell_value",
    "RESOLVED", "PARTIAL", "UNSUPPORTED", "CONST",
    "WorksheetFinding", "WorksheetReport", "scan_worksheet", "segment_line",
    "WorksheetSource", "OpenpyxlWorksheet", "InMemoryWorksheet", "as_source", "snapshot_from_com",
    "a1_sheet_reference", "navigation_target", "findings_payload", "findings_to_json",
    "SUPPORTED_OPENPYXL", "openpyxl_supported",
]
