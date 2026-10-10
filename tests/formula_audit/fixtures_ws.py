"""Worksheet-level regression corpus: realistic financial-model layouts with labelled
anomalies and legitimate variations.

Labels are grouped into *blocks* (a singleton is a block of size 1):
- ``expected_blocks``  : runs that SHOULD be flagged;
- ``known_fn_blocks``  : genuine anomalies we deliberately DON'T flag (conservative limit),
                         e.g. a 3-cell run (longer than the supported short-block max);
- ``axes``             : expected axes per flagged coord.
Everything else is a legitimate cell that must NOT be flagged (incl. quarterly exceptions,
financial-year boundaries and actual→forecast transitions).
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter

MONTHS = list("BCDEFGHIJKLM")  # Jan..Dec in columns B..M


def _revenue_grid(ws):
    ws["A20"], ws["B20"] = "Growth", 0.05
    ws["A21"], ws["B21"] = "COGS %", 0.40
    ws["A1"] = "Item"
    for i, month in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun"]):
        ws[f"{get_column_letter(2 + i)}1"] = month
    ws["H1"] = "Total"
    ws["A2"], ws["B2"] = "Revenue", 1000
    for c in range(3, 8):
        ws[f"{get_column_letter(c)}2"] = f"={get_column_letter(c - 1)}2*(1+$B$20)"
    ws["H2"] = "=SUM(B2:G2)"
    ws["A3"], ws["B3"] = "COGS", 400
    for c in range(3, 8):
        ws[f"{get_column_letter(c)}3"] = f"={get_column_letter(c)}2*$B$21"
    ws["H3"] = "=SUM(B3:G3)"
    ws["A4"] = "Gross Profit"
    for c in range(2, 8):
        col = get_column_letter(c)
        ws[f"{col}4"] = f"={col}2-{col}3"
    ws["E4"] = "=E2-D3"          # singleton anomaly
    ws["H4"] = "=SUM(B4:G4)"


def _balance_roll(ws):
    ws["D1"] = 10
    ws["A2"], ws["B2"] = "Opening", 100
    for r in range(3, 9):
        ws[f"A{r}"] = f"Period {r - 2}"
        ws[f"B{r}"] = f"=B{r - 1}+$D$1"
    ws["B6"] = "=B5*$D$1"        # singleton anomaly


def _overlap(ws):
    for c in range(2, 6):
        ws[f"{get_column_letter(c)}1"] = c * 10
    for r in range(2, 6):
        ws[f"A{r}"] = r
    for r in range(2, 6):
        for c in range(2, 6):
            col = get_column_letter(c)
            ws[f"{col}{r}"] = f"=$A{r}*{col}$1"
    ws["D3"] = "=$A3*C$1"        # anomaly in both row 3 and column D


def _consecutive_block(ws):
    # two adjacent identical reference errors -> a short minority block (Gate 3 flags this)
    for c in range(2, 8):
        ws[f"{get_column_letter(c)}1"] = c
    ws["A2"] = "Doubler"
    for c in range(2, 8):
        ws[f"{get_column_letter(c)}2"] = f"={get_column_letter(c)}1*2"
    ws["D2"] = "=D1*3"
    ws["E2"] = "=E1*3"


def _triple_suppressed(ws):
    # three adjacent errors -> longer than the supported short-block max, suppressed (known FN)
    for c in range(2, 9):
        ws[f"{get_column_letter(c)}1"] = c
    ws["A2"] = "Doubler"
    for c in range(2, 9):
        ws[f"{get_column_letter(c)}2"] = f"={get_column_letter(c)}1*2"
    ws["D2"] = "=D1*3"
    ws["E2"] = "=E1*3"
    ws["F2"] = "=F1*3"


def _consec_vertical(ws):
    ws["D1"] = 1
    ws["A2"], ws["B2"] = "Opening", 100
    for r in range(3, 9):
        ws[f"B{r}"] = f"=B{r - 1}+$D$1"
    ws["B5"] = "=B4*$D$1"       # consecutive vertical block B5:B6
    ws["B6"] = "=B5*$D$1"


def _quarterly(ws):
    # legitimate recurring quarter-end exception (Mar/Jun/Sep/Dec use a second driver)
    ws["A1"], ws["A2"], ws["A4"] = 0.05, 0.10, 100
    q_ends = {"D", "G", "J", "M"}  # Mar, Jun, Sep, Dec
    prev = "A"
    for col in MONTHS:
        driver = "$A$2" if col in q_ends else "$A$1"
        ws[f"{col}4"] = f"={prev}4+{driver}"
        prev = col


def _fy_boundary(ws):
    # legitimate two-regime boundary: FY1 additive (B..G), FY2 multiplicative (H..M)
    ws["A1"], ws["A2"], ws["A4"] = 0.05, 1.10, 100
    prev = "A"
    for col in MONTHS:
        op, driver = ("+", "$A$1") if col in list("BCDEFG") else ("*", "$A$2")
        ws[f"{col}4"] = f"={prev}4{op}{driver}"
        prev = col


def _actual_forecast(ws):
    # legitimate actuals (constants) then forecast (formulas)
    ws["A1"] = 0.05
    ws["B4"], ws["C4"], ws["D4"] = 100, 110, 120
    prev = "D"
    for col in list("EFGH"):
        ws[f"{col}4"] = f"={prev}4+$A$1"
        prev = col


def build_ws_fixtures(path: Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    for name, builder in (
        ("revenue_grid", _revenue_grid),
        ("balance_roll", _balance_roll),
        ("overlap", _overlap),
        ("consecutive_block", _consecutive_block),
        ("triple_suppressed", _triple_suppressed),
        ("consec_vertical", _consec_vertical),
        ("quarterly", _quarterly),
        ("fy_boundary", _fy_boundary),
        ("actual_forecast", _actual_forecast),
        ("totals_headings", _totals_headings),
    ):
        builder(wb.create_sheet(name))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _totals_headings(ws):
    for c in range(2, 5):
        ws[f"{get_column_letter(c)}10"] = 5
        ws[f"{get_column_letter(c)}11"] = 3
    ws["A1"] = "Metric"
    for i, m in enumerate(["M1", "M2", "M3"]):
        ws[f"{get_column_letter(2 + i)}1"] = m
    ws["E1"] = "Total"
    ws["A2"] = "Sales"
    for c in range(2, 5):
        ws[f"{get_column_letter(c)}2"] = f"={get_column_letter(c)}10*2"
    ws["E2"] = "=SUM(B2:D2)"
    ws["A3"] = "Costs"
    for c in range(2, 5):
        ws[f"{get_column_letter(c)}3"] = f"={get_column_letter(c)}11*2"
    ws["E3"] = "=SUM(B3:D3)"
    ws["A5"] = "Total"
    for c in range(2, 6):
        col = get_column_letter(c)
        ws[f"{col}5"] = f"=SUM({col}2:{col}3)"


SHEET_LABELS = {
    "revenue_grid":      {"expected_blocks": [{"E4"}],         "known_fn_blocks": [], "axes": {"E4": ("row",)}},
    "balance_roll":      {"expected_blocks": [{"B6"}],         "known_fn_blocks": [], "axes": {"B6": ("column",)}},
    "overlap":           {"expected_blocks": [{"D3"}],         "known_fn_blocks": [], "axes": {"D3": ("column", "row")}},
    "consecutive_block": {"expected_blocks": [{"D2", "E2"}],   "known_fn_blocks": [], "axes": {"D2": ("row",), "E2": ("row",)}},
    "triple_suppressed": {"expected_blocks": [],               "known_fn_blocks": [{"D2", "E2", "F2"}], "axes": {}},
    "consec_vertical":   {"expected_blocks": [{"B5", "B6"}],   "known_fn_blocks": [], "axes": {"B5": ("column",), "B6": ("column",)}},
    "quarterly":         {"expected_blocks": [],               "known_fn_blocks": [], "axes": {}},
    "fy_boundary":       {"expected_blocks": [],               "known_fn_blocks": [], "axes": {}},
    "actual_forecast":   {"expected_blocks": [],               "known_fn_blocks": [], "axes": {}},
    "totals_headings":   {"expected_blocks": [],               "known_fn_blocks": [], "axes": {}},
}


def expected_cells(label) -> set:
    return set().union(*label["expected_blocks"]) if label["expected_blocks"] else set()


def known_fn_cells(label) -> set:
    return set().union(*label["known_fn_blocks"]) if label["known_fn_blocks"] else set()
