"""Parity between the saved-workbook (openpyxl) and in-memory (future live-Excel) sources.

The whole point: the *same* detector produces *identical* findings, coverage, ordering and
confidence regardless of where the formulas came from. No second detection pathway.
"""

import pytest
from openpyxl import Workbook, load_workbook

from excel_toolkit.formula_audit import (
    InMemoryWorksheet,
    scan_worksheet,
    snapshot_from_com,
)

from .fixtures_ws import SHEET_LABELS, build_ws_fixtures, expected_cells


@pytest.fixture(scope="module")
def wb(tmp_path_factory):
    path = build_ws_fixtures(tmp_path_factory.mktemp("ws") / "corpus.xlsx")
    return load_workbook(path)


def _assert_reports_equal(a, b):
    assert a.sheet == b.sheet
    assert a.findings == b.findings                      # coords, kind, axes, expected, conflict, message
    assert (a.total_formulas, a.resolved, a.unsupported, a.analysed_formulas) == \
           (b.total_formulas, b.resolved, b.unsupported, b.analysed_formulas)
    assert a.unsupported_cells == b.unsupported_cells
    assert a.skipped_cells == b.skipped_cells


@pytest.mark.parametrize("sheet", list(SHEET_LABELS))
def test_saved_and_snapshot_produce_identical_reports(wb, sheet):
    ws = wb[sheet]
    snapshot = InMemoryWorksheet.from_openpyxl(ws)
    saved = scan_worksheet(ws)
    snap = scan_worksheet(snapshot)
    _assert_reports_equal(saved, snap)
    # and the expected anomalies are still found via the snapshot
    assert expected_cells(SHEET_LABELS[sheet]) <= {f.coord for f in snap.findings}


def test_snapshot_carries_sheet_identity():
    snap = InMemoryWorksheet({"B2": "=B1*2"}, title="Budget 2026")
    assert scan_worksheet(snap).sheet == "Budget 2026"


def test_in_memory_snapshot_hand_built_with_constructs():
    # horizontal + vertical anomalies, blanks, constants, actual->forecast, unsupported, identity
    cells = {
        "A1": 0.05,
        # actuals (constants) then forecast (formulas) in row 2 -> boundary, no FP
        "A2": "Rev", "B2": 100, "C2": 120, "D2": "=C2+$A$1", "E2": "=D2+$A$1", "F2": "=E2+$A$1",
        # base row 3 (constants)
        "B3": 10, "C3": 10, "D3": 10, "E3": 10, "F3": 10,
        # row 4 consistent (=X2-X3) with a singleton anomaly at E4 (=E2-D3)
        "A4": "GP", "B4": "=B2-B3", "C4": "=C2-C3", "D4": "=D2-D3", "E4": "=E2-D3", "F4": "=F2-F3",
        # vertical column H consistent (=above+1) with an INDIRECT (unsupported) in the middle
        "H1": 1, "H2": "=H1+1", "H3": '=INDIRECT("A1")', "H4": "=H3+1", "H5": "=H4+1",
    }
    snap = InMemoryWorksheet(cells, title="S")
    report = scan_worksheet(snap)
    flagged = {f.coord for f in report.findings}
    assert "E4" in flagged                               # row singleton found
    assert any(c == "H3" for c, _s in report.unsupported_cells)  # INDIRECT surfaced
    assert report.sheet == "S"


def test_unsaved_edit_is_detected_via_snapshot():
    # The saved file is clean; an unsaved edit introduces an anomaly the snapshot carries.
    saved = {f"{c}1": 1 for c in "BCDE"}
    saved.update({f"{c}2": f"={c}1*2" for c in "BCDE"})
    assert scan_worksheet(InMemoryWorksheet(saved, "S")).findings == []
    unsaved = dict(saved)
    unsaved["D2"] = "=D1*3"                               # edited in Excel, not yet saved
    flagged = {f.coord for f in scan_worksheet(InMemoryWorksheet(unsaved, "S")).findings}
    assert flagged == {"D2"}


# --- snapshot-conversion logic from a FAKE COM worksheet (no real Excel) -----------------

class _FakeCount:
    def __init__(self, n): self.Count = n


class _FakeRange:
    def __init__(self, row, col, nrows, ncols):
        self.Row, self.Column = row, col
        self.Rows, self.Columns = _FakeCount(nrows), _FakeCount(ncols)


class _FakeCell:
    def __init__(self, value):
        if isinstance(value, str) and value.startswith("="):
            self.HasFormula, self.Formula, self.Value = True, value, None
        else:
            self.HasFormula, self.Formula, self.Value = False, value, value


class _FakeComWorksheet:
    """Mimics the handful of COM members snapshot_from_com uses."""

    def __init__(self, cells: dict, name: str):
        self._cells = {k if isinstance(k, tuple) else _a1(k): v for k, v in cells.items()}
        self.Name = name
        rows = [r for r, _c in self._cells]
        cols = [c for _r, c in self._cells]
        self._top, self._left = min(rows), min(cols)
        self.UsedRange = _FakeRange(self._top, self._left,
                                    max(rows) - self._top + 1, max(cols) - self._left + 1)

    def Cells(self, row, col):
        return _FakeCell(self._cells.get((row, col)))


def _a1(coord):
    from openpyxl.utils.cell import coordinate_to_tuple
    return coordinate_to_tuple(coord)


def test_snapshot_from_fake_com_matches_in_memory():
    cells = {"B1": 1, "C1": 1, "D1": 1, "E1": 1,
             "B2": "=B1*2", "C2": "=C1*2", "D2": "=D1*3", "E2": "=E1*2"}  # D2 anomaly
    com_snapshot = snapshot_from_com(_FakeComWorksheet(cells, "Live"))
    reference = InMemoryWorksheet(cells, "Live")
    a, b = scan_worksheet(com_snapshot), scan_worksheet(reference)
    _assert_reports_equal(a, b)
    assert {f.coord for f in a.findings} == {"D2"}
    assert a.sheet == "Live"
