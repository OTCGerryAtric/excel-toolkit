"""A small read-only worksheet interface so the detector is independent of its input source.

The detector only needs: sheet identity, bounds, and the value at each coordinate
(a formula string, a constant, or ``None``). Both a saved workbook (openpyxl) and a live
Excel workbook (a COM snapshot, built later) implement this one interface, so the *same*
analysis and confidence rules run against either — there is no second detection pathway.

    saved .xlsx ──openpyxl──┐
                            ├─▶ WorksheetSource ─▶ scan_worksheet (one detector)
    live Excel ──COM snap──┘

Reads are non-mutating by contract (see ``OpenpyxlWorksheet.value_at``).
"""

from __future__ import annotations

from typing import Protocol, runtime_checkable

from openpyxl.utils.cell import coordinate_to_tuple


@runtime_checkable
class WorksheetSource(Protocol):
    """Read-only view a detector needs. 1-based ``row``/``col`` like Excel/openpyxl."""

    @property
    def title(self) -> str: ...
    @property
    def max_row(self) -> int: ...
    @property
    def max_col(self) -> int: ...
    def value_at(self, row: int, col: int): ...  # formula str | constant | None


class OpenpyxlWorksheet:
    """Adapter over an openpyxl worksheet (a saved workbook)."""

    def __init__(self, ws):
        self._ws = ws

    @property
    def title(self) -> str:
        return self._ws.title

    @property
    def max_row(self) -> int:
        return self._ws.max_row

    @property
    def max_col(self) -> int:
        return self._ws.max_column

    def value_at(self, row: int, col: int):
        # ws._cells.get never creates a Cell (unlike ws.cell()/iter_rows) -> no mutation.
        cell = self._ws._cells.get((row, col))
        return cell.value if cell is not None else None


class InMemoryWorksheet:
    """A plain in-memory snapshot — the shape a live Excel COM read will supply.

    Holds exactly the formulas/constants present (including *unsaved* edits, which simply
    means the snapshot carries different values than the file on disk). Build from A1-keyed
    or (row, col)-keyed cells.
    """

    def __init__(self, cells: dict, title: str = "Sheet1"):
        self._cells: dict[tuple[int, int], object] = {}
        for key, value in cells.items():
            rc = key if isinstance(key, tuple) else coordinate_to_tuple(key)
            self._cells[rc] = value
        self._title = title
        self._max_row = max((r for r, _c in self._cells), default=0)
        self._max_col = max((c for _r, c in self._cells), default=0)

    @classmethod
    def from_openpyxl(cls, ws) -> "InMemoryWorksheet":
        cells = {rc: cell.value for rc, cell in ws._cells.items() if cell.value is not None}
        return cls(cells, title=ws.title)

    @property
    def title(self) -> str:
        return self._title

    @property
    def max_row(self) -> int:
        return self._max_row

    @property
    def max_col(self) -> int:
        return self._max_col

    def value_at(self, row: int, col: int):
        return self._cells.get((row, col))


def as_source(obj) -> WorksheetSource:
    """Return ``obj`` if it already implements the interface, else wrap an openpyxl ws."""
    if hasattr(obj, "value_at") and hasattr(obj, "max_row") and hasattr(obj, "max_col"):
        return obj
    return OpenpyxlWorksheet(obj)


def _as_grid(data) -> list[list]:
    """Normalise what ``Range.Formula`` returns into a list of rows.

    Excel/pywin32 returns a 2-D tuple-of-tuples for a multi-cell range, but a **scalar** for
    a single cell (and a 1-D tuple can appear for a single row/column), so we normalise all
    three shapes."""
    if data is None:
        return []
    if isinstance(data, (list, tuple)):
        if len(data) and isinstance(data[0], (list, tuple)):
            return [list(row) for row in data]          # 2-D
        return [list(data)]                              # 1-D -> one row
    return [[data]]                                      # scalar -> single cell


def _coerce_constant(value):
    """Normalise a constant cell's **typed** ``.Value`` for the detector.

    Preserves real types — number, text, bool, date pass through unchanged — so a numeric
    cell stays numeric and a *text* cell that happens to contain ``"100"`` stays text. Blanks
    become ``None``; Excel error values / unexpected COM types become their string form,
    which the detector treats as a (non-formula) boundary.
    """
    import datetime as _dt
    if value is None or value == "":
        return None
    if isinstance(value, (str, bool, int, float, _dt.datetime, _dt.date)):
        return value
    return str(value)


def _cell_from(formula, value):
    """Decide a cell's content from its paired ``.Formula`` and ``.Value`` entries.

    A formula is identified by its ``.Formula`` text starting with ``=`` (kept verbatim).
    Everything else is a constant and takes its **typed** ``.Value`` — crucially NOT the
    ``.Formula`` text, which Excel returns as a string even for numbers (that was what
    misclassified numeric constants as text boundaries in the first live run).
    """
    if isinstance(formula, str) and formula.startswith("="):
        return formula
    return _coerce_constant(value)


def snapshot_from_com(com_ws) -> InMemoryWorksheet:
    """Build an InMemoryWorksheet from a live Excel COM worksheet object.

    Reads the used range in **two bulk COM calls** — ``UsedRange.Formula`` (formula text)
    and ``UsedRange.Value`` (typed constant values) — rather than tens of thousands of
    per-cell accesses, then combines them per ``_cell_from`` so formulas keep their text and
    constants keep their real type. Snapshot-conversion logic only; tested against a FAKE COM
    object offline, NOT run against real Excel here.

    Unresolved live-Excel assumptions (to confirm on the authorised test machine):
      * ``.Formula`` is always US A1 syntax with ',' separators (what the tokenizer expects),
        NOT ``.FormulaLocal``; dynamic-array/spill (``Formula2``) is future work;
      * a cell is a formula iff its ``.Formula`` starts with ``=``; otherwise its ``.Value``
        gives the typed constant (number/text/bool/date); errors/unexpected types → string;
      * ``.Formula`` and ``.Value`` share the used range's shape and 1-based top-left
        (``UsedRange.Row``/``.Column``); ``UsedRange`` can over-report and ignores merged cells;
      * both return a scalar for a single-cell range (handled by ``_as_grid``);
      * ``Worksheet.Name`` gives identity.
    """
    used = com_ws.UsedRange
    top, left = used.Row, used.Column
    formulas = _as_grid(used.Formula)
    values = _as_grid(used.Value)
    cells: dict[tuple[int, int], object] = {}
    for i, frow in enumerate(formulas):
        vrow = values[i] if i < len(values) else []
        for j, f in enumerate(frow):
            v = vrow[j] if j < len(vrow) else None
            cell = _cell_from(f, v)
            if cell is not None:
                cells[(top + i, left + j)] = cell
    return InMemoryWorksheet(cells, title=com_ws.Name)
