"""M3d experiment - Excel's own Paste Values run from VBA. DO NOT RUN until authorised.

Question (roadmap, "Clipboard Paste Values (M3d)"): does Excel's built-in Paste Values, invoked from VBA
with CommandBars.ExecuteMso "PasteValues", (1) keep the clipboard, (2) give a native Undo for the paste,
(3) leave the EARLIER Undo history in place, and (4) handle single and multi-area destinations? Two
controls (an object-model write and the legacy Range.PasteSpecial) show what the same measurements look
like when Excel's Undo history IS cleared, so a "kept" result can be trusted.

Boundaries as m2-m3c: a SEPARATE disposable Excel process (DispatchEx: no XLSTART, no PERSONAL.XLSB, no
add-ins), workbooks it creates and closes without saving, the probe module (m3d_paste_probe.bas)
imported into a disposable workbook - nothing installed, the add-in not involved. The run USES THE WINDOWS
CLIPBOARD (shared with the owner's session): do not copy anything while it runs; it is emptied at the end.
Excel message boxes raised by a paste are recorded and dismissed with OK.

  <scratch venv>/python addin/live/m3d_native_paste_experiment.py [--visible]

--visible shows the disposable Excel window (ExecuteMso may need a visible window; the run records this).
Not covered here, for the later manual stage: invocation from the ribbon and a keyboard shortcut, and a
real Ctrl+Z key press (automation can only call Application.Run and Excel's Undo command).
"""

from __future__ import annotations

import argparse
import json
import sys
import threading
import time
from pathlib import Path

import win32clipboard
import win32com.client
import win32con
import win32gui
import win32process

from m2_live_check import PERSONAL, Watchdog, _terminate, sha256

HERE = Path(__file__).resolve().parent
PROBE = HERE / "m3d_paste_probe.bas"

# Source A1:B3 (formula, number formats, text, date, logical); destinations are in D1:H6, which starts as "old".
SOURCE = {"A1": (1234.5, "#,##0.00"), "A2": ("=A1*2", None), "A3": ("text", None),
          "B1": (45000, "dd/mm/yyyy"), "B2": (0.25, "0%"), "B3": (True, None)}
BOX = "D1:H6"
BOX_ROWS, BOX_COLS = 6, 5
HISTORY_CELL = "J1"


class DialogWatcher:
    """Records and dismisses (OK) any message box the disposable Excel shows."""

    def __init__(self, pid: int):
        self.pid, self.captured, self._stop = pid, [], threading.Event()
        threading.Thread(target=self._run, daemon=True).start()

    def take(self) -> list:
        found, self.captured = self.captured, []
        return found

    def stop(self):
        self._stop.set()

    def _run(self):
        while not self._stop.is_set():
            def visit(hwnd, _):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770" \
                        and win32process.GetWindowThreadProcessId(hwnd)[1] == self.pid:
                    texts = []
                    win32gui.EnumChildWindows(hwnd, lambda h, _: texts.append(win32gui.GetWindowText(h)) or True, None)
                    self.captured.append({"title": win32gui.GetWindowText(hwnd), "texts": [t for t in texts if t]})
                    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, win32con.IDOK, 0)
                return True
            win32gui.EnumWindows(visit, None)
            time.sleep(0.3)


def set_clipboard_text(text: str | None) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        if text is not None:
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


class Experiment:
    def __init__(self, visible: bool):
        self.excel = win32com.client.DispatchEx("Excel.Application")
        _, self.pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
        self.watchdog = Watchdog(self.pid)
        self.dialogs = DialogWatcher(self.pid)
        self.results: list[dict] = []
        try:
            self.excel.Visible = visible
            self.excel.DisplayAlerts = False
            self.wb = self.excel.Workbooks.Add()
            self.wb.VBProject.VBComponents.Import(str(PROBE))
            self.source_values = None
        except Exception:
            self.stop()
            raise

    # ----- plumbing -----

    def probe(self, procedure: str) -> str:
        self.watchdog.tick()
        value = self.excel.Run(f"'{self.wb.Name}'!{procedure}")
        self.watchdog.tick()
        return value

    def undo_state(self) -> dict:
        _, enabled, count, entries = self.probe("Probe_UndoState").split("|", 3)
        return {"enabled": enabled == "True", "count": int(count), "entries": [e for e in entries.split(";") if e]}

    def record(self, scenario: str, **observed):
        observed["dialogs"] = self.dialogs.take()
        self.results.append({"scenario": scenario, **observed})
        print(json.dumps({"scenario": scenario, **observed}, default=str))

    def fresh_sheet(self, name: str):
        """A new sheet: source A1:B3, D1:H6 'old', J1 'history'. Object-model writes - done BEFORE any
        Undo history is built, since they clear it."""
        self.wb.Activate()
        ws = self.wb.Worksheets.Add()
        ws.Name = name
        for address, (content, number_format) in SOURCE.items():
            if isinstance(content, str) and content.startswith("="):
                ws.Range(address).Formula = content
            else:
                ws.Range(address).Value2 = content
            if number_format:
                ws.Range(address).NumberFormat = number_format
        ws.Range(BOX).Value2 = "old"
        ws.Range(HISTORY_CELL).Value2 = "history"
        ws.Activate()
        if self.source_values is None:
            self.source_values = [list(row) for row in ws.Range("A1:B3").Value2]
        return ws

    def build_history(self, ws) -> dict:
        """One earlier Undo entry from a user-interface command: Bold on J1."""
        ws.Range(HISTORY_CELL).Select()
        return {"bold": self.probe("Probe_NativeBold"), "undo": self.undo_state()}

    @staticmethod
    def box(ws) -> list:
        return [list(row) for row in ws.Range(BOX).Value2]

    @staticmethod
    def box_details(ws) -> dict:
        cells = ws.Range(BOX)
        return {"formulas": [c.GetAddress(False, False) for c in cells.Cells if c.HasFormula],
                "formats": sorted({c.NumberFormat for c in cells.Cells})}

    def expected_box(self, placements) -> list:
        """'old' everywhere, with the source grid placed at each (row, col) offset inside D1:H6."""
        grid = [["old"] * BOX_COLS for _ in range(BOX_ROWS)]
        for top, left in placements:
            for r, row in enumerate(self.source_values):
                for c, value in enumerate(row):
                    grid[top + r][left + c] = value
        return grid

    # ----- controls: what a cleared Undo history looks like -----

    def controls(self):
        ws = self.fresh_sheet("control_write")
        history = self.build_history(ws)
        ws.Range("D1").Select()
        write = self.probe("Probe_ControlObjectModelWrite")
        after = self.undo_state()
        undo = self.probe("Probe_NativeUndo")
        self.record("control: object-model write after a Bold", history=history, write=write, undo_after=after,
                    native_undo=undo, bold_still_on=bool(ws.Range(HISTORY_CELL).Font.Bold))

        ws = self.fresh_sheet("control_pastespecial")
        history = self.build_history(ws)
        ws.Range("A1:B3").Copy()
        ws.Range("D1").Select()
        paste = self.probe("Probe_ControlPasteSpecial")
        after = self.undo_state()
        mode = self.excel.CutCopyMode
        undo = self.probe("Probe_NativeUndo")
        self.record("control: Range.PasteSpecial (legacy) after a Bold", history=history, paste=paste, undo_after=after,
                    cut_copy_mode=mode, native_undo=undo, box_after_undo_attempt=self.box(ws),
                    bold_still_on=bool(ws.Range(HISTORY_CELL).Font.Bold))
        self.excel.CutCopyMode = False

    # ----- the native command -----

    def native_case(self, scenario: str, destination: str, placements, *, source: str = "A1:B3", cut: bool = False):
        ws = self.fresh_sheet(f"case{len(self.results) + 1}")
        history = self.build_history(ws)
        (ws.Range(source).Cut() if cut else ws.Range(source).Copy())
        after_copy = self.undo_state()
        ws.Range(destination).Select()
        paste = self.probe("Probe_NativePasteValues")
        mode_after = self.excel.CutCopyMode
        box_after, details = self.box(ws), self.box_details(ws)
        undo_after_paste = self.undo_state()
        observed = {"destination": destination, "history": history, "undo_after_copy": after_copy,
                    "paste": paste, "cut_copy_mode_after": mode_after, "undo_after_paste": undo_after_paste,
                    "box_after": box_after, "box_details": details}
        if placements is not None:
            observed["values_as_expected"] = box_after == self.expected_box(placements)
        observed["no_formulas_or_formats_pasted"] = not details["formulas"] and details["formats"] == ["General"]
        undo1 = self.probe("Probe_NativeUndo")
        observed["undo_1"] = undo1
        observed["paste_reversed_by_undo"] = self.box(ws) == self.expected_box([])
        observed["bold_kept_after_undo_1"] = bool(ws.Range(HISTORY_CELL).Font.Bold)
        undo2 = self.probe("Probe_NativeUndo")
        observed["undo_2"] = undo2
        observed["earlier_history_reversed_by_undo_2"] = not ws.Range(HISTORY_CELL).Font.Bold
        self.record(scenario, **observed)
        self.excel.CutCopyMode = False

    def clipboard_reuse(self):
        ws = self.fresh_sheet("reuse")
        ws.Range("A1:B3").Copy()
        ws.Range("D1").Select()
        first = self.probe("Probe_NativePasteValues")
        ws.Range("G1").Select()
        second = self.probe("Probe_NativePasteValues")
        got = self.box(ws)
        self.record("clipboard kept: two pastes from one copy (D1 then G1)", first=first, second=second,
                    cut_copy_mode_after=self.excel.CutCopyMode, box_after=got,
                    both_pastes_as_expected=got == self.expected_box([(0, 0), (0, 3)]))
        self.excel.CutCopyMode = False

    def other_workbook_source(self):
        other = self.excel.Workbooks.Add()
        try:
            src = other.Worksheets(1)
            for address, (content, number_format) in SOURCE.items():
                if isinstance(content, str) and content.startswith("="):
                    src.Range(address).Formula = content
                else:
                    src.Range(address).Value2 = content
            ws = self.fresh_sheet("other_book")
            history = self.build_history(ws)
            src.Range("A1:B3").Copy()
            self.wb.Activate()
            ws.Activate()
            ws.Range("D1").Select()
            paste = self.probe("Probe_NativePasteValues")
            self.record("source in another workbook", history=history, paste=paste,
                        values_as_expected=self.box(ws) == self.expected_box([(0, 0)]),
                        undo_after_paste=self.undo_state(), cut_copy_mode_after=self.excel.CutCopyMode)
            self.excel.CutCopyMode = False
        finally:
            other.Close(SaveChanges=False)

    def outside_clipboard(self):
        ws = self.fresh_sheet("outside_text")
        self.excel.CutCopyMode = False
        set_clipboard_text("42\tabc\r\n7\tdef")
        ws.Range("D1").Select()
        paste = self.probe("Probe_NativePasteValues")
        self.record("clipboard holds text from outside Excel", paste=paste, box_after=self.box(ws),
                    undo_after_paste=self.undo_state())
        ws = self.fresh_sheet("empty_clipboard")
        set_clipboard_text(None)
        ws.Range("D1").Select()
        paste = self.probe("Probe_NativePasteValues")
        self.record("clipboard empty", paste=paste, box_unchanged=self.box(ws) == self.expected_box([]))

    def protected_and_filtered(self):
        ws = self.fresh_sheet("protected")
        ws.Range("A1:B3").Copy()
        ws.Protect()
        ws.Range("D1").Select()
        paste = self.probe("Probe_NativePasteValues")
        self.record("protected sheet (destination locked)", paste=paste, box_unchanged=self.box(ws) == self.expected_box([]))
        self.excel.CutCopyMode = False
        ws.Unprotect()

        ws = self.fresh_sheet("filtered")
        ws.Range("D1:D6").Value2 = (("head",), ("a",), ("b",), ("a",), ("b",), ("a",))
        ws.Range("D1:D6").AutoFilter(Field=1, Criteria1="a")
        ws.Range("A1:A3").Copy()
        ws.Range("E2").Select()
        paste = self.probe("Probe_NativePasteValues")
        self.record("filtered destination (rows 3 and 5 hidden), 3 cells pasted at E2", paste=paste,
                    column_e=[c.Value2 for c in ws.Range("E1:E6").Cells],
                    hidden_rows=[r for r in range(1, 7) if ws.Rows(r).Hidden])
        self.excel.CutCopyMode = False

    # ----- run -----

    def run_all(self):
        self.controls()
        self.native_case("single cell destination D1", "D1", [(0, 0)])
        self.native_case("same-size destination D1:E3", "D1:E3", [(0, 0)])
        self.native_case("twice-the-size destination D1:E6 (Excel repeats the copy)", "D1:E6", [(0, 0), (3, 0)])
        self.native_case("multi-area destination D1,G1", "D1,G1", None)
        self.native_case("multi-area destination D1:E3,G1:H3", "D1:E3,G1:H3", None)
        self.native_case("multi-area source A1,A3 to D1", "D1", None, source="A1,A3")
        self.native_case("after Cut (not Copy)", "D1", None, cut=True)
        self.clipboard_reuse()
        self.other_workbook_source()
        self.outside_clipboard()
        self.protected_and_filtered()

    def stop(self):
        self.dialogs.stop()
        self.watchdog.stop()
        try:
            for book in list(self.excel.Workbooks):
                book.Close(SaveChanges=False)
            self.excel.Quit()
        except Exception:  # noqa: BLE001 - terminated below anyway
            pass
        del self.excel
        time.sleep(1)
        _terminate(self.pid)


def summarise(results: list[dict]) -> dict:
    by = {r["scenario"]: r for r in results}
    core = [by[s] for s in ("single cell destination D1", "same-size destination D1:E3",
                            "twice-the-size destination D1:E6 (Excel repeats the copy)")]
    controls = [by["control: object-model write after a Bold"], by["control: Range.PasteSpecial (legacy) after a Bold"]]
    return {
        "controls show a cleared history (measurement trustworthy)": all(c["bold_still_on"] for c in controls),
        "native paste: values correct, no formulas or formats": all(c.get("values_as_expected")
                                                                    and c["no_formulas_or_formats_pasted"] for c in core),
        "native paste keeps the clipboard": all(c["cut_copy_mode_after"] == 1 for c in core)
        and by["clipboard kept: two pastes from one copy (D1 then G1)"]["both_pastes_as_expected"],
        "native Undo reverses the paste": all(c["paste_reversed_by_undo"] and c["bold_kept_after_undo_1"] for c in core),
        "earlier Undo history kept": all(c["earlier_history_reversed_by_undo_2"] for c in core),
        "multi-area destinations": "see scenarios (behaviour recorded, decision for review)",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--visible", action="store_true")
    parser.add_argument("--out", type=Path, default=HERE / "results" / f"m3d_native_paste_{time.strftime('%Y%m%d_%H%M%S')}.json")
    args = parser.parse_args()
    personal_before = sha256(PERSONAL)
    experiment = Experiment(args.visible)
    error = None
    try:
        experiment.run_all()
    except Exception as exc:  # noqa: BLE001 - recorded; the disposable Excel is always closed
        error = repr(exc)
        print("STOPPED:", error)
    finally:
        fired = experiment.watchdog.fired
        experiment.stop()
        set_clipboard_text(None)
    summary = summarise(experiment.results) if error is None else {}
    out = {"visible": args.visible, "error": error, "watchdog_fired": fired, "summary": summary,
           "personal_unchanged": sha256(PERSONAL) == personal_before, "results": experiment.results}
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(out, indent=2, default=str), encoding="utf-8")
    print(json.dumps(summary, indent=2))
    print("PERSONAL.XLSB unchanged:", out["personal_unchanged"], "->", args.out)
    return 0 if error is None else 1


if __name__ == "__main__":
    sys.exit(main())
