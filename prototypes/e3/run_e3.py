"""E3 orchestrator: reversibility by command category, and connection resilience.

All changes happen in disposable copies opened in new Excel processes. The owner's own
Excel is never contacted. The orchestrator's own COM reference plays the part of the
user (setup edits, renames, closing workbooks); the code under test always attaches
through an explicit identity.
"""

import json
import shutil
import sys
import threading
import time
import traceback
from pathlib import Path

import pywintypes
import win32api
import win32com.client
import win32con
import win32gui
import win32process

sys.path.insert(0, str(Path(__file__).resolve().parent))
from excel_link import ExcelBusy, ExcelGone, Identity, IdentityError, attach, call, process_start_time  # noqa: E402
from journal import (  # noqa: E402
    OperationFailed, Journal, delete_names, preview_delete_names, read_block,
    run_block_operation, undo_blocks, undo_delete_names, value_input,
)

HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"
WORK = HERE / "work"
JOURNALS = WORK / "journals"
RESULTS = HERE / "results"
XL_OPEN_XML_ADDIN = 55
XL_OPEN_XML_WORKBOOK = 51
WM_KEYDOWN, WM_KEYUP, WM_CHAR, WM_COMMAND, WM_CLOSE = 0x0100, 0x0101, 0x0102, 0x0111, 0x0010
VK_F2, VK_ESCAPE, IDOK = 0x71, 0x1B, 1

results: list[dict] = []


def record(test: str, passed, detail: str, category: str) -> None:
    passed = bool(passed)
    results.append({"test": test, "category": category, "passed": passed, "detail": detail})
    print(f"[{'PASS' if passed else 'FAIL'}] {test}: {detail}")


def guarded(test: str, category: str):
    def wrap(fn):
        try:
            fn()
        except Exception as error:  # noqa: BLE001 - record and carry on with the other tests
            record(test, False, f"unexpected {type(error).__name__}: {error}\n{traceback.format_exc(limit=3)}", category)
    return wrap


# ---------- Excel helpers ----------

def start_excel():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


def pid_of(excel) -> int:
    return win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1]


def build_xlam(target: Path) -> None:
    crlf = WORK / "E3Commands.bas"
    crlf.write_bytes((HERE / "E3Commands.bas").read_text(encoding="ascii").replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Add()
        workbook.VBProject.VBComponents.Import(str(crlf))
        workbook.SaveAs(str(target), FileFormat=XL_OPEN_XML_ADDIN)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()


def identity_for(excel, workbook) -> Identity:
    workbook.Activate()
    return Identity.from_handoff(excel.Run("E3Commands.xlam!E3_Identity"))


def dialogs_of(pid: int) -> list[int]:
    found = []

    def visit(hwnd, _):
        cls = win32gui.GetClassName(hwnd)
        if (cls == "#32770" or cls.startswith("bosa_sdm")) and win32gui.IsWindowVisible(hwnd) \
                and win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
            found.append(hwnd)
        return True

    win32gui.EnumWindows(visit, None)
    return found


def close_dialogs(pid: int) -> int:
    hwnds = dialogs_of(pid)
    for hwnd in hwnds:
        if win32gui.GetClassName(hwnd) == "#32770":
            win32gui.PostMessage(hwnd, WM_COMMAND, IDOK, 0)
        else:
            win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
    return len(hwnds)


def wait_until(condition, timeout: float) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if condition():
            return True
        time.sleep(0.05)
    return False


def values_transform(before):
    return [[value_input(cell[1]) for cell in row] for row in before]


# ---------- Tests ----------

def test_toggle_ctrl_z(x, wb) -> None:
    category = "Formatting toggle: Excel Ctrl+Z via Application.OnUndo"
    data = wb.Worksheets("Data")
    data.Activate()
    data.Range("C1:C5,B5").Select()
    before = [data.Range(a).NumberFormat for a in ("C1", "C3", "B5")]
    x.Run("E3Commands.xlam!E3_Number_Format_Toggle")
    toggled = [data.Range(a).NumberFormat for a in ("C1", "C3", "B5")]
    record("Toggle changes mixed multi-area selection to the first format", len(set(toggled)) == 1 and toggled != before,
           f"before {before} -> {toggled}", category)

    for how, undo in (("Application.Undo called from automation", x.Undo),
                      ("Application.Undo called inside VBA", lambda: x.Run("E3Commands.xlam!E3_CallUndoFromVba"))):
        x.Run("E3Commands.xlam!E3_Number_Format_Toggle") if how.endswith("VBA") else None
        try:
            undo()
            outcome = "no error"
        except pywintypes.com_error as error:
            outcome = f"error {error.excepinfo[2] if error.excepinfo else error}"
        after_undo = [data.Range(a).NumberFormat for a in ("C1", "C3", "B5")]
        record(f"Informational: {how} triggers the OnUndo restore", True,
               f"restored: {after_undo == before} ({outcome}); formats now {after_undo}", category)
        data.Range("C1:C5").NumberFormat = "#,##0.0"
        data.Range("B5").NumberFormat = "General"
        data.Range("C1:C5,B5").Select()

    x.Run("E3Commands.xlam!E3_RegisterUndo", "Undo paste as values")
    try:
        x.Run("E3Commands.xlam!E3_CallUndoFromVba")
    except pywintypes.com_error:
        pass
    marker = x.Run("E3Commands.xlam!E3_GetUndoMarker")
    record("Informational: OnUndo hook registered for a Python operation fires on Application.Undo", True,
           f"marker after Undo: {marker!r} - real Ctrl+Z needs a manual keypress test", category)


def test_bulk(x, wb) -> Path | None:
    category = "Bulk transformation: custom undo"
    data = wb.Worksheets("Data")

    # What a naive 'Value = Value' paste-as-values does to text constants, on scratch cells.
    data.Range("F1").Value = "'0012"
    data.Range("F2").Value = "'1/2"
    data.Range("F1:F2").Value2 = data.Range("F1:F2").Value2
    naive = [data.Range("F1").Value2, data.Range("F2").Value2]
    data.Range("F1:F2").Clear()
    record("Evidence: naive Value2 = Value2 re-parses text constants", naive != ["0012", "1/2"],
           f"text '0012' and '1/2' became {naive!r}", category)

    identity = identity_for(x, wb)
    blocks = [("Data", "A1:A5"), ("Data", "B1:B5"), ("Data", "C1:C5")]
    formats_before = [data.Range(f"C{r}").NumberFormat for r in range(1, 6)]
    journal = run_block_operation(identity, "paste_as_values", blocks, values_transform, JOURNALS)
    after = {b["address"]: b["after"] for b in journal.record["blocks"]}
    checks = {
        "A1 now a value (10)": after["A1:A5"][0][0] == ["10", 10.0] or after["A1:A5"][0][0][1] == 10.0 and not str(after["A1:A5"][0][0][0]).startswith("="),
        "B1 text '0012' kept as text": after["B1:B5"][0][0][1] == "0012",
        "B2 text '1/2' kept as text": after["B1:B5"][1][0][1] == "1/2",
        "B3 #N/A kept as error constant": after["B1:B5"][2][0][1] == {"error": "#N/A"} and not str(after["B1:B5"][2][0][0]).startswith("="),
        "B4 TRUE kept as boolean": after["B1:B5"][3][0][1] is True,
        "C formats kept": [data.Range(f"C{r}").NumberFormat for r in range(1, 6)] == formats_before,
    }
    record("Paste as values: values, text, errors, booleans and formats correct", all(checks.values()),
           json.dumps(checks), category)

    outcome = undo_blocks(journal.path)
    app, workbook = attach(identity)
    restored = all(read_block(app, workbook, b["sheet"], b["address"]) == b["before"]
                   for b in Journal.load(journal.path).record["blocks"])
    record("Custom undo restores formulas and constants exactly", outcome["ok"] and restored,
           f"{outcome}; verified equal to before: {restored}; formula A1 now {data.Range('A1').Formula2!r}", category)
    return journal.path


def test_conflict(x, wb) -> None:
    category = "Bulk transformation: conflict detection"
    identity = identity_for(x, wb)
    journal = run_block_operation(identity, "paste_as_values", [("Data", "A1:A5")], values_transform, JOURNALS)
    wb.Worksheets("Data").Range("A2").Value = 777  # the user edits an affected cell afterwards
    outcome = undo_blocks(journal.path)
    untouched = wb.Worksheets("Data").Range("A1").Formula2
    record("Undo refused when an affected cell changed since; conflicting cell reported", not outcome["ok"]
           and outcome.get("conflicts") == ["Data!$A$2"], f"{outcome}", category)
    record("Nothing restored after a refused undo (no partial overwrite)", untouched == "10" and
           wb.Worksheets("Data").Range("A2").Value2 == 777, f"A1 formula {untouched!r}, A2 {wb.Worksheets('Data').Range('A2').Value2}", category)
    # Put the sheet back for later tests.
    wb.Worksheets("Data").Range("A1:A5").Formula = "=ROW()*10"


def test_partial_failure(x, wb) -> None:
    category = "Bulk transformation: rollback after partial failure"
    prot = wb.Worksheets("Prot")
    before = [prot.Range("A1:A3").Formula2, prot.Range("C1:C3").Formula2]
    identity = identity_for(x, wb)
    try:
        run_block_operation(identity, "paste_as_values", [("Prot", "A1:A3"), ("Prot", "C1:C3")], values_transform, JOURNALS)
        record("Operation over unlocked + locked cells fails", False, "no error raised", category)
        return
    except OperationFailed as error:
        message = str(error)
    journal_path = max(JOURNALS.glob("*_paste_as_values.json"), key=lambda p: p.stat().st_mtime)
    journal = Journal.load(journal_path)
    after = [prot.Range("A1:A3").Formula2, prot.Range("C1:C3").Formula2]
    record("First block written, second refused by sheet protection, first block rolled back",
           journal.record["status"] == "rolled_back" and after == before,
           f"status {journal.record['status']}; events {journal.record['events'][-2:]}; error: {message[:120]}", category)


def test_names(x, wb) -> None:
    category = "Destructive: delete names"
    identity = identity_for(x, wb)
    preview = preview_delete_names(identity)
    record("Preview lists every name, including hidden, sheet-scoped, print area and broken",
           {"Rate", "HiddenName", "Broken", "Data!LocalName", "Data!Print_Area"} <= {n["full"] for n in preview},
           f"{[(n['full'], n['visible']) for n in preview]}", category)
    try:
        delete_names(identity, JOURNALS, confirmed=False)
        record("Refuses without confirmation", False, "ran without confirmation", category)
    except OperationFailed as error:
        record("Refuses without confirmation", True, str(error), category)

    try:
        delete_names(identity, JOURNALS, confirmed=True, fail_after=2)
    except OperationFailed as error:
        after = preview_delete_names(identity)
        record("Failure after 2 deletions rolls back to the identical name set", after == preview, str(error), category)

    journal = delete_names(identity, JOURNALS, confirmed=True)
    remaining = preview_delete_names(identity)
    record("Delete names removes them all", remaining == [], f"remaining {remaining}", category)
    outcome = undo_delete_names(journal.path)
    restored = preview_delete_names(identity)
    record("Undo restores every name with scope, definition and visibility", outcome["ok"] and restored == preview,
           f"{outcome}; differences: {[n for n in preview if n not in restored]}", category)


def test_book1(x) -> None:
    category = "Identity: unsaved new workbook"
    book = x.Workbooks.Add()
    book.Worksheets(1).Range("A1:A3").Formula = "=ROW()*5"
    identity = identity_for(x, book)
    record("Unsaved workbook has no path; identity uses its name", "\\" not in identity.full_name and book.Path == "",
           f"name {identity.workbook!r}, full name {identity.full_name!r}, Path {book.Path!r}", category)
    sheet = book.Worksheets(1).Name
    journal = run_block_operation(identity, "paste_as_values", [(sheet, "A1:A3")], values_transform, JOURNALS)
    outcome = undo_blocks(journal.path)
    record("Operation and undo work on an unsaved workbook", outcome["ok"] and book.Worksheets(1).Range("A2").Formula2 == "=ROW()*5",
           f"{outcome}", category)
    book.Close(SaveChanges=False)


def test_rename_and_reopen(x) -> None:
    category = "Identity: renamed, closed and reopened workbook"
    original = WORK / "x" / "E3_Rename.xlsx"
    shutil.copy(FIXTURES / "E3_Fixture.xlsx", original)
    book = x.Workbooks.Open(str(original))
    identity = identity_for(x, book)
    journal = run_block_operation(identity, "paste_as_values", [("Data", "A1:A5")], values_transform, JOURNALS)

    book.SaveAs(str(WORK / "x" / "E3_Renamed.xlsx"), FileFormat=XL_OPEN_XML_WORKBOOK)
    try:
        outcome = undo_blocks(journal.path)
        record("Undo after Save As under a new name is refused", False, f"undo ran: {outcome}", category)
    except IdentityError as error:
        record("Undo after Save As under a new name is refused", True, str(error), category)

    book.Close(SaveChanges=False)
    reopened = x.Workbooks.Open(str(original))  # same name and path; disk version still has formulas
    outcome = undo_blocks(journal.path)
    record("Undo against a closed-and-reopened workbook (same name/path) is refused by the content check",
           not outcome["ok"] and outcome.get("conflicts"), f"{outcome}", category)
    reopened.Close(SaveChanges=False)

    try:
        attach(identity)
        record("Stale identity after close is refused", False, "attached", category)
    except IdentityError as error:
        record("Stale identity after close is refused", True, str(error), category)


def test_stale_process(identity: Identity) -> None:
    category = "Identity: stale process"
    stale = Identity(identity.hwnd, identity.pid, identity.workbook, identity.full_name, identity.process_started - 3600)
    try:
        attach(stale)
        record("Identity whose process start time doesn't match (pid reused) is refused", False, "attached", category)
    except IdentityError as error:
        record("Identity whose process start time doesn't match (pid reused) is refused", True, str(error), category)


def test_busy(x, wb, x_pid: int) -> None:
    category = "Resilience: Excel busy"
    identity = identity_for(x, wb)

    def probe(timeout: float) -> tuple[str, float]:
        """Read Data!A1 and write a scratch cell; report what happened and how long it took."""
        started = time.monotonic()
        try:
            app, workbook = attach(identity, timeout=timeout)
            read_block(app, workbook, "Data", "A1")
            sheet = call(app, lambda: workbook.Worksheets("Data"), timeout)
            call(app, lambda: setattr(sheet.Range("H1"), "Value2", 1), timeout)
            return "read and write succeeded", time.monotonic() - started
        except ExcelBusy as error:
            return f"ExcelBusy: {error}", time.monotonic() - started

    for label, macro in (("VBA MsgBox", "E3_ScheduleBlockingMessage"), ("built-in dialog (Format Number)", "E3_ScheduleBuiltInDialog")):
        x.Run(f"E3Commands.xlam!{macro}")
        opened = wait_until(lambda: dialogs_of(x_pid), 10)
        if not opened:
            record(f"{label}: dialog opened", False, "no dialog window found", category)
            continue
        outcome, elapsed = probe(5)
        blocked = outcome.startswith("ExcelBusy")
        record(f"{label} open: outcome of read + write with a 5 s bounded wait", True if not blocked else 4.5 <= elapsed <= 8,
               f"{outcome} after {elapsed:.1f} s", category)
        if blocked:
            timer = threading.Timer(2.0, close_dialogs, args=(x_pid,))
            timer.start()
            outcome, elapsed = probe(5)
            timer.join()
            record(f"{label}: succeeds when the dialog closes within the retry window", outcome == "read and write succeeded",
                   f"{outcome} after {elapsed:.1f} s", category)
        close_dialogs(x_pid)
        wait_until(lambda: not dialogs_of(x_pid), 5)

    # Cell edit mode, entered by messages posted only to the target workbook window.
    desk = win32gui.FindWindowEx(int(x.Hwnd), 0, "XLDESK", None)
    workbook_window = win32gui.FindWindowEx(desk, 0, "EXCEL7", None)
    win32gui.PostMessage(workbook_window, WM_KEYDOWN, VK_F2, 0)
    win32gui.PostMessage(workbook_window, WM_KEYUP, VK_F2, 0)
    win32gui.PostMessage(workbook_window, WM_CHAR, ord("9"), 0)
    time.sleep(0.5)
    started = time.monotonic()
    try:
        app, workbook = attach(identity, timeout=3)
        value = read_block(app, workbook, "Data", "A1")
        record("Cell edit mode (informational)", True,
               f"read succeeded in {time.monotonic() - started:.1f} s ({value}); edit mode was probably not entered via posted keys - verify manually", category)
    except ExcelBusy as error:
        record("Read in cell edit mode fails after the bounded wait with a clear message", True,
               f"{time.monotonic() - started:.1f} s: {error}", category)
    win32gui.PostMessage(workbook_window, WM_KEYDOWN, VK_ESCAPE, 0)
    win32gui.PostMessage(workbook_window, WM_KEYUP, VK_ESCAPE, 0)
    responsive = wait_until(lambda: _responsive(identity), 10)
    record("Excel responsive again and the edited cell unchanged after Escape", responsive and wb.Worksheets("Data").Range("A1").Formula2 == "=ROW()*10",
           f"responsive {responsive}, A1 {wb.Worksheets('Data').Range('A1').Formula2!r}", category)


def _responsive(identity: Identity) -> bool:
    try:
        app, workbook = attach(identity, timeout=1)
        read_block(app, workbook, "Data", "A1")
        return True
    except (ExcelBusy, IdentityError):
        return False


def test_excel_dies_mid_operation() -> None:
    category = "Resilience: Excel closes during an operation"
    shutil.copy(FIXTURES / "E3_Big.xlsx", WORK / "y" / "E3_Big.xlsx")
    y = start_excel()
    book = y.Workbooks.Open(str(WORK / "y" / "E3_Big.xlsx"))
    y_pid = pid_of(y)
    identity = Identity.from_handoff(f"{int(y.Hwnd)}|{y_pid}|{book.Name}|{book.FullName}")
    del book

    def kill():
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, y_pid)
        win32api.TerminateProcess(handle, 1)

    timer = threading.Timer(1.0, kill)
    started = time.monotonic()
    timer.start()
    try:
        run_block_operation(identity, "paste_as_values_slow", [("Data", "A1:C1000")], values_transform, JOURNALS, cell_by_cell=True)
        record("Operation interrupted by Excel closing", False, "operation completed before Excel was closed", category)
    except ExcelGone as error:
        elapsed = time.monotonic() - started
        journal_path = max(JOURNALS.glob("*_paste_as_values_slow.json"), key=lambda p: p.stat().st_mtime)
        journal = Journal.load(journal_path)
        written = journal.record["blocks"][0].get("cells_written")
        record("Excel killed mid-operation: clear error, no hang, journal marked incomplete",
               journal.record["status"] == "incomplete" and elapsed < 15,
               f"{elapsed:.1f} s; {error}; journal status {journal.record['status']}, cells written before loss ~{written} of 3000", category)
    except Exception as error:  # noqa: BLE001
        record("Excel killed mid-operation: clear error, no hang, journal marked incomplete", False,
               f"{type(error).__name__}: {error}", category)
    timer.join()
    try:
        attach(identity)
        record("Identity of the killed Excel is refused afterwards", False, "attached", category)
    except IdentityError as error:
        record("Identity of the killed Excel is refused afterwards", True, str(error), category)
    del y


# ---------- Main ----------

def main() -> None:
    shutil.rmtree(WORK, ignore_errors=True)
    for folder in (WORK / "x", WORK / "y", JOURNALS, RESULTS):
        folder.mkdir(parents=True, exist_ok=True)
    shutil.copy(FIXTURES / "E3_Fixture.xlsx", WORK / "x" / "E3_Fixture.xlsx")
    build_xlam(WORK / "E3Commands.xlam")

    x = start_excel()
    x_pid = pid_of(x)
    try:
        wb = x.Workbooks.Open(str(WORK / "x" / "E3_Fixture.xlsx"))
        x.Workbooks.Open(str(WORK / "E3Commands.xlam"))

        guarded("Formatting toggle Ctrl+Z", "toggle")(lambda: test_toggle_ctrl_z(x, wb))
        guarded("Bulk paste as values + undo", "bulk")(lambda: test_bulk(x, wb))
        guarded("Conflict detection", "bulk")(lambda: test_conflict(x, wb))
        guarded("Partial failure rollback", "bulk")(lambda: test_partial_failure(x, wb))
        guarded("Delete names", "destructive")(lambda: test_names(x, wb))
        guarded("Unsaved Book1", "identity")(lambda: test_book1(x))
        guarded("Rename / reopen", "identity")(lambda: test_rename_and_reopen(x))
        guarded("Stale process", "identity")(lambda: test_stale_process(identity_for(x, wb)))
        guarded("Busy Excel", "busy")(lambda: test_busy(x, wb, x_pid))
        guarded("Excel dies mid-operation", "gone")(test_excel_dies_mid_operation)
    finally:
        close_dialogs(x_pid)
        try:
            for workbook in list(x.Workbooks):
                workbook.Close(SaveChanges=False)
            x.Quit()
        except pywintypes.com_error:
            pass
        del x
        if not wait_until(lambda: process_start_time(x_pid) is None, 15):
            # Our own disposable test instance only (pid recorded at start-up).
            win32api.TerminateProcess(win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, x_pid), 1)
            results.append({"test": "Cleanup", "category": "housekeeping", "passed": True,
                            "detail": f"test Excel {x_pid} did not quit within 15 s; terminated"})

    (RESULTS / "e3_results.json").write_text(json.dumps(results, indent=1), encoding="utf-8")
    passed = sum(r["passed"] for r in results)
    lines = [f"# E3 results: {passed}/{len(results)} passed", ""]
    category = None
    for r in results:
        if r["category"] != category:
            category = r["category"]
            lines += ["", f"## {category}", ""]
        lines.append(f"- **{'PASS' if r['passed'] else 'FAIL'}** {r['test']} — {r['detail']}")
    (RESULTS / "e3_summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"\n{passed}/{len(results)} passed")


if __name__ == "__main__":
    main()
