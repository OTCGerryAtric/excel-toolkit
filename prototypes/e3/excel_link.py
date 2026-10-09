"""E3 prototype adapter: explicit attach (from E2) plus bounded retries for busy Excel.

Never falls back to GetActiveObject / Dispatch("Excel.Application").
"""

import time
from dataclasses import dataclass

import pythoncom
import pywintypes
import win32api
import win32com.client
import win32con
import win32gui
import win32process

RPC_E_CALL_REJECTED = -2147418111        # 0x80010001: Excel busy (dialog, edit mode)
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
BUSY = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
GONE = {
    -2147023174,  # 0x800706BA RPC server unavailable
    -2147023170,  # 0x800706BE remote procedure call failed
    -2147417848,  # 0x80010108 object disconnected from its clients
    -2147418094,  # 0x80010012 server has died / call did not execute
}
WM_GETOBJECT = 0x003D
OBJID_NATIVEOM = -16
SMTO_ABORTIFHUNG = 0x0002
DEFAULT_TIMEOUT = 5.0


class ExcelBusy(Exception):
    pass


class ExcelGone(Exception):
    pass


class IdentityError(Exception):
    pass


@dataclass(frozen=True)
class Identity:
    hwnd: int
    pid: int
    workbook: str
    full_name: str
    process_started: float

    @classmethod
    def from_handoff(cls, handoff: str) -> "Identity":
        hwnd, pid, workbook, full_name = handoff.split("|", 3)
        started = process_start_time(int(pid))
        if started is None:
            raise IdentityError(f"Excel process {pid} is not running")
        return cls(int(hwnd), int(pid), workbook, full_name, started)


def process_start_time(pid: int) -> float | None:
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    except pywintypes.error:
        return None
    if win32process.GetExitCodeProcess(handle) != 259:  # STILL_ACTIVE
        return None
    return win32process.GetProcessTimes(handle)["CreationTime"].timestamp()


def _classify(app, error: Exception) -> str:
    """'busy', 'gone' or 'other' for a COM failure (pywin32 sometimes reports a rejected
    call as AttributeError, so probe the application to find out)."""
    if isinstance(error, pywintypes.com_error):
        hresult = error.hresult
    else:
        try:
            app._oleobj_.GetIDsOfNames("Ready")
            return "other"
        except pywintypes.com_error as probe:
            hresult = probe.hresult
    if hresult in BUSY:
        return "busy"
    if hresult in GONE:
        return "gone"
    return "other"


def call(app, operation, timeout: float = DEFAULT_TIMEOUT):
    """Run one COM operation, retrying while Excel is busy, for at most `timeout` seconds."""
    deadline = time.monotonic() + timeout
    while True:
        try:
            return operation()
        except (pywintypes.com_error, AttributeError) as error:
            kind = _classify(app, error)
            if kind == "busy":
                if time.monotonic() >= deadline:
                    raise ExcelBusy(
                        f"Excel stayed busy for {timeout:.0f} s - a dialog may be open or a cell is being edited. "
                        "Finish or cancel it in Excel and try again."
                    ) from error
                time.sleep(0.1)
                continue
            if kind == "gone":
                raise ExcelGone("Excel closed or stopped responding during the operation.") from error
            raise


def attach(identity: Identity, timeout: float = DEFAULT_TIMEOUT):
    """Return (app, workbook) for exactly this identity, or raise."""
    started = process_start_time(identity.pid)
    if started is None:
        raise IdentityError(f"Excel process {identity.pid} is no longer running")
    if abs(started - identity.process_started) > 1:
        raise IdentityError(f"process id {identity.pid} now belongs to a different process (stale identity)")
    if not win32gui.IsWindow(identity.hwnd) or win32process.GetWindowThreadProcessId(identity.hwnd)[1] != identity.pid:
        raise IdentityError(f"window {identity.hwnd} no longer belongs to Excel process {identity.pid}")

    desk = win32gui.FindWindowEx(identity.hwnd, 0, "XLDESK", None)
    workbook_window = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
    if not workbook_window:
        raise IdentityError("no workbook window under the Excel window")
    deadline = time.monotonic() + timeout
    while True:
        try:
            _, lresult = win32gui.SendMessageTimeout(
                workbook_window, WM_GETOBJECT, 0, OBJID_NATIVEOM, SMTO_ABORTIFHUNG, int(timeout * 1000)
            )
        except pywintypes.error:
            lresult = 0
        if lresult:
            break
        if time.monotonic() >= deadline:
            raise ExcelBusy(f"Excel did not respond within {timeout:.0f} s")
        time.sleep(0.1)
    window = win32com.client.Dispatch(pythoncom.ObjectFromLresult(lresult, pythoncom.IID_IDispatch, 0))
    app = call(window, lambda: window.Application, timeout)
    app_pid = win32process.GetWindowThreadProcessId(int(call(app, lambda: app.Hwnd, timeout)))[1]
    if app_pid != identity.pid:
        raise IdentityError(f"attached to Excel process {app_pid}, expected {identity.pid}")

    def find():
        for workbook in app.Workbooks:
            if workbook.Name == identity.workbook:
                return workbook
        return None

    workbook = call(app, find, timeout)
    if workbook is None:
        raise IdentityError(f"workbook {identity.workbook!r} is no longer open (closed or renamed)")
    if call(app, lambda: workbook.FullName, timeout) != identity.full_name:
        raise IdentityError(f"workbook {identity.workbook!r} is now a different file")
    return app, workbook
