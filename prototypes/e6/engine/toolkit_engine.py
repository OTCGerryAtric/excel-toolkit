"""E6 packaged engine (prototype): one executable, two modes.

  ExcelToolkitEngine.exe worker  <state_dir> [--idle-seconds N]
  ExcelToolkitEngine.exe oneshot <request_file>

Read-only audit requests only. Hardened from E4:
- named pipe restricted to the current Windows user (explicit DACL) and local clients
- bounded message size, strict request validation, bounded request size
- single worker per user; idle timeout; pid file for clean-up
"""

import argparse
import ctypes
import gc
import json
import os
import queue
import re
import sys
import threading
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32api
import win32file
import win32pipe
import win32security

from excel_link import ExcelBusy, ExcelGone, Identity, IdentityError, attach, call

VERSION = "0.0.0"  # replaced at build time
CALLBACK = "ExcelToolkit.xlam!Toolkit_OnResult"
MAX_MESSAGE_BYTES = 16 * 1024
MAX_CELLS = 1_000_000
RECYCLE_AFTER_CELLS = int(os.environ.get("EXCEL_TOOLKIT_RECYCLE_CELLS", "10000000"))  # backstop: fresh process if latency degrades
ALLOWED_OPS = {"classify", "read"}
ADDRESS = re.compile(r"^\$?[A-Z]{1,3}\$?\d{1,7}(:\$?[A-Z]{1,3}\$?\d{1,7})?(,\$?[A-Z]{1,3}\$?\d{1,7}(:\$?[A-Z]{1,3}\$?\d{1,7})?)*$")
REQUEST_ID = re.compile(r"^[A-Za-z0-9_-]{1,64}$")
PIPE_ACCESS_INBOUND = 0x00000001
FILE_FLAG_FIRST_PIPE_INSTANCE = 0x00080000
PIPE_REJECT_REMOTE_CLIENTS = 0x00000008
ERROR_NAMES = {
    2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?",
    2036: "#NUM!", 2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!", 2046: "#CONNECT!",
    2047: "#BLOCKED!", 2048: "#UNKNOWN!", 2049: "#FIELD!", 2050: "#CALC!",
}
COM_ERROR_BASE = -2146828288
NUMBER_LITERAL = re.compile(r'(?<![A-Za-z_$\d.])\d+(\.\d+)?(?![\d(:A-Za-z_])')


# ---------- Validation ----------

class InvalidRequest(Exception):
    pass


def validate(request) -> dict:
    if not isinstance(request, dict):
        raise InvalidRequest("request is not an object")
    checks = [
        ("id", str, lambda v: REQUEST_ID.match(v)),
        ("op", str, lambda v: v in ALLOWED_OPS),
        ("sheet", str, lambda v: 0 < len(v) <= 31),
        ("address", str, lambda v: len(v) <= 2048 and ADDRESS.match(v.upper())),
        ("hwnd", int, lambda v: v > 0),
        ("pid", int, lambda v: v > 0),
        ("workbook", str, lambda v: 0 < len(v) <= 260),
        ("full_name", str, lambda v: 0 < len(v) <= 1024),
    ]
    for key, kind, ok in checks:
        value = request.get(key)
        if not isinstance(value, kind) or isinstance(value, bool) or not ok(value):
            raise InvalidRequest(f"field {key!r} is missing or invalid")
    extra = set(request) - {key for key, _, _ in checks}
    if extra:
        raise InvalidRequest(f"unexpected fields {sorted(extra)}")
    return request


# ---------- Reading and classifying ----------

def as_2d(value) -> list[list]:
    if isinstance(value, (tuple, list)):
        return [list(row) for row in value]
    return [[value]]


def clean(value):
    if isinstance(value, int) and not isinstance(value, bool):
        code = value - COM_ERROR_BASE
        if 2000 <= code <= 2100:
            return {"error": ERROR_NAMES.get(code, f"#ERR{code}")}
    return value


def column_letters(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def top_left(address: str) -> tuple[int, int]:
    match = re.match(r"\$?([A-Z]+)\$?(\d+)", address.upper())
    column = 0
    for char in match.group(1):
        column = column * 26 + ord(char) - 64
    return int(match.group(2)), column


def area_size(address: str) -> int:
    first, _, last = address.upper().replace("$", "").partition(":")
    r1, c1 = top_left(first)
    r2, c2 = top_left(last) if last else (r1, c1)
    return (abs(r2 - r1) + 1) * (abs(c2 - c1) + 1)


def classify(sheet: str, address: str, formulas, values) -> list[tuple]:
    top, left = top_left(address)
    findings = []
    for r, row in enumerate(formulas):
        for c, formula in enumerate(row):
            value = values[r][c]
            if isinstance(formula, str) and formula.startswith("="):
                without_strings = re.sub(r'"[^"]*"', "", formula)
                kind = ("formula_error" if isinstance(value, dict)
                        else "formula_with_constant" if NUMBER_LITERAL.search(without_strings) else "formula")
                detail = formula
            elif value is None or value == "":
                continue
            elif isinstance(value, dict):
                kind, detail = "error_constant", value["error"]
            elif isinstance(value, bool):
                kind, detail = "boolean", str(value)
            elif isinstance(value, (int, float)):
                kind, detail = "number", str(value)
            else:
                kind, detail = "text", str(value)
            findings.append((sheet, f"{column_letters(left + c)}{top + r}", kind, detail))
    return findings


def identity_of(request: dict) -> Identity:
    return Identity.from_handoff(f"{request['hwnd']}|{request['pid']}|{request['workbook']}|{request['full_name']}")


def handle(request: dict) -> dict:
    response = {"id": request["id"], "status": "ok", "error": "", "probe": None, "count": 0, "findings": [], "timings": {}}
    started = time.perf_counter()
    app = workbook = None
    try:
        areas = request["address"].split(",")
        if sum(area_size(a) for a in areas) > MAX_CELLS:
            raise InvalidRequest(f"request covers more than {MAX_CELLS:,} cells")
        app, workbook = attach(identity_of(request))
        sheet = call(app, lambda: workbook.Worksheets(request["sheet"]))
        for area in areas:
            rng = call(app, lambda: sheet.Range(area))
            formulas = as_2d(call(app, lambda: rng.Formula2))
            values = [[clean(v) for v in row] for row in as_2d(call(app, lambda: rng.Value2))]
            if request["op"] == "classify":
                response["findings"] += classify(request["sheet"], area, formulas, values)
            else:
                response["count"] += sum(len(row) for row in formulas)
        if request["op"] == "classify":
            response["count"] = len(response["findings"])
        response["probe"] = call(app, lambda: workbook.Worksheets(1).Range("A1").Value2)
    except (IdentityError, ExcelBusy, ExcelGone, InvalidRequest) as error:
        response.update(status="failed", error=f"{type(error).__name__}: {error}")
    except Exception as error:  # noqa: BLE001
        response.update(status="failed", error=f"{type(error).__name__}: {error}")
    finally:
        app = workbook = sheet = rng = None  # noqa: F841 - short-lived COM connection
        gc.collect()
        pythoncom.CoFreeUnusedLibraries()
    response["timings"]["work_ms"] = (time.perf_counter() - started) * 1000
    return response


def deliver(request: dict, response: dict) -> str:
    app = None
    try:
        app, _ = attach(identity_of(request))
        findings = tuple(response["findings"]) if response["findings"] else ""
        return call(app, lambda: app.Run(
            CALLBACK, response["id"], request["pid"], request["workbook"], request["full_name"],
            response["status"], response["error"], response["probe"], response["count"], findings))
    except Exception as error:  # noqa: BLE001
        return f"not delivered: {type(error).__name__}: {error}"
    finally:
        app = None  # noqa: F841
        gc.collect()
        pythoncom.CoFreeUnusedLibraries()


# ---------- Worker ----------

def pipe_name() -> str:
    return rf"\\.\pipe\excel_toolkit_{os.environ.get('USERNAME', 'user')}"


def user_only_security():
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32security.TOKEN_QUERY)
    sid = win32security.ConvertSidToStringSid(win32security.GetTokenInformation(token, win32security.TokenUser)[0])
    attributes = win32security.SECURITY_ATTRIBUTES()
    attributes.SECURITY_DESCRIPTOR = win32security.ConvertStringSecurityDescriptorToSecurityDescriptor(
        f"D:P(A;;GA;;;{sid})", win32security.SDDL_REVISION_1)  # protected DACL: this user only
    return attributes


def create_pipe(first: bool):
    return win32pipe.CreateNamedPipe(
        pipe_name(),
        PIPE_ACCESS_INBOUND | (FILE_FLAG_FIRST_PIPE_INSTANCE if first else 0),
        win32pipe.PIPE_TYPE_BYTE | win32pipe.PIPE_READMODE_BYTE | win32pipe.PIPE_WAIT | PIPE_REJECT_REMOTE_CLIENTS,
        win32pipe.PIPE_UNLIMITED_INSTANCES, 0, MAX_MESSAGE_BYTES, 0, user_only_security(),
    )


ERROR_NO_DATA = 232          # client already wrote and closed before we called ConnectNamedPipe
ERROR_PIPE_CONNECTED = 535   # client connected before we called ConnectNamedPipe


def accept_loop(first_pipe, inbox: queue.Queue, log) -> None:
    """Never lets one bad connection stop the listener."""
    pipe = first_pipe
    while True:
        try:
            win32pipe.ConnectNamedPipe(pipe, None)
        except pywintypes.error as error:
            if error.winerror not in (ERROR_NO_DATA, ERROR_PIPE_CONNECTED):
                log({"event": "connect error", "error": str(error)})
        next_pipe = create_pipe(first=False)
        chunks, size, too_big = [], 0, False
        try:
            while True:
                _, data = win32file.ReadFile(pipe, 4096)
                size += len(data)
                if size > MAX_MESSAGE_BYTES:
                    too_big = True
                    break
                chunks.append(data)
        except pywintypes.error:
            pass  # end of message: client closed its end
        try:
            win32pipe.DisconnectNamedPipe(pipe)
            win32file.CloseHandle(pipe)
        except pywintypes.error:
            pass
        pipe = next_pipe
        if too_big:
            log({"event": "refused", "reason": f"message larger than {MAX_MESSAGE_BYTES} bytes"})
            continue
        text = b"".join(chunks).decode("mbcs", errors="replace").strip()
        if text:
            inbox.put(text)


class _PowerThrottlingState(ctypes.Structure):
    _fields_ = [("Version", ctypes.c_ulong), ("ControlMask", ctypes.c_ulong), ("StateMask", ctypes.c_ulong)]


def opt_out_of_power_throttling() -> dict:
    """A hidden, windowless process started by Excel can be put into Windows 11 'efficiency mode'
    (EcoQoS), which slowed the worker ~3x in E4/E6. Opt out explicitly; report what we found."""
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.GetCurrentProcess.restype = ctypes.c_void_p
    kernel32.SetProcessInformation.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p, ctypes.c_ulong]
    kernel32.GetPriorityClass.argtypes = [ctypes.c_void_p]
    process = kernel32.GetCurrentProcess()
    state = _PowerThrottlingState(1, 0x1, 0x0)  # control EXECUTION_SPEED, state off = not throttled
    ok = bool(kernel32.SetProcessInformation(process, 4, ctypes.byref(state), ctypes.sizeof(state)))  # 4 = ProcessPowerThrottling
    return {"power_throttling_opt_out": ok, "priority_class": hex(kernel32.GetPriorityClass(process))}


def run_worker(state: Path, idle_seconds: float) -> int:
    state.mkdir(parents=True, exist_ok=True)
    throttling = opt_out_of_power_throttling() if os.environ.get("EXCEL_TOOLKIT_ALLOW_THROTTLING") != "1" else {"power_throttling_opt_out": "skipped"}
    try:
        first_pipe = create_pipe(first=True)
    except pywintypes.error:
        return 0  # another worker already owns the pipe
    (state / "worker.pid").write_text(f"{os.getpid()}", encoding="ascii")
    log_path = state / "worker_log.jsonl"

    def log(entry: dict) -> None:
        entry["at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        with log_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, default=str) + "\n")

    inbox: queue.Queue = queue.Queue()
    threading.Thread(target=accept_loop, args=(first_pipe, inbox, log), daemon=True).start()
    log({"event": "started", "pid": os.getpid(), "version": VERSION, **throttling})
    pythoncom.CoInitialize()
    cells_served = 0
    try:
        while True:
            idle_since = time.monotonic()
            text = None
            while text is None:
                try:
                    text = inbox.get(timeout=1)
                except queue.Empty:
                    pythoncom.PumpWaitingMessages()  # keep the STA's message queue drained while idle
                    if time.monotonic() - idle_since >= idle_seconds:
                        log({"event": "idle timeout"})
                        return 0
            try:
                raw = json.loads(text)
            except json.JSONDecodeError:
                log({"event": "refused", "reason": "not JSON"})
                continue
            if raw == {"op": "shutdown"}:
                log({"event": "shutdown requested"})
                return 0
            try:
                request = validate(raw)
            except InvalidRequest as error:
                log({"event": "refused", "reason": str(error)})
                continue
            started = time.perf_counter()
            response = handle(request)
            verdict = deliver(request, response)
            log({"id": request["id"], "op": request["op"], "pid": request["pid"], "workbook": request["workbook"],
                 "status": response["status"], "error": response["error"], "count": response["count"],
                 "verdict": verdict, "total_ms": (time.perf_counter() - started) * 1000})
            pythoncom.PumpWaitingMessages()  # STA thread: drain COM's window messages after each request
            cells_served += response["count"]
            if cells_served >= RECYCLE_AFTER_CELLS and inbox.empty():
                log({"event": "recycling", "cells_served": cells_served})
                return 0
    finally:
        try:
            (state / "worker.pid").unlink()
        except OSError:
            pass


def run_oneshot(request_file: Path) -> int:
    pythoncom.CoInitialize()
    try:
        request = validate(json.loads(request_file.read_text(encoding="mbcs")))
    except (InvalidRequest, json.JSONDecodeError, OSError):
        return 2
    finally:
        try:
            request_file.unlink()
        except OSError:
            pass
    response = handle(request)
    deliver(request, response)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(prog="ExcelToolkitEngine")
    sub = parser.add_subparsers(dest="mode", required=True)
    worker = sub.add_parser("worker")
    worker.add_argument("state_dir")
    worker.add_argument("--idle-seconds", type=float, default=1800)
    oneshot = sub.add_parser("oneshot")
    oneshot.add_argument("request_file")
    sub.add_parser("version")
    args = parser.parse_args()
    if args.mode == "worker":
        return run_worker(Path(args.state_dir), args.idle_seconds)
    if args.mode == "oneshot":
        return run_oneshot(Path(args.request_file))
    print(VERSION)
    return 0


if __name__ == "__main__":
    sys.exit(main())
