"""E6 runtime tests against the PACKAGED build (frozen engine + built .xlam), no install needed.

Fresh Excel processes; E4 fixtures (target + second workbook; same-named decoy in another
process). Covers: worker, one-shot, auto fallback, crash fallback, routing, pipe ACL,
request validation, idle clean-up, and the E4 slowdown investigation.
"""

import ctypes
import json
import os
import shutil
import statistics
import subprocess
import sys
import threading
import time
from pathlib import Path

import pywintypes
import win32api
import win32com.client
import win32con
import win32file
import win32pipe
import win32process
import win32security

HERE = Path(__file__).resolve().parent
VERSION = sys.argv[1] if len(sys.argv) > 1 else "0.1.0"
PAYLOAD = HERE / "build" / f"ExcelToolkit-{VERSION}" / "payload"
ENGINE = PAYLOAD / "engine" / "ExcelToolkitEngine.exe"
E4_FIXTURES = HERE.parent / "e4" / "fixtures"
WORK = HERE / "work" / time.strftime("runtime_%Y%m%d_%H%M%S")
STATE = Path(os.environ["LOCALAPPDATA"]) / "ExcelToolkit" / "state"
RESULTS = HERE / "results"
PIPE = rf"\\.\pipe\excel_toolkit_{os.environ.get('USERNAME', 'user')}"
TARGET = "E4_Fixture.xlsx"
report: dict = {"version": VERSION, "tests": [], "perf": {}}


def record(test: str, passed, detail) -> None:
    report["tests"].append({"test": test, "passed": bool(passed), "detail": detail})
    print(f"[{'PASS' if passed else 'FAIL'}] {test}: {detail}")


def alive(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    except pywintypes.error:
        return False
    return win32process.GetExitCodeProcess(handle) == 259


def terminate(pid: int) -> None:
    try:
        win32api.TerminateProcess(win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid), 1)
    except pywintypes.error:
        pass


def wait_until(condition, timeout: float, step: float = 0.02) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if condition():
            return True
        time.sleep(step)
    return False


def worker_pid() -> int | None:
    try:
        pid = int((STATE / "worker.pid").read_text(encoding="ascii"))
    except (OSError, ValueError):
        return None
    return pid if alive(pid) else None


def stop_worker() -> None:
    pid = worker_pid()
    if pid:
        send_raw(b'{"op": "shutdown"}\n')
        if not wait_until(lambda: not alive(pid), 10):
            terminate(pid)


def send_raw(data: bytes) -> bool:
    deadline = time.monotonic() + 5
    while True:  # like the VBA client: retry while the worker is between connections
        try:
            handle = win32file.CreateFile(PIPE, win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, 0, None)
            break
        except pywintypes.error:
            if time.monotonic() > deadline:
                return False
            time.sleep(0.02)
    try:
        for start in range(0, len(data), 4096):
            win32file.WriteFile(handle, data[start:start + 4096])
    except pywintypes.error:
        pass
    win32file.CloseHandle(handle)
    return True


def start_excel():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


def pid_of(excel) -> int:
    return win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1]


def parse(text: str) -> dict:
    status, count, ms, probe, error, route, workbook = (text.split("|") + [""] * 7)[:7]
    return {"status": status, "count": int(count or 0), "ms": float(ms or "nan"), "probe": probe,
            "error": error, "route": route, "workbook": workbook}


def request(excel, workbook, mode: str, address: str, sheet: str = "Bulk", timeout: float = 90) -> dict:
    request_id = excel.Run("ExcelToolkit.xlam!Toolkit_Request", mode, "classify", sheet, address, workbook.Name)
    if request_id.startswith("ERROR"):
        return {"status": "vba_error", "error": request_id, "route": "", "probe": "", "workbook": "", "count": 0, "ms": float("nan")}
    result: dict = {}

    def done() -> bool:
        nonlocal result
        result = parse(excel.Run("ExcelToolkit.xlam!Toolkit_GetResult", request_id))
        return result["status"] != "pending"

    wait_until(done, timeout, 0.01)
    return result


def log_lines() -> list[dict]:
    path = STATE / "worker_log.jsonl"
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()] if path.exists() else []


def trivial_call_ms(excel, workbook, n: int = 30) -> float:
    cell = workbook.Worksheets("Inputs").Range("A1")
    times = []
    for _ in range(n):
        started = time.perf_counter()
        _ = cell.Value2
        times.append((time.perf_counter() - started) * 1000)
    return statistics.median(times)


def process_stats(pid: int) -> dict:
    handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
    memory = win32process.GetProcessMemoryInfo(handle)
    count = ctypes.c_ulong()
    ctypes.windll.kernel32.GetProcessHandleCount(int(handle), ctypes.byref(count))
    gdi = ctypes.windll.user32.GetGuiResources(int(handle), 0)
    user = ctypes.windll.user32.GetGuiResources(int(handle), 1)
    return {"working_set_mb": round(memory["WorkingSetSize"] / 1e6, 1), "private_mb": round(memory["PagefileUsage"] / 1e6, 1),
            "handles": count.value, "gdi": gdi, "user_objects": user}


# ---------- Tests ----------

def test_pipe_acl() -> None:
    handle = win32file.CreateFile(PIPE, win32file.GENERIC_WRITE | win32con.READ_CONTROL, 0, None, win32file.OPEN_EXISTING, 0, None)
    descriptor = win32security.GetSecurityInfo(handle, win32security.SE_KERNEL_OBJECT, win32security.DACL_SECURITY_INFORMATION)
    win32file.CloseHandle(handle)
    dacl = descriptor.GetSecurityDescriptorDacl()
    aces = []
    for i in range(dacl.GetAceCount()):
        ace = dacl.GetAce(i)
        account = win32security.LookupAccountSid(None, ace[2])
        aces.append(f"{account[1]}\\{account[0]} mask=0x{ace[1]:08X}")
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32security.TOKEN_QUERY)
    me = win32security.LookupAccountSid(None, win32security.GetTokenInformation(token, win32security.TokenUser)[0])
    record("Worker pipe access list contains only the current Windows user", len(aces) == 1 and f"{me[1]}\\{me[0]}" in aces[0],
           f"ACEs: {aces}")


def test_validation(x, target) -> None:
    before = len(log_lines())
    good = {"id": "v1", "op": "classify", "sheet": "Bulk", "address": "A1", "hwnd": int(x.Hwnd), "pid": pid_of(x),
            "workbook": TARGET, "full_name": target.FullName}
    bad = [
        ("not JSON", b"this is not json\n"),
        ("unexpected field", (json.dumps({**good, "code": "import os"}) + "\n").encode()),
        ("unknown operation", (json.dumps({**good, "op": "exec"}) + "\n").encode()),
        ("malformed address", (json.dumps({**good, "address": "A1;DROP"}) + "\n").encode()),
        ("oversized message (20 KB)", b'{"id": "' + b"x" * 20000 + b'"}\n'),
    ]
    sent = [send_raw(data) for _, data in bad]
    if not all(sent):
        record("Test client could send every bad message", False, sent)
    wait_until(lambda: len(log_lines()) >= before + len(bad), 10, 0.1)
    refused = [e for e in log_lines()[before:] if e.get("event") == "refused"]
    record("Malformed, unexpected, unknown-operation, injected and oversized requests are refused (none executed)",
           len(refused) == len(bad), [e["reason"] for e in refused])
    result = request(x, target, "worker", "A1:Z100000")
    record("Request over the 1,000,000-cell limit is answered with a clear refusal", result["status"] == "failed"
           and "more than 1,000,000 cells" in result["error"], result["error"])


def test_routing(x, y, target, other, decoy) -> None:
    plan = [(x, target, "10"), (y, decoy, "77"), (x, other, "5555")] * 10
    wrong = []
    for excel, workbook, expected in plan:
        result = request(excel, workbook, "auto", "A1:B2")
        if result["status"] != "ok" or result["probe"] != expected or result["workbook"] != workbook.Name:
            wrong.append((workbook.Name, expected, result))
    record("Packaged build: 30 interleaved requests over 2 Excel processes / 3 workbooks (2 same-named) all routed correctly",
           not wrong, wrong[:2] or "all correct")


def test_oneshot(x, target) -> None:
    result = request(x, target, "oneshot", "A1:J10")
    record("Packaged one-shot path works on its own", result["status"] == "ok" and result["count"] == 100,
           f"{result['ms']:.0f} ms, route {result['route']}")


def test_crash_fallback(x, target) -> None:
    request(x, target, "auto", "A1")  # make sure a worker is running
    pid = worker_pid()
    request_id = x.Run("ExcelToolkit.xlam!Toolkit_Request", "auto", "classify", "Bulk", "A1:J1000", target.Name)
    time.sleep(0.05)
    terminate(pid)
    result: dict = {}
    wait_until(lambda: (result.update(parse(x.Run("ExcelToolkit.xlam!Toolkit_GetResult", request_id))) or result["status"] != "pending"), 60, 0.05)
    record("Worker killed mid-request: watchdog re-sends it one-shot and the user still gets the result",
           result["status"] == "ok" and result["count"] == 10000 and "worker stopped" in result["route"],
           f"{result['status']}, {result['count']} findings, route '{result['route']}', {result['ms']:.0f} ms")


def test_worker_unavailable(x, target) -> None:
    stop_worker()
    # Something else holds the pipe name and never answers: worker can't start, request must still complete.
    squatter = win32pipe.CreateNamedPipe(PIPE, 0x00000001 | 0x00080000, 0, 1, 0, 1024, 0, None)
    try:
        result = request(x, target, "auto", "A1:J10", timeout=60)
    finally:
        win32file.CloseHandle(squatter)
    record("Another process holds the pipe name and swallows the request: it still completes via one-shot fallback",
           result["status"] == "ok" and result["route"].startswith("oneshot"), f"route '{result['route']}', {result['ms']:.0f} ms")


def test_idle_cleanup() -> None:
    stop_worker()
    process = subprocess.Popen([str(ENGINE), "worker", str(STATE), "--idle-seconds", "3"])
    exited = wait_until(lambda: process.poll() is not None, 15, 0.1)
    record("Idle worker exits on its own and removes its pid file", exited and not (STATE / "worker.pid").exists(),
           f"exit code {process.poll()}, pid file present: {(STATE / 'worker.pid').exists()}")


def test_performance() -> None:
    shutil.copy(E4_FIXTURES / TARGET, WORK / "perf.xlsx")
    stop_worker()
    excel = start_excel()
    book = excel.Workbooks.Open(str(WORK / "perf.xlsx"))
    excel.Workbooks.Open(str(PAYLOAD / "ExcelToolkit.xlam"))
    excel.Run("ExcelToolkit.xlam!Toolkit_SetQuiet", True)
    pid = pid_of(excel)
    phases = []

    def phase(name: str, mode: str | None, address: str, n: int) -> None:
        latencies = [request(excel, book, mode, address)["ms"] for _ in range(n)] if mode else []
        small = [request(excel, book, "worker", "A1:J10")["ms"] for _ in range(5)]
        phases.append({"phase": name, "requests": n, "median_ms": round(statistics.median(latencies), 1) if latencies else None,
                       "probe_100_cells_ms": round(statistics.median(small), 1),
                       "trivial_com_call_ms": round(trivial_call_ms(excel, book), 3), **process_stats(pid)})
        print(f"   {phases[-1]}")

    try:
        phase("fresh Excel, worker warm-up", "worker", "A1:J10", 10)
        phase("20 x worker classify 10,000", "worker", "A1:J1000", 20)
        phase("20 x worker read 100,000", "worker", "A1:J10000", 20)
        phase("10 x one-shot classify 10,000", "oneshot", "A1:J1000", 10)
        phase("20 x worker classify 10,000 again", "worker", "A1:J1000", 20)
        phase("after 30 s idle", None, "", 0) if not time.sleep(30) else None
    finally:
        for workbook in list(excel.Workbooks):
            workbook.Close(SaveChanges=False)
        excel.Quit()
        excel = book = None
        if not wait_until(lambda: not alive(pid), 15, 0.2):
            terminate(pid)
    report["perf"] = phases
    first, worst = phases[0]["probe_100_cells_ms"], max(p["probe_100_cells_ms"] for p in phases)
    record("Performance retest on a fresh Excel (worker drains its COM message queue): 100-cell probe latency across heavy workloads",
           worst <= 2 * first, f"fresh {first} ms, worst {worst} ms; see phases")


def main() -> None:
    WORK.mkdir(parents=True, exist_ok=True)
    RESULTS.mkdir(exist_ok=True)
    stop_worker()
    for name in (TARGET, "E4_Other.xlsx"):
        shutil.copy(E4_FIXTURES / name, WORK / name)
    (WORK / "y").mkdir()
    shutil.copy(E4_FIXTURES / "E4_Fixture_decoy.xlsx", WORK / "y" / TARGET)

    x = start_excel()
    held = {"y": start_excel()}
    pids = [pid_of(x), pid_of(held["y"])]
    try:
        other = x.Workbooks.Open(str(WORK / "E4_Other.xlsx"))
        target = x.Workbooks.Open(str(WORK / TARGET))
        held["decoy"] = held["y"].Workbooks.Open(str(WORK / "y" / TARGET))
        for excel in (x, held["y"]):
            excel.Workbooks.Open(str(PAYLOAD / "ExcelToolkit.xlam"))
            excel.Run("ExcelToolkit.xlam!Toolkit_SetQuiet", True)
        excel = None
        status = x.Run("ExcelToolkit.xlam!Toolkit_Status")
        record("Packaged add-in finds its engine next to itself", "engine_exists=True" in status, status)

        first = request(x, target, "auto", "A1:J10")
        record("First request starts the packaged worker and returns a result", first["status"] == "ok" and first["count"] == 100,
               f"{first['ms']:.0f} ms, route {first['route']}")
        if first["status"] != "ok":
            print("First request failed; worker log tail:", log_lines()[-3:])
            raise SystemExit("stopping early: the packaged engine is not answering")
        test_pipe_acl()
        test_validation(x, target)
        test_routing(x, held["y"], target, other, held["decoy"])
        test_oneshot(x, target)
        test_crash_fallback(x, target)
        test_worker_unavailable(x, target)
    finally:
        for pid, excel in zip(pids, (x, held.get("y"))):
            try:
                for workbook in list(excel.Workbooks):
                    workbook.Close(SaveChanges=False)
                excel.Quit()
            except Exception:  # noqa: BLE001
                pass
        x = None
        held.clear()
        for pid in pids:
            if not wait_until(lambda pid=pid: not alive(pid), 15, 0.2):
                terminate(pid)
    test_idle_cleanup()
    test_performance()
    stop_worker()

    (RESULTS / f"e6_runtime_{VERSION}.json").write_text(json.dumps(report, indent=1, default=str), encoding="utf-8")
    passed = sum(t["passed"] for t in report["tests"])
    print(f"\n{passed}/{len(report['tests'])} passed")


if __name__ == "__main__":
    main()
