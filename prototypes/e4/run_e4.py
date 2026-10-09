"""E4 orchestrator: one-shot vs persistent worker, results returned to VBA.

Two new Excel processes (X: target + second workbook; Y: same-named decoy), each with the
E4Bridge add-in. The orchestrator plays the user: it activates workbooks and asks the add-in
for requests; timings are measured inside Excel (request -> callback). Audit requests are
read-only (verified by snapshot). Write timings run only on a new scratch workbook.
"""

import gc
import json
import os
import shutil
import statistics
import sys
import threading
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32api
import win32com.client
import win32con
import win32file
import win32process

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "e3"))
from excel_link import process_start_time  # noqa: E402

FIXTURES = HERE / "fixtures"
WORK = HERE / "work" / time.strftime("run_%Y%m%d_%H%M%S")  # fresh folder per run; earlier runs can't block it
STATE = WORK / "state"
RESULTS = HERE / "results"
PYTHON = sys.executable
XL_OPEN_XML_ADDIN = 55
TARGET = "E4_Fixture.xlsx"
REPEATS = 10

report: dict = {"tests": [], "timings": {}, "sizes": {}, "writes": {}, "idle": {}}


def record(test: str, passed, detail: str) -> None:
    passed = bool(passed)
    report["tests"].append({"test": test, "passed": passed, "detail": detail})
    print(f"[{'PASS' if passed else 'FAIL'}] {test}: {detail}")


# ---------- Excel helpers ----------

def start_excel():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


def pid_of(excel) -> int:
    return win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1]


def build_xlam(target: Path) -> None:
    crlf = WORK / "E4Bridge.bas"
    crlf.write_bytes((HERE / "E4Bridge.bas").read_text(encoding="ascii").replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Add()
        workbook.VBProject.VBComponents.Import(str(crlf))
        workbook.SaveAs(str(target), FileFormat=XL_OPEN_XML_ADDIN)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()


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


# ---------- Requests through the add-in ----------

def parse_result(text: str) -> dict:
    status, count, rows, ms, first, last, probe, error, worker_started, workbook = (text.split("|") + [""] * 10)[:10]
    return {"status": status, "count": int(count or 0), "rows": int(rows or 0), "ms": float(ms or "nan"),
            "first": first, "last": last, "probe": probe, "error": error, "worker_started": worker_started == "True",
            "workbook": workbook}


def request(excel, workbook, mode: str, op: str, sheet: str, address: str, timeout: float = 120) -> dict:
    workbook.Activate()
    request_id = excel.Run("E4Bridge.xlam!E4_Request", mode, op, sheet, address, workbook.Name)
    if request_id.startswith("ERROR"):
        return {"status": "vba_error", "error": request_id, "count": 0, "rows": 0, "ms": float("nan"),
                "probe": "", "first": "", "last": "", "worker_started": False, "id": None,
                "where": f"{workbook.Name} in pid {pid_of(excel)}"}
    result: dict = {}

    def done() -> bool:
        nonlocal result
        result = parse_result(excel.Run("E4Bridge.xlam!E4_GetResult", request_id))
        return result["status"] != "pending"

    wait_until(done, timeout, 0.01)
    result["id"] = request_id
    return result


def worker_pid() -> int | None:
    path = STATE / "worker.pid"
    try:
        pid = int(path.read_text(encoding="ascii"))
    except (OSError, ValueError):
        return None
    return pid if process_start_time(pid) is not None else None


def send_to_pipe(payload: dict) -> None:
    name = rf"\\.\pipe\excel_toolkit_e4_{os.environ.get('USERNAME', 'user')}"
    handle = win32file.CreateFile(name, win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, 0, None)
    win32file.WriteFile(handle, (json.dumps(payload) + "\n").encode("mbcs"))
    win32file.CloseHandle(handle)


def stop_worker() -> None:
    pid = worker_pid()
    if pid is None:
        return
    try:
        send_to_pipe({"op": "shutdown"})
    except pywintypes.error:
        pass
    if not wait_until(lambda: process_start_time(pid) is None, 10):
        terminate(pid)


def summarise(values: list[float]) -> dict:
    ordered = sorted(values)
    return {"n": len(values), "median": statistics.median(values),
            "p95": ordered[max(0, round(0.95 * len(ordered)) - 1)], "min": ordered[0], "max": ordered[-1]}


def snapshot(workbook) -> dict:
    return {s.Name: (s.UsedRange.Address, s.UsedRange.Formula2, s.UsedRange.Value2) for s in workbook.Worksheets}


# ---------- Tests ----------

def test_latency(x, target) -> None:
    for mode in ("oneshot", "worker"):
        runs = [request(x, target, mode, "classify", "Bulk", "A1:J10") for _ in range(REPEATS)]
        ok = [r for r in runs if r["status"] == "ok" and r["count"] == 100]
        warm = [r["ms"] for r in ok[1:]] if mode == "worker" else [r["ms"] for r in ok]
        report["timings"][mode] = {"first_ms": runs[0]["ms"], "first_started_worker": runs[0]["worker_started"],
                                   "warm": summarise(warm) if warm else None}
        record(f"{mode}: {REPEATS} classify requests (100 findings) answered correctly", len(ok) == REPEATS,
               f"first {runs[0]['ms']:.0f} ms (worker started: {runs[0]['worker_started']}); "
               f"median {statistics.median(warm):.0f} ms, p95 {summarise(warm)['p95']:.0f} ms; first finding {runs[0]['first']}")


def test_sizes(x, target) -> None:
    cases = [("classify", "A1:J10", 100), ("classify", "A1:J100", 1000), ("classify", "A1:J1000", 10000),
             ("read", "A1", 1), ("read", "A1:J100", 1000), ("read", "A1:J1000", 10000), ("read", "A1:J10000", 100000)]
    for mode in ("oneshot", "worker"):
        for op, address, expected in cases:
            runs = [request(x, target, mode, op, "Bulk", address) for _ in range(3)]
            good = [r for r in runs if r["status"] == "ok" and r["count"] == expected and (op == "read" or r["rows"] == expected)]
            report["sizes"][f"{mode} {op} {expected}"] = summarise([r["ms"] for r in good]) if good else None
            record(f"{mode}: {op} {expected:,} cells returned to VBA", len(good) == len(runs),
                   f"median {statistics.median([r['ms'] for r in good]):.0f} ms" if good else f"{runs[0]}")


def test_kinds(x, target) -> None:
    result = request(x, target, "worker", "classify", "Mixed", "A1:A6")
    record("Finding kinds classified (formula, constant in formula, number, text, formula error, boolean)",
           result["status"] == "ok" and result["count"] == 6 and result["first"] == "Mixed!A1=formula"
           and result["last"] == "Mixed!A6=boolean", f"{result['first']} ... {result['last']}, count {result['count']}")


def test_no_mixing(x, y, target, other, decoy) -> None:
    plan = [(x, target, "10"), (y, decoy, "77"), (x, other, "5555")] * 20
    wrong = []
    for mode in ("worker", "oneshot"):
        for excel, workbook, expected in plan if mode == "worker" else plan[:9]:
            result = request(excel, workbook, mode, "classify", "Bulk", "A1:B2")
            if result["status"] != "ok" or result["probe"] != expected or result["workbook"] != workbook.Name:
                wrong.append((mode, workbook.Name, expected, result))
    record("Interleaved requests across 2 processes and 3 workbooks (two share a name): every result came from the right workbook",
           not wrong, f"{len(plan) + 9} requests; wrong: {wrong[:3]}")
    rejected_before = x.Run("E4Bridge.xlam!E4_GetRejectedCount")
    verdict_unknown = x.Run("E4Bridge.xlam!E4_OnResult", "no-such-request", pid_of(x), TARGET, target.FullName, "ok", "", 1, 0, "")
    target.Activate()
    pending_id = x.Run("E4Bridge.xlam!E4_Request", "none", "classify", "Bulk", "A1")  # mode 'none': nothing sent
    verdict_identity = x.Run("E4Bridge.xlam!E4_OnResult", pending_id, pid_of(y), TARGET, decoy.FullName, "ok", "", 77, 0, "")
    rejected_after = x.Run("E4Bridge.xlam!E4_GetRejectedCount")
    record("VBA rejects a result for an unknown request and one with another process's identity",
           verdict_unknown.startswith("rejected") and verdict_identity.startswith("rejected") and rejected_after - rejected_before == 2,
           f"{verdict_unknown!r}; {verdict_identity!r}")


def test_worker_crash(x, target) -> None:
    pid = worker_pid()
    target.Activate()
    request_id = x.Run("E4Bridge.xlam!E4_Request", "worker", "classify", "Bulk", "A1:J10000")
    time.sleep(0.3)
    terminate(pid)
    result: dict = {}

    def settled() -> bool:
        nonlocal result
        result = parse_result(x.Run("E4Bridge.xlam!E4_GetResult", request_id))
        return result["status"] != "pending"

    started = time.monotonic()
    wait_until(settled, 15, 0.05)
    record("Worker killed mid-request: VBA watchdog reports a clear failure", result["status"] == "failed"
           and "worker stopped" in result["error"], f"{time.monotonic() - started:.1f} s: {result['status']} - {result['error']}")
    after = request(x, target, "worker", "classify", "Bulk", "A1:J10")
    record("Next request restarts the worker and succeeds", after["status"] == "ok" and after["worker_started"],
           f"{after['ms']:.0f} ms, worker started: {after['worker_started']}, new pid {worker_pid()} (old {pid})")


def test_excel_lifecycle(x, held: dict, target) -> None:
    y, decoy = held.pop("y"), held.pop("decoy")
    y_pid = pid_of(y)
    request(y, decoy, "worker", "classify", "Bulk", "A1:J10")  # worker has now talked to Y
    decoy.Close(SaveChanges=False)
    y.Workbooks("E4Bridge.xlam").Close(SaveChanges=False)
    y.Quit()
    y = decoy = None  # the orchestrator's own references must go too, or Excel can't exit
    gc.collect()
    pythoncom.CoFreeUnusedLibraries()
    exited = wait_until(lambda: process_start_time(y_pid) is None, 20, 0.2)
    record("Excel closes normally after the worker served it (worker holds no stale COM references)", exited,
           f"Y process {y_pid} exited: {exited}")
    if not exited:
        terminate(y_pid)

    before = sum(1 for _ in (STATE / "worker_log.jsonl").open(encoding="utf-8"))
    send_to_pipe({"id": "stale-1", "op": "classify", "sheet": "Bulk", "address": "A1", "hwnd": 1,
                  "pid": y_pid, "workbook": TARGET, "full_name": str(decoy_path())})
    wait_until(lambda: sum(1 for _ in (STATE / "worker_log.jsonl").open(encoding="utf-8")) > before, 15, 0.1)
    last = json.loads((STATE / "worker_log.jsonl").read_text(encoding="utf-8").splitlines()[-1])
    record("Request for a closed Excel: worker answers with a clear error and keeps running",
           last.get("status") == "failed" and "IdentityError" in last.get("error", "") and worker_pid() is not None,
           f"{last.get('error')}; verdict: {last.get('verdict')}")

    after = request(x, target, "worker", "classify", "Bulk", "A1:J10")
    record("Worker still serves the other Excel afterwards", after["status"] == "ok" and after["probe"] == "10",
           f"{after['ms']:.0f} ms")


def decoy_path() -> Path:
    return WORK / "y" / TARGET


def test_killed_mid_request(x, target) -> None:
    z = start_excel()
    shutil.copy(FIXTURES / "E4_Fixture.xlsx", WORK / "z" / TARGET)
    book = z.Workbooks.Open(str(WORK / "z" / TARGET))
    z.Workbooks.Open(str(WORK / "E4Bridge.xlam"))
    z.Run("E4Bridge.xlam!E4_Configure", PYTHON, str(HERE), str(STATE))
    z_pid = pid_of(z)
    book.Activate()
    z.Run("E4Bridge.xlam!E4_Request", "worker", "classify", "Bulk", "A1:J10000")
    del book, z
    time.sleep(0.4)
    terminate(z_pid)
    time.sleep(3)
    after = request(x, target, "worker", "classify", "Bulk", "A1:J10")
    tail = [json.loads(line) for line in (STATE / "worker_log.jsonl").read_text(encoding="utf-8").splitlines()[-3:]]
    lost = [t for t in tail if t.get("pid") == z_pid]
    record("Excel killed while the worker was serving it: worker survives and serves the next request",
           after["status"] == "ok" and worker_pid() is not None,
           f"worker log for the killed Excel: {[(t.get('status'), t.get('error'), t.get('verdict')) for t in lost]}")


def test_idle(pid: int) -> None:
    handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)

    def cpu() -> float:
        times = win32process.GetProcessTimes(handle)
        return (times["KernelTime"] + times["UserTime"]) / 1e7

    memory = win32process.GetProcessMemoryInfo(handle)
    cpu_before = cpu()
    time.sleep(10)
    cpu_after = cpu()
    report["idle"] = {"working_set_mb": memory["WorkingSetSize"] / 1e6, "private_mb": memory["PagefileUsage"] / 1e6,
                      "cpu_seconds_over_10s_idle": cpu_after - cpu_before}
    record("Idle worker resource use measured", True,
           f"working set {memory['WorkingSetSize'] / 1e6:.0f} MB, private {memory['PagefileUsage'] / 1e6:.0f} MB, "
           f"CPU over 10 s idle {cpu_after - cpu_before:.3f} s")


def test_writes(x) -> None:
    scratch = x.Workbooks.Add()
    sheet = scratch.Worksheets(1)
    for label, address in (("1", "A1"), ("1k", "A1:J100"), ("10k", "A1:J1000"), ("100k", "A1:J10000")):
        rows = int(address.split(":")[-1][1:]) if ":" in address else 1
        cols = 10 if ":" in address else 1
        timings = {}
        started = time.perf_counter()
        sheet.Range(address).NumberFormat = "#,##0.00"
        timings["python_number_format_ms"] = (time.perf_counter() - started) * 1000
        timings["vba_number_format_ms"] = x.Run("E4Bridge.xlam!E4_BenchFormat", scratch.Name, sheet.Name, address, "0.0")
        payload = 1.5 if rows == 1 else tuple(tuple(float(r * cols + c) for c in range(cols)) for r in range(rows))
        started = time.perf_counter()
        sheet.Range(address).Value2 = payload
        timings["python_write_values_ms"] = (time.perf_counter() - started) * 1000
        report["writes"][label] = timings
    record("Write timings on a scratch workbook (formats and values, 1 to 100k cells)", True, json.dumps(report["writes"]))
    scratch.Close(SaveChanges=False)


# ---------- Main ----------

def main() -> None:
    stop_worker()
    for folder in (WORK / "x", WORK / "y", WORK / "z", STATE, RESULTS):
        folder.mkdir(parents=True, exist_ok=True)
    shutil.copy(FIXTURES / "E4_Fixture.xlsx", WORK / "x" / TARGET)
    shutil.copy(FIXTURES / "E4_Other.xlsx", WORK / "x" / "E4_Other.xlsx")
    shutil.copy(FIXTURES / "E4_Fixture_decoy.xlsx", decoy_path())
    build_xlam(WORK / "E4Bridge.xlam")

    x = start_excel()
    held = {"y": start_excel()}
    pids = [pid_of(x), pid_of(held["y"])]
    try:
        other = x.Workbooks.Open(str(WORK / "x" / "E4_Other.xlsx"))
        target = x.Workbooks.Open(str(WORK / "x" / TARGET))
        held["decoy"] = held["y"].Workbooks.Open(str(decoy_path()))
        for excel in (x, held["y"]):
            excel.Workbooks.Open(str(WORK / "E4Bridge.xlam"))
            excel.Run("E4Bridge.xlam!E4_Configure", PYTHON, str(HERE), str(STATE))
        excel = None  # don't keep an extra reference to Y alive in this scope
        before = snapshot(target)

        test_latency(x, target)
        test_kinds(x, target)
        test_sizes(x, target)
        test_no_mixing(x, held["y"], target, other, held["decoy"])
        test_worker_crash(x, target)
        test_excel_lifecycle(x, held, target)
        test_killed_mid_request(x, target)
        record("Audit requests left the target workbook unchanged", snapshot(target) == before, "full snapshot compared")
        pid = worker_pid()
        if pid:
            test_idle(pid)
        test_writes(x)
    finally:
        stop_worker()
        for pid in pids:
            if process_start_time(pid) is not None:
                try:
                    excel = x if pid == pids[0] else held.get("y")
                    if excel is None:
                        continue
                    for workbook in list(excel.Workbooks):
                        workbook.Close(SaveChanges=False)
                    excel.Quit()
                except Exception:  # noqa: BLE001
                    pass
        x = None
        held.clear()
        for pid in pids:
            if not wait_until(lambda pid=pid: process_start_time(pid) is None, 15, 0.2):
                terminate(pid)

    (RESULTS / "e4_report.json").write_text(json.dumps(report, indent=1, default=str), encoding="utf-8")
    shutil.copy(STATE / "worker_log.jsonl", RESULTS / "worker_log.jsonl")
    passed = sum(t["passed"] for t in report["tests"])
    print(f"\n{passed}/{len(report['tests'])} passed")


if __name__ == "__main__":
    main()
