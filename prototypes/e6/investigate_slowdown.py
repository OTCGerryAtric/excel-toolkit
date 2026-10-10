"""Locate the E4/E6 slowdown: worker process or Excel?

Fresh Excel + packaged add-in. Measure the 100-cell probe, run heavy 10k-finding requests,
measure again, restart only the worker, measure again. Also time a brand-new COM
connection from this script (attach + one read) at each step.
"""

import shutil
import statistics
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "engine"))
import run_e6_runtime as rt  # noqa: E402
from excel_link import Identity, attach, call  # noqa: E402


def fresh_connection_ms(identity: Identity) -> float:
    times = []
    for _ in range(5):
        started = time.perf_counter()
        app, book = attach(identity)
        call(app, lambda: book.Worksheets("Bulk").Range("A1:J10").Formula2)
        app = book = None
        times.append((time.perf_counter() - started) * 1000)
    return round(statistics.median(times), 1)


def main() -> None:
    rt.WORK.mkdir(parents=True, exist_ok=True)
    rt.stop_worker()
    shutil.copy(rt.E4_FIXTURES / rt.TARGET, rt.WORK / "slow.xlsx")
    excel = rt.start_excel()
    pid = rt.pid_of(excel)
    book = excel.Workbooks.Open(str(rt.WORK / "slow.xlsx"))
    excel.Workbooks.Open(str(rt.PAYLOAD / "ExcelToolkit.xlam"))
    excel.Run("ExcelToolkit.xlam!Toolkit_SetQuiet", True)
    identity = Identity.from_handoff(f"{int(excel.Hwnd)}|{pid}|{book.Name}|{book.FullName}")
    steps = []

    def measure(label: str) -> None:
        probes = [rt.request(excel, book, "worker", "A1:J10")["ms"] for _ in range(8)]
        steps.append((label, round(statistics.median(probes[3:]), 1), fresh_connection_ms(identity), rt.worker_pid()))
        print(f"{label:55} worker probe {steps[-1][1]:7} ms | new COM connection from a 2nd process {steps[-1][2]:6} ms | worker pid {steps[-1][3]}")

    # Start the worker ourselves with recycling disabled, so only the pumping change is measured.
    import os
    import subprocess
    launch_by_script = "script" in sys.argv  # default: let VBA start the worker, as in real use
    if launch_by_script:
        subprocess.Popen([str(rt.ENGINE), "worker", str(rt.STATE)], env=dict(os.environ, EXCEL_TOOLKIT_RECYCLE_CELLS=str(10**12)))
        rt.wait_until(lambda: rt.worker_pid() is not None, 10)
    print("worker started by", "script" if launch_by_script else "VBA Shell")
    try:
        measure("fresh")
        for _ in range(20):
            rt.request(excel, book, "worker", "A1:J1000")
        measure("after 20 x classify 10,000 (worker)")
        rt.stop_worker()
        if launch_by_script:
            subprocess.Popen([str(rt.ENGINE), "worker", str(rt.STATE)], env=dict(os.environ, EXCEL_TOOLKIT_RECYCLE_CELLS=str(10**12)))
            rt.wait_until(lambda: rt.worker_pid() is not None, 10)
        measure("after restarting ONLY the worker")
        for _ in range(20):
            rt.request(excel, book, "worker", "A1:J1000")
        measure("after another 20 x classify 10,000")
    finally:
        for workbook in list(excel.Workbooks):
            workbook.Close(SaveChanges=False)
        excel.Quit()
        excel = book = None
        if not rt.wait_until(lambda: not rt.alive(pid), 15, 0.2):
            rt.terminate(pid)
        rt.stop_worker()


if __name__ == "__main__":
    main()
