"""E2 orchestrator: two Excel processes, VBA -> Python handoff, pywin32 vs xlwings.

Setup (the only writes): copies fixtures into prototypes/e2/work, opens them in two new
Excel processes, makes unsaved edits and a multi-area selection in the target workbook.
Everything else is read-only and verified by before/after snapshots. The user's own
Excel is only touched by the 'naive' comparison, which reads the instance's window handle
and workbook count.
"""

import gc
import json
import shutil
import statistics
import subprocess
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32api
import win32com.client
import win32con
import win32process

HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"
WORK = HERE / "work"
RESULTS = HERE / "results"
RAW = RESULTS / "raw"
READER = HERE / "e2_reader.py"
PYTHON = sys.executable
RUNS = 5
TARGET = "E2_FixtureA.xlsx"
XL_OPEN_XML_ADDIN = 55
STILL_ACTIVE = 259


# ---------- Excel helpers ----------

def start_excel():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


def pid_of(excel) -> int:
    return win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1]


def process_alive(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    except pywintypes.error:
        return False
    return win32process.GetExitCodeProcess(handle) == STILL_ACTIVE


def build_xlam(target: Path) -> None:
    crlf = WORK / "E2Handoff.bas"
    source = (HERE / "E2Handoff.bas").read_text(encoding="ascii")
    crlf.write_bytes(source.replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Add()
        workbook.VBProject.VBComponents.Import(str(crlf))
        workbook.SaveAs(str(target), FileFormat=XL_OPEN_XML_ADDIN)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()


def formulas_or_marker(used):
    try:
        return used.Formula2
    except pywintypes.com_error:
        return "<formula not readable>"


def snapshot(excel, workbook) -> dict:
    """Everything a read could plausibly disturb."""
    state = {
        "saved": bool(workbook.Saved),
        "calculation": int(excel.Calculation),
        "active_sheet": workbook.ActiveSheet.Name,
        "selection": workbook.Windows(1).RangeSelection.Address,
        "names": sorted(name.Name for name in workbook.Names),
        "sheet_order": [sheet.Name for sheet in workbook.Worksheets],
        "sheets": {},
    }
    for sheet in workbook.Worksheets:
        used = sheet.UsedRange
        entry = {
            "visible": int(sheet.Visible),
            "protected": bool(sheet.ProtectContents),
            "used": used.Address,
            "formula2": formulas_or_marker(used),
            "value2": used.Value2,
        }
        if sheet.Name != "Bulk":
            entry["formats"] = [
                (cell.Address, cell.NumberFormat, int(cell.Interior.Color), bool(cell.Font.Bold), bool(cell.FormulaHidden))
                for cell in used.Cells
            ]
        state["sheets"][sheet.Name] = entry
    return state


# ---------- Reader invocation ----------

def wait_for(path: Path, timeout: float) -> bool:
    deadline = time.perf_counter() + timeout
    while time.perf_counter() < deadline:
        if path.exists():
            return True
        time.sleep(0.02)
    return False


def handoff_via_vba(excel, backend: str, label: str) -> dict:
    out = RAW / f"{label}.json"
    out.unlink(missing_ok=True)
    started = time.perf_counter()
    command = excel.Run("E2Handoff.xlam!E2_Handoff", PYTHON, str(READER), backend, str(out))
    if not wait_for(out, 90):
        return {"ok": False, "error": "timed out waiting for reader", "command": command}
    result = json.loads(out.read_text(encoding="utf-8"))
    result["wall_ms"] = (time.perf_counter() - started) * 1000
    result["command"] = command
    return result


def run_reader(backend: str, hwnd: int, pid: int, workbook: str, label: str) -> dict:
    out = RAW / f"{label}.json"
    out.unlink(missing_ok=True)
    started = time.perf_counter()
    completed = subprocess.run(
        [PYTHON, str(READER), "--backend", backend, "--hwnd", str(hwnd), "--pid", str(pid),
         "--workbook", workbook, "--out", str(out)],
        capture_output=True, text=True, timeout=120,
    )
    wall = (time.perf_counter() - started) * 1000
    result = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {"ok": False, "error": completed.stderr[-500:]}
    result.update(wall_ms=wall, exit_code=completed.returncode)
    return result


# ---------- Correctness checks ----------

def cell(data: dict, sheet: str, row: int, col: int, kind: str = "value"):
    return data["sheets"][sheet][kind][row][col]


def check_target(data: dict) -> list[tuple[str, bool, str]]:
    checks: list[tuple[str, bool, str]] = []

    def add(name: str, actual, expected) -> None:
        checks.append((name, actual == expected, f"got {actual!r}, expected {expected!r}"))

    add("Correct instance and file (work/x)", Path(data["full_name"]).parent.name, "x")
    add("Correct workbook among several open (Inputs!A1 = 10, not decoy 77 or FixtureB 5555)", cell(data, "Inputs", 0, 0), 10)
    add("Unsaved edit: Calc!B2 value", cell(data, "Calc", 1, 1), 999)
    add("Unsaved edit: Calc!B3 formula", cell(data, "Calc", 2, 1, "formula2"), "=B2*3")
    add("Unsaved edit: Calc!B3 recalculated value", cell(data, "Calc", 2, 1), 2997)
    add("Workbook reported as unsaved", data["saved"], False)
    add("Multi-area selection: sheet", data["selection"]["sheet"], "Calc")
    add("Multi-area selection: areas", data["selection"]["areas"], ["$B$2:$B$5", "$D$7:$E$9"])
    add("Cross-sheet formula: Calc!A1", cell(data, "Calc", 0, 0, "formula2"), "=Inputs!A1*2")
    add("Cross-sheet value: Calc!A1", cell(data, "Calc", 0, 0), 20)
    add("Reference into hidden sheet: Calc!C1", cell(data, "Calc", 0, 2), 43)
    expected_errors = ["#DIV/0!", "#N/A", "#REF!", "#VALUE!", "#NAME?", "#NUM!", "#SPILL!"]
    actual_errors = [cell(data, "Errors", row, 0) for row in range(7)]
    add("Error types preserved (Errors!A1:A7)", actual_errors, [{"error": e} for e in expected_errors])
    add("Text that looks like an error stays text (Errors!B1)", cell(data, "Errors", 0, 1), "#N/A")
    add("Dynamic array formula: Dyn!A1", cell(data, "Dyn", 0, 0, "formula2"), "=SEQUENCE(5)")
    add("Spilled cell has no formula: Dyn!A2", cell(data, "Dyn", 1, 0, "formula2"), "")
    add("Spilled values: Dyn!A1:A5", [cell(data, "Dyn", r, 0) for r in range(5)], [1, 2, 3, 4, 5])
    add("Spill reference: Dyn!E1 formula", cell(data, "Dyn", 0, 4, "formula2"), "=SUM(A1#)")
    add("Spill reference: Dyn!E1 value", cell(data, "Dyn", 0, 4), 15)
    add("Spill range detected", data["spill"], {"Dyn!A1_has_spill": True, "Dyn!A1_spill": "$A$1:$A$5"})
    add("Hidden sheet visibility (0)", data["sheets"]["Hidden"]["visible"], 0)
    add("Very hidden sheet visibility (2)", data["sheets"]["VeryHidden"]["visible"], 2)
    add("Very hidden sheet formula readable", cell(data, "VeryHidden", 0, 0, "formula2"), "=Inputs!A1")
    add("Protected sheet detected", data["sheets"]["Protected"]["protected"], True)
    add("Protected sheet value readable", cell(data, "Protected", 0, 0), 110)
    checks.append((
        "Protected sheet, hidden formula: what COM returns (informational)", True,
        f"got {cell(data, 'Protected', 0, 0, 'formula2')!r}",
    ))
    return checks


def check_bulk(timings: dict) -> list[tuple[str, bool, str]]:
    out = []
    for label, cells, last_value in (("1k", 1000, 1010), ("10k", 10000, 10010)):
        bulk = timings["bulk"][label]
        ok = (bulk["cells"] == cells and bulk["first"] == ["=Inputs!$A$1*ROW()+COLUMN()", 11]
              and bulk["last"][1] == last_value)
        out.append((f"Bulk {label}: count, first and last cell", ok, f"{bulk['cells']} cells, first {bulk['first']}, last {bulk['last']}"))
    return out


# ---------- Report ----------

def median(values: list[float]) -> float:
    return statistics.median(values) if values else float("nan")


def main() -> None:
    shutil.rmtree(WORK, ignore_errors=True)
    (WORK / "x").mkdir(parents=True)
    (WORK / "y").mkdir(parents=True)
    RAW.mkdir(parents=True, exist_ok=True)
    shutil.copy(FIXTURES / "E2_FixtureA.xlsx", WORK / "x" / TARGET)
    shutil.copy(FIXTURES / "E2_FixtureB.xlsx", WORK / "x" / "E2_FixtureB.xlsx")
    shutil.copy(FIXTURES / "E2_FixtureA_decoy.xlsx", WORK / "y" / TARGET)  # same name, other process
    build_xlam(WORK / "E2Handoff.xlam")

    report: dict = {"runs": {}, "naive": None, "decoy_read": {}, "failures": {}, "read_only": {}}
    x = start_excel()
    y = start_excel()
    try:
        x.Workbooks.Open(str(WORK / "x" / "E2_FixtureB.xlsx"))
        target = x.Workbooks.Open(str(WORK / "x" / TARGET))
        x.Workbooks.Open(str(WORK / "E2Handoff.xlam"))
        decoy = y.Workbooks.Open(str(WORK / "y" / TARGET))

        calc = target.Worksheets("Calc")
        calc.Range("B2").Value = 999
        calc.Range("B3").Formula = "=B2*3"
        target.Activate()
        calc.Activate()
        calc.Range("B2:B5,D7:E9").Select()

        x_hwnd, x_pid = int(x.Hwnd), pid_of(x)
        y_hwnd, y_pid = int(y.Hwnd), pid_of(y)
        report["instances"] = {"x": {"hwnd": x_hwnd, "pid": x_pid}, "y": {"hwnd": y_hwnd, "pid": y_pid}}

        before = {"target": snapshot(x, target), "decoy": snapshot(y, decoy)}

        for backend in ("pywin32", "xlwings"):
            report["runs"][backend] = [handoff_via_vba(x, backend, f"{backend}_vba_{i}") for i in range(RUNS)]

        report["naive"] = run_reader("naive", x_hwnd, x_pid, TARGET, "naive")
        for backend in ("pywin32", "xlwings"):
            report["decoy_read"][backend] = run_reader(backend, y_hwnd, y_pid, TARGET, f"{backend}_decoy")

        after = {"target": snapshot(x, target), "decoy": snapshot(y, decoy)}
        for key in before:
            diffs = [k for k in before[key] if before[key][k] != after[key][k]]
            report["read_only"][key] = diffs

        # Failure handling
        for backend in ("pywin32", "xlwings"):
            report["failures"][f"{backend}: workbook not open"] = run_reader(backend, x_hwnd, x_pid, "Missing.xlsx", f"{backend}_fail_missing")
            report["failures"][f"{backend}: window/pid mismatch"] = run_reader(backend, x_hwnd, y_pid, TARGET, f"{backend}_fail_mismatch")
        decoy.Close(SaveChanges=False)
        decoy = None
        for backend in ("pywin32", "xlwings"):
            report["failures"][f"{backend}: target workbook closed"] = run_reader(backend, y_hwnd, y_pid, TARGET, f"{backend}_fail_closed_wb")
        y.Quit()
        y = None
        gc.collect()
        pythoncom.CoFreeUnusedLibraries()
        deadline = time.perf_counter() + 30
        while process_alive(y_pid) and time.perf_counter() < deadline:
            time.sleep(0.2)
        report["y_process_exited"] = not process_alive(y_pid)
        for backend in ("pywin32", "xlwings"):
            report["failures"][f"{backend}: Excel instance closed"] = run_reader(backend, y_hwnd, y_pid, TARGET, f"{backend}_fail_closed_app")
    finally:
        for excel in (x, y):
            if excel is None:
                continue
            try:
                for workbook in list(excel.Workbooks):
                    workbook.Close(SaveChanges=False)
                excel.Quit()
            except pywintypes.com_error:
                pass

    (RESULTS / "e2_report.json").write_text(json.dumps(report, indent=1, default=str), encoding="utf-8")
    write_summary(report)


def write_summary(report: dict) -> None:
    lines = ["# E2 results", ""]
    lines += [f"Instances: target X pid {report['instances']['x']['pid']}, decoy Y pid {report['instances']['y']['pid']}", ""]

    lines += ["## Correctness (first VBA handoff per backend; later runs compared to it)", ""]
    for backend, runs in report["runs"].items():
        first = runs[0]
        if not first.get("ok"):
            lines.append(f"- **{backend}: FAILED** — {first.get('error')}")
            continue
        checks = check_target(first["data"]) + check_bulk(first["timings"])
        passed = sum(ok for _, ok, _ in checks)
        consistent = all(run.get("ok") and run["data"] == first["data"] for run in runs)
        lines.append(f"### {backend}: {passed}/{len(checks)} checks passed; all {len(runs)} runs identical: {consistent}")
        lines += [f"- {'PASS' if ok else 'FAIL'} — {name}: {detail}" for name, ok, detail in checks]
        lines.append(f"- API fallbacks needed: {first['data']['api_fallbacks'] or 'none'}")
        lines.append("")

    lines += ["## Explicit handoff to the other instance (decoy, same file name)", ""]
    for backend, result in report["decoy_read"].items():
        if result.get("ok"):
            data = result["data"]
            lines.append(f"- {backend}: read {Path(data['full_name']).parent.name}/{data['name']}, Inputs!A1 = {cell(data, 'Inputs', 0, 0)} (expected y, 77)")
        else:
            lines.append(f"- {backend}: FAILED — {result.get('error')}")
    naive = report["naive"]
    lines += ["", "## Naive GetActiveObject (comparison only)", ""]
    if naive.get("ok"):
        n = naive["naive"]
        lines.append(f"- Attached to pid {n['attached_pid']} (target X is {report['instances']['x']['pid']}); "
                     f"target workbook present: {n['target_present']}; workbooks open there: {n['workbook_count']}"
                     + (f"; Inputs!A1 = {n.get('Inputs!A1')}" if n["target_present"] else ""))
    else:
        lines.append(f"- Failed: {naive.get('error')}")

    lines += ["", "## Timings (ms), warm runs 2–5 median; run 1 shown separately", "",
              "| Backend | Run 1 wall | Warm wall | Warm process | Warm attach | Warm read small | Warm bulk 1k | Warm bulk 10k |",
              "|---|---|---|---|---|---|---|---|"]
    for backend, runs in report["runs"].items():
        ok_runs = [r for r in runs if r.get("ok")]
        warm = ok_runs[1:]
        if not ok_runs:
            continue

        def m(key):
            return median([r["timings"][key] for r in warm])

        lines.append(
            f"| {backend} | {ok_runs[0]['wall_ms']:.0f} | {median([r['wall_ms'] for r in warm]):.0f} | {m('process_ms'):.0f} | "
            f"{m('attach_ms'):.1f} | {m('read_small_ms'):.1f} | "
            f"{median([r['timings']['bulk']['1k']['ms'] for r in warm]):.1f} | {median([r['timings']['bulk']['10k']['ms'] for r in warm]):.1f} |"
        )

    lines += ["", "## Failure handling", "", "| Case | ok | Exit code | Wall ms | Error |", "|---|---|---|---|---|"]
    for case, result in report["failures"].items():
        lines.append(f"| {case} | {result.get('ok')} | {result.get('exit_code')} | {result.get('wall_ms', 0):.0f} | {str(result.get('error', ''))[:120]} |")
    lines.append(f"\nDecoy Excel process exited before the 'instance closed' test: {report.get('y_process_exited')}")

    lines += ["", "## Read-only verification", ""]
    for key, diffs in report["read_only"].items():
        lines.append(f"- {key}: {'unchanged' if not diffs else 'CHANGED: ' + ', '.join(diffs)}")

    summary = "\n".join(lines) + "\n"
    (RESULTS / "e2_summary.md").write_text(summary, encoding="utf-8")
    print(summary)


if __name__ == "__main__":
    main()
