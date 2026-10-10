"""E6 install lifecycle on THIS PC (current Windows user, not elevated). Requires Excel closed.

Simulated download (Mark of the Web) -> blocked-file evidence -> install 0.1.0 -> first run in
a normally started Excel (PERSONAL.XLSB and other add-ins load as usual) -> refusals (Excel
running, tampered package) -> upgrade 0.2.0 -> rollback -> uninstall -> verify the PC is back
to its baseline. Every Excel this script starts is closed again without saving.
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import winreg
import zipfile
from pathlib import Path

import pywintypes
import win32api
import win32con
import win32gui
import win32process

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "engine"))
from excel_link import ExcelBusy, Identity, IdentityError, attach, call  # noqa: E402

DIST = HERE / "dist"
WORK = HERE / "work" / time.strftime("install_%Y%m%d_%H%M%S")
RESULTS = HERE / "results"
ROOT = Path(os.environ["LOCALAPPDATA"]) / "ExcelToolkit"
OPTIONS = r"Software\Microsoft\Office\16.0\Excel\Options"
PERSONAL = Path(os.environ["APPDATA"]) / "Microsoft" / "Excel" / "XLSTART" / "PERSONAL.XLSB"
FIXTURE = HERE.parent / "e4" / "fixtures" / "E4_Fixture.xlsx"
report: dict = {"tests": []}


def record(test: str, passed, detail) -> None:
    report["tests"].append({"test": test, "passed": bool(passed), "detail": detail})
    print(f"[{'PASS' if passed else 'FAIL'}] {test}: {detail}")


def open_values() -> dict:
    values = {}
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, OPTIONS) as key:
            i = 0
            while True:
                try:
                    name, data, _ = winreg.EnumValue(key, i)
                except OSError:
                    break
                if name.upper().startswith("OPEN"):
                    values[name] = data
                i += 1
    except OSError:
        pass
    return values


def set_open(name: str, data: str | None) -> None:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, OPTIONS, 0, winreg.KEY_SET_VALUE) as key:
        if data is None:
            winreg.DeleteValue(key, name)
        else:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, data)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def has_motw(path: Path) -> bool:
    try:
        with open(f"{path}:Zone.Identifier", encoding="utf-8") as f:
            return "ZoneId=3" in f.read()
    except OSError:
        return False


def add_motw(path: Path) -> None:
    with open(f"{path}:Zone.Identifier", "w", encoding="utf-8") as f:
        f.write("[ZoneTransfer]\nZoneId=3\nReferrerUrl=https://example.invalid/\nHostUrl=https://example.invalid/ExcelToolkit.zip\n")


def excel_pids() -> set[int]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True).stdout
    return {int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"EXCEL.EXE"')}


def engine_running() -> list[int]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq ExcelToolkitEngine.exe", "/FO", "CSV", "/NH"], capture_output=True, text=True).stdout
    return [int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"ExcelToolkitEngine.exe"')]


def run_installer(package: Path, script: str, *args: str) -> tuple[int, str]:
    completed = subprocess.run(["cmd", "/c", str(package / script), "-NoPause", *args], capture_output=True, text=True, timeout=300)
    return completed.returncode, (completed.stdout + completed.stderr).strip()


def extract(zip_path: Path, dest: Path) -> Path:
    """Like Explorer's 'Extract all': extracted files inherit the zip's Mark of the Web."""
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(dest)
    package = next(dest.iterdir())
    if has_motw(zip_path):
        for file in package.rglob("*"):
            if file.is_file():
                add_motw(file)
    return package


# ---------- Normally started Excel ----------

class UserExcel:
    """Excel started the normal way (loads XLSTART and registered add-ins), on a test workbook."""

    def __init__(self, workbook: Path):
        excel_exe = winreg.QueryValue(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")
        before = excel_pids()
        self.process = subprocess.Popen([excel_exe, str(workbook)])
        self.pid = self.process.pid
        self.name = workbook.name
        deadline = time.monotonic() + 60
        self.identity = None
        while time.monotonic() < deadline and self.identity is None:
            time.sleep(0.5)
            for pid in excel_pids() - before:
                self.pid = pid
            hwnd = self._main_window()
            if hwnd:
                try:
                    self.identity = Identity.from_handoff(f"{hwnd}|{self.pid}|{self.name}|{workbook}")
                    self.app, self.book = attach(self.identity, timeout=2)
                except (IdentityError, ExcelBusy):
                    self.identity = None
        if self.identity is None:
            raise RuntimeError("Excel did not open the test workbook within 60 s")
        time.sleep(3)  # let startup add-ins finish loading

    def _main_window(self) -> int:
        found = []

        def visit(hwnd, _):
            if win32gui.GetClassName(hwnd) == "XLMAIN" and win32process.GetWindowThreadProcessId(hwnd)[1] == self.pid \
                    and self.name in win32gui.GetWindowText(hwnd):
                found.append(hwnd)
            return True

        win32gui.EnumWindows(visit, None)
        return found[0] if found else 0

    def run(self, macro: str, *args):
        return call(self.app, lambda: self.app.Run(macro, *args), timeout=10)

    def loaded(self) -> list[str]:
        return call(self.app, lambda: [b.Name for b in self.app.Workbooks] + [a.Name for a in self.app.AddIns if a.Installed])

    def request(self, mode: str, address: str) -> dict:
        request_id = self.run("ExcelToolkit.xlam!Toolkit_Request", mode, "classify", "Bulk", address, self.name)
        deadline = time.monotonic() + 90
        while time.monotonic() < deadline:
            status, count, ms, probe, error, route, workbook = (self.run("ExcelToolkit.xlam!Toolkit_GetResult", request_id).split("|") + [""] * 7)[:7]
            if status != "pending":
                return {"status": status, "count": count, "ms": ms, "route": route, "error": error}
            time.sleep(0.05)
        return {"status": "timeout"}

    def close(self) -> None:
        try:
            call(self.app, lambda: setattr(self.app, "DisplayAlerts", False))
            call(self.app, lambda: self.book.Close(False))
            call(self.app, lambda: self.app.Quit())
        except Exception:  # noqa: BLE001
            pass
        self.app = self.book = None
        deadline = time.monotonic() + 30
        while self.pid in excel_pids() and time.monotonic() < deadline:
            time.sleep(0.5)
        if self.pid in excel_pids():
            win32api.TerminateProcess(win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, self.pid), 1)


def first_run(label: str, expected_version: str) -> None:
    shutil.copy(FIXTURE, WORK / f"{label}.xlsx")
    excel = UserExcel(WORK / f"{label}.xlsx")
    try:
        status = excel.run("ExcelToolkit.xlam!Toolkit_Status")
        excel.run("ExcelToolkit.xlam!Toolkit_SetQuiet", True)
        loaded = excel.loaded()
        result = excel.request("auto", "A1:J10")
        record(f"{label}: add-in loads at Excel start-up (Auto_Open ran), version {expected_version}, audit returns",
               f"version={expected_version}" in status and "auto_open=True" in status and str(ROOT) in status
               and result["status"] == "ok" and result["count"] == "100",
               f"{status.split('|engine=')[0]}; result {result}")
        record(f"{label}: PERSONAL.XLSB and the other add-ins still load alongside", "PERSONAL.XLSB" in loaded
               and any("Arixcel" in name for name in loaded), loaded)
        if label == "first run":
            request_id = excel.run("ExcelToolkit.xlam!Toolkit_Request", "auto", "classify", "Bulk", "A1:J1000", excel.name)
            time.sleep(0.05)
            for pid in engine_running():
                win32api.TerminateProcess(win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid), 1)
            deadline = time.monotonic() + 60
            while time.monotonic() < deadline:
                text = excel.run("ExcelToolkit.xlam!Toolkit_GetResult", request_id)
                if not text.startswith("pending"):
                    break
                time.sleep(0.1)
            record("Installed build: worker killed mid-request, result still arrives via one-shot", text.startswith("ok|10000"), text)
    finally:
        excel.close()


# ---------- Main ----------

def main() -> None:
    if excel_pids():
        sys.exit("Excel is running. Close every Excel window, then run this again. Nothing was changed.")
    WORK.mkdir(parents=True, exist_ok=True)
    RESULTS.mkdir(exist_ok=True)
    for pid in engine_running():
        win32api.TerminateProcess(win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid), 1)
    if ROOT.exists() and not (ROOT / "current.json").exists():
        shutil.rmtree(ROOT)  # only test state from the runtime tests; no installation present
    baseline = {"open": open_values(), "personal": sha256(PERSONAL), "root_exists": ROOT.exists()}
    report["baseline"] = baseline
    print("Baseline OPEN values:", baseline["open"])

    # 1. Simulated download
    downloads = WORK / "Downloads"
    downloads.mkdir()
    packages = {}
    for version in ("0.1.0", "0.2.0"):
        zip_path = downloads / f"ExcelToolkit-{version}.zip"
        shutil.copy(DIST / f"ExcelToolkit-{version}.zip", zip_path)
        add_motw(zip_path)
        packages[version] = extract(zip_path, WORK / f"extracted_{version}")
    xlam_marked = has_motw(packages["0.1.0"] / "payload" / "ExcelToolkit.xlam")
    record("Downloaded package carries Mark of the Web into the extracted files", xlam_marked, f"xlam marked: {xlam_marked}")

    # 2. Evidence: what Excel does with a still-marked add-in (temporary OPEN entry, removed afterwards)
    blocked_copy = WORK / "blocked" / "ExcelToolkit.xlam"
    shutil.copytree(packages["0.1.0"] / "payload", blocked_copy.parent)
    for file in blocked_copy.parent.rglob("*"):
        if file.is_file():
            add_motw(file)
    free_name = next(n for n in (["OPEN"] + [f"OPEN{i}" for i in range(1, 20)]) if n not in baseline["open"])
    set_open(free_name, f'/R "{blocked_copy}"')
    try:
        shutil.copy(FIXTURE, WORK / "blocked_test.xlsx")
        excel = UserExcel(WORK / "blocked_test.xlsx")
        try:
            try:
                outcome = f"macros RAN: {excel.run('ExcelToolkit.xlam!Toolkit_Status')}"
            except Exception as error:  # noqa: BLE001
                outcome = f"macros blocked: {getattr(error, 'excepinfo', None) and error.excepinfo[2] or error}"
            loaded = excel.loaded()
        finally:
            excel.close()
        record("Evidence: an add-in still carrying Mark of the Web does not run its macros", outcome.startswith("macros blocked"),
               f"{outcome[:160]}; loaded: {[n for n in loaded if 'Toolkit' in n]}")
    finally:
        set_open(free_name, None)

    # 3. Refusals leave nothing changed
    tampered = WORK / "tampered" / packages["0.1.0"].name
    shutil.copytree(packages["0.1.0"], tampered)
    xlam = tampered / "payload" / "ExcelToolkit.xlam"
    xlam.write_bytes(xlam.read_bytes() + b"tamper")
    code, output = run_installer(tampered, "install.cmd")
    record("Tampered package: installer refuses, nothing changed", code == 1 and "damaged or altered" in output
           and open_values() == baseline["open"] and not (ROOT / "versions").exists(), output.splitlines()[-3:])

    blocker = UserExcel(WORK / "blocked_test.xlsx")
    try:
        code, output = run_installer(packages["0.1.0"], "install.cmd")
    finally:
        blocker.close()
    record("Excel running: installer refuses with instructions, nothing changed", code == 1 and "close every Excel window" in output
           and open_values() == baseline["open"], output.splitlines()[-3:])

    # 4. Install 0.1.0 from the marked download
    code, output = run_installer(packages["0.1.0"], "install.cmd")
    installed = ROOT / "versions" / "0.1.0"
    still_marked = [str(p.relative_to(installed)) for p in installed.rglob("*") if p.is_file() and has_motw(p)]
    added = {k: v for k, v in open_values().items() if k not in baseline["open"] or baseline["open"][k] != v}
    record("Install 0.1.0 (standard user, no admin): succeeds", code == 0, output.splitlines()[-1:])
    record("Installed files verified and unblocked (no Mark of the Web left)", not still_marked, still_marked[:5] or "none marked")
    record("Only one Excel registry value added; other add-ins untouched",
           len(added) == 1 and all(open_values().get(k) == v for k, v in baseline["open"].items()), added)
    record("PERSONAL.XLSB unchanged by install", sha256(PERSONAL) == baseline["personal"], "hash compared")

    # 5. First run
    first_run("first run", "0.1.0")

    # 6. Upgrade, 7. Rollback
    code, output = run_installer(packages["0.2.0"], "install.cmd")
    versions = sorted(p.name for p in (ROOT / "versions").iterdir())
    record("Upgrade to 0.2.0: succeeds, same registry value updated, previous version kept for rollback",
           code == 0 and versions == ["0.1.0", "0.2.0"] and len(open_values()) == len(baseline["open"]) + 1, f"versions {versions}")
    first_run("after upgrade", "0.2.0")
    code, output = run_installer(packages["0.2.0"], "install.cmd", "-Rollback")
    record("Rollback: succeeds", code == 0, output.splitlines()[-1:])
    first_run("after rollback", "0.1.0")

    # 8. Uninstall
    code, output = run_installer(packages["0.1.0"], "uninstall.cmd")
    record("Uninstall: Excel registry values exactly back to baseline, files and engine gone",
           code == 0 and open_values() == baseline["open"] and not ROOT.exists() and not engine_running(),
           {"open_now": open_values(), "root_exists": ROOT.exists(), "engines": engine_running()})
    record("PERSONAL.XLSB unchanged after the whole cycle", sha256(PERSONAL) == baseline["personal"], "hash compared")

    (RESULTS / "e6_install.json").write_text(json.dumps(report, indent=1, default=str), encoding="utf-8")
    passed = sum(t["passed"] for t in report["tests"])
    print(f"\n{passed}/{len(report['tests'])} passed")


if __name__ == "__main__":
    main()
