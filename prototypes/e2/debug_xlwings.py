"""Reproduce the xlwings reader hang against one disposable Excel instance."""

import shutil
import subprocess
import sys
from pathlib import Path

import win32com.client
import win32process

HERE = Path(__file__).resolve().parent
WORK = HERE / "work" / "debug"


def main() -> None:
    WORK.mkdir(parents=True, exist_ok=True)
    shutil.copy(HERE / "fixtures" / "E2_FixtureA.xlsx", WORK / "E2_FixtureA.xlsx")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        excel.Workbooks.Open(str(WORK / "E2_FixtureA.xlsx"))
        hwnd = int(excel.Hwnd)
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        out = WORK / "xlwings_debug.json"
        out.unlink(missing_ok=True)
        completed = subprocess.run(
            [sys.executable, str(HERE / "e2_reader.py"), "--backend", "xlwings", "--hwnd", str(hwnd),
             "--pid", str(pid), "--workbook", "E2_FixtureA.xlsx", "--out", str(out)],
            capture_output=True, text=True, timeout=90,
        )
        print("exit", completed.returncode)
        print("stderr:", completed.stderr[-4000:])
        if out.exists():
            print(out.read_text(encoding="utf-8")[:1500])
    finally:
        for workbook in list(excel.Workbooks):
            workbook.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    main()
