"""M3c manual checklist: create the disposable test workbook the owner opens in normal Excel.

Runs in a SEPARATE disposable Excel process (DispatchEx: no XLSTART, no PERSONAL.XLSB, no add-ins), builds
the workbook and saves it to the path given (a scratch folder), then closes. No add-in, no macros.

Sheet "Model" (sheets in order Model, A, AB, Other; numbers in A1:C3 on each): one formula per row in E,
the expected formula after Clear Sheet References in F (text), the live formula in G (FORMULATEXT), the
value before in H (pasted as a value) and "same value?" in I. Sheet "Constants" for Find Constants.

  <scratch venv>/python addin/live/m3c_manual_book.py <output.xlsx>
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import win32com.client
import win32process

from m2_live_check import Watchdog, _terminate

# (label, formula entered, expected formula afterwards)
ROWS = [
    ("Own sheet", "=Model!B2+1", "=B2+1"),
    ("Range, both halves named", "=SUM(Model!A1:Model!B2)", "=SUM(A1:B2)"),
    ("Entered quoted", "=SUM('Model'!A1:C1)", "=SUM(A1:C1)"),
    ("Other sheet kept", "=Other!B2+Model!B2", "=Other!B2+B2"),
    ("INDIRECT text kept", '=INDIRECT("Model!B2")+Model!C2', '=INDIRECT("Model!B2")+C2'),
    ("String literal kept", '=Model!B2&" see Model!B2"', '=B2&" see Model!B2"'),
    ("Sheets named A, AB: other sheets", "=A!B2+AB!B2", "=A!B2+AB!B2"),
    ("3D range (Model to AB)", "=SUM(Model:AB!B2)", "=SUM(Model:AB!B2)"),
    ("First half on another sheet (#VALUE!)", "=SUM(Other!A1:Model!B2)", "=SUM(Other!A1:Model!B2)"),
    ("TWO WAYS: 3D range + own ref", "=SUM(Model:AB!B2)+Model!C1", "=SUM(Model:AB!B2)+Model!C1"),
    ("TWO WAYS: stored like a 3D range", "=SUM(AB:'Model'!B2)+Model!C1", "=SUM(AB:Model!B2)+Model!C1"),
]
FIRST = 10


def build(path: Path) -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
    watchdog = Watchdog(pid)
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        names = ("Model", "A", "AB", "Other", "Constants")
        while wb.Worksheets.Count < len(names):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for index in range(1, len(names) + 1):
            wb.Worksheets(index).Name = f"tmp{index}"
        for index, name in enumerate(names, start=1):
            ws = wb.Worksheets(index)
            ws.Name = name
            if name != "Constants":
                ws.Range("A1:C3").Value2 = tuple(tuple(index * 1000 + r * 10 + c for c in range(3)) for r in range(3))
        watchdog.tick()

        model = wb.Worksheets("Model")
        model.Range("D9:I9").Value2 = (("Case", "Formula", "Expected after (text)", "Formula now",
                                        "Value before", "Same value?"),)
        model.Range("D9:I9").Font.Bold = True
        for row, (label, formula, expected) in enumerate(ROWS, start=FIRST):
            model.Range(f"D{row}").Value2 = label
            model.Range(f"E{row}").Formula2 = formula
            model.Range(f"F{row}").Value2 = "'" + expected
            model.Range(f"G{row}").Formula2 = f"=FORMULATEXT(E{row})"
            model.Range(f"I{row}").Formula2 = (f"=IFERROR(E{row}=H{row},"
                                               f"IFERROR(ERROR.TYPE(E{row})=ERROR.TYPE(H{row}),FALSE))")
        last = FIRST + len(ROWS) - 1
        model.Range(f"E{FIRST}:E{last}").Copy()
        model.Range(f"H{FIRST}").PasteSpecial(Paste=-4163)  # xlPasteValues
        excel.CutCopyMode = False
        model.Range(f"D{last + 2}").Value2 = "Text cell: Model!B2 is a note"
        model.Range(f"D{last + 3}").Value2 = "'Model'!A1 typed as text"
        model.Columns("D:I").AutoFit()
        watchdog.tick()

        constants = wb.Worksheets("Constants")
        constants.Range("A1:D1").Value2 = (("Typed numbers", "Formulas (numbers)", "Text", "Other"),)
        constants.Range("A1:D1").Font.Bold = True
        constants.Range("A2:A5").Value2 = ((1,), (2.5,), (-3,), (0.05,))
        constants.Range("A5").NumberFormat = "0%"
        constants.Range("B2:B5").Formula2 = "=A2*2"
        constants.Range("C2:C4").Value2 = (("abc",), ("'123",), ("Total",))
        constants.Range("D2").Value2 = True
        constants.Range("D3").Formula2 = '="text from a formula"'
        constants.Columns("A:D").AutoFit()

        model.Activate()
        model.Range("A1").Select()
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            path.unlink()
        wb.SaveAs(str(path), FileFormat=51)  # xlOpenXMLWorkbook (.xlsx, no macros)
        print({r: (model.Range(f"E{r}").Formula2, model.Range(f"H{r}").Value2) for r in range(FIRST, last + 1)})
        wb.Close(SaveChanges=False)
        excel.Quit()
    finally:
        watchdog.stop()
        del excel
        time.sleep(1)
        _terminate(pid)
    print("saved", path)


if __name__ == "__main__":
    build(Path(sys.argv[1]).resolve())
