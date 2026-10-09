"""Create disposable E2 fixture workbooks in a separate hidden Excel instance.

E2_FixtureA.xlsx        - the target workbook
E2_FixtureA_decoy.xlsx  - same sheets and addresses, different values; opened under the
                          same file name in a second Excel process to catch wrong-instance reads
E2_FixtureB.xlsx        - a second workbook open in the target instance
"""

from pathlib import Path

import win32com.client

HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"
XL_OPEN_XML_WORKBOOK = 51
XL_SHEET_HIDDEN = 0
XL_SHEET_VERY_HIDDEN = 2


def build_fixture_a(excel, path: Path, input_base: int) -> None:
    workbook = excel.Workbooks.Add()
    while workbook.Worksheets.Count < 8:
        workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    names = ["Inputs", "Calc", "Errors", "Dyn", "Hidden", "VeryHidden", "Protected", "Bulk"]
    for index, name in enumerate(names, start=1):
        workbook.Worksheets(index).Name = name
    ws = {name: workbook.Worksheets(name) for name in names}

    for row in range(1, 11):
        ws["Inputs"].Cells(row, 1).Value = input_base + row - 1
    ws["Inputs"].Range("B1").Value = "label"

    calc = ws["Calc"]
    calc.Range("A1").Formula = "=Inputs!A1*2"
    calc.Range("A2").Formula = "=SUM(Inputs!A1:A10)"
    calc.Range("B2").Value = 5
    calc.Range("B3").Formula = "=B2*2"
    calc.Range("B4:B5").Value = 7
    calc.Range("C1").Formula = "=Hidden!A1+1"
    calc.Range("D7:E9").Value = 3

    errors = ws["Errors"]
    errors.Range("A1").Formula = "=1/0"
    errors.Range("A2").Formula = "=NA()"
    errors.Range("A3").Formula = "=#REF!+1"
    errors.Range("A4").Formula = '="a"+1'
    errors.Range("A5").Formula = "=NOSUCHFUNCTION()"
    errors.Range("A6").Formula = "=SQRT(-1)"
    errors.Range("A7").Formula2 = "=SEQUENCE(2)"
    errors.Range("A8").Value = "blocks the spill"
    errors.Range("B1").Value = "'#N/A"  # text that looks like an error (apostrophe keeps it text)

    dyn = ws["Dyn"]
    dyn.Range("A1").Formula2 = "=SEQUENCE(5)"
    dyn.Range("C1").Formula2 = "=SORT(Inputs!A1:A10,,-1)"
    dyn.Range("E1").Formula2 = "=SUM(A1#)"

    ws["Hidden"].Range("A1").Value = 42
    ws["Hidden"].Visible = XL_SHEET_HIDDEN
    ws["VeryHidden"].Range("A1").Formula = "=Inputs!A1"
    ws["VeryHidden"].Visible = XL_SHEET_VERY_HIDDEN

    protected = ws["Protected"]
    protected.Range("A1").Formula = "=Inputs!A1+100"
    protected.Range("A1").FormulaHidden = True
    protected.Protect()

    ws["Bulk"].Range("A1:J1000").Formula = "=Inputs!$A$1*ROW()+COLUMN()"

    workbook.Worksheets("Inputs").Activate()
    path.unlink(missing_ok=True)
    workbook.SaveAs(str(path), FileFormat=XL_OPEN_XML_WORKBOOK)
    workbook.Close(SaveChanges=False)


def build_fixture_b(excel, path: Path) -> None:
    workbook = excel.Workbooks.Add()
    workbook.Worksheets(1).Name = "Inputs"
    workbook.Worksheets(1).Range("A1").Value = 5555
    path.unlink(missing_ok=True)
    workbook.SaveAs(str(path), FileFormat=XL_OPEN_XML_WORKBOOK)
    workbook.Close(SaveChanges=False)


def main() -> None:
    FIXTURES.mkdir(exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        build_fixture_a(excel, FIXTURES / "E2_FixtureA.xlsx", input_base=10)
        build_fixture_a(excel, FIXTURES / "E2_FixtureA_decoy.xlsx", input_base=77)
        build_fixture_b(excel, FIXTURES / "E2_FixtureB.xlsx")
    finally:
        excel.Quit()
    print(f"Fixtures written to {FIXTURES}")


if __name__ == "__main__":
    main()
