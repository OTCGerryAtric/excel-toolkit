"""Create disposable E3 fixtures in a separate hidden Excel instance.

E3_Fixture.xlsx - Data (formulas and awkward constants), Prot (protected, with unlocked
                  and locked cells), and a mix of names for the destructive-command test.
E3_Big.xlsx     - 3,000 formulas, for the 'Excel closes mid-operation' test.
"""

from pathlib import Path

import win32com.client

HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"
XL_OPEN_XML_WORKBOOK = 51


def build_main(excel, path: Path) -> None:
    workbook = excel.Workbooks.Add()
    while workbook.Worksheets.Count < 2:
        workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    data, prot = workbook.Worksheets(1), workbook.Worksheets(2)
    data.Name, prot.Name = "Data", "Prot"

    data.Range("A1:A5").Formula = "=ROW()*10"
    data.Range("B1").Value = "'0012"      # text that looks like a number
    data.Range("B2").Value = "'1/2"       # text that looks like a date
    data.Range("B3").Formula = "=NA()"    # error from a formula
    data.Range("B4").Value = True
    data.Range("B5").Value = 3.25
    data.Range("C1:C5").Formula = "=A1+1"
    data.Range("C1:C5").NumberFormat = "#,##0.0"
    data.Range("D1").Formula = "=SUM(A1:A5)"

    prot.Range("A1:A3").Formula = "=ROW()"
    prot.Range("C1:C3").Formula = "=ROW()*2"
    prot.Range("A1:A3").Locked = False
    prot.Protect()

    workbook.Names.Add("Rate", "=Data!$B$5")
    workbook.Names.Add("HiddenName", "=Data!$A$1", False)
    workbook.Names.Add("Broken", "=#REF!")
    data.Names.Add("LocalName", "=Data!$C$1")
    data.PageSetup.PrintArea = "$A$1:$D$5"

    data.Activate()
    path.unlink(missing_ok=True)
    workbook.SaveAs(str(path), FileFormat=XL_OPEN_XML_WORKBOOK)
    workbook.Close(SaveChanges=False)


def build_big(excel, path: Path) -> None:
    workbook = excel.Workbooks.Add()
    workbook.Worksheets(1).Name = "Data"
    workbook.Worksheets(1).Range("A1:C1000").Formula = "=ROW()+COLUMN()"
    path.unlink(missing_ok=True)
    workbook.SaveAs(str(path), FileFormat=XL_OPEN_XML_WORKBOOK)
    workbook.Close(SaveChanges=False)


def main() -> None:
    FIXTURES.mkdir(exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        build_main(excel, FIXTURES / "E3_Fixture.xlsx")
        build_big(excel, FIXTURES / "E3_Big.xlsx")
    finally:
        excel.Quit()
    print(f"Fixtures written to {FIXTURES}")


if __name__ == "__main__":
    main()
