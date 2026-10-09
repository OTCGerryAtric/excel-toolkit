"""Create disposable E4 fixtures in a separate hidden Excel instance.

E4_Fixture.xlsx       - target: Inputs, Mixed (one of each finding kind), Bulk (100,000 formulas)
E4_Fixture_decoy.xlsx - same layout, Inputs!A1 = 77; opened under the target's name in another process
E4_Other.xlsx         - second workbook in the target process, Inputs!A1 = 5555
"""

from pathlib import Path

import win32com.client

HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"
XL_OPEN_XML_WORKBOOK = 51


def build(excel, path: Path, inputs_a1: float, bulk_rows: int) -> None:
    workbook = excel.Workbooks.Add()
    while workbook.Worksheets.Count < 3:
        workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    inputs, mixed, bulk = (workbook.Worksheets(i) for i in (1, 2, 3))
    inputs.Name, mixed.Name, bulk.Name = "Inputs", "Mixed", "Bulk"
    inputs.Range("A1").Value = inputs_a1

    mixed.Range("A1").Formula = "=Inputs!A1"     # formula
    mixed.Range("A2").Formula = "=A1*1.2"        # formula with a hard-coded number
    mixed.Range("A3").Value = 5                  # number
    mixed.Range("A4").Value = "'label"           # text
    mixed.Range("A5").Formula = "=NA()"          # formula returning an error
    mixed.Range("A6").Value = True               # boolean

    bulk.Range(f"A1:J{bulk_rows}").Formula = "=Inputs!$A$1*ROW()+COLUMN()"
    inputs.Activate()
    path.unlink(missing_ok=True)
    workbook.SaveAs(str(path), FileFormat=XL_OPEN_XML_WORKBOOK)
    workbook.Close(SaveChanges=False)


def main() -> None:
    FIXTURES.mkdir(exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        build(excel, FIXTURES / "E4_Fixture.xlsx", 10, 10000)
        build(excel, FIXTURES / "E4_Fixture_decoy.xlsx", 77, 10000)
        build(excel, FIXTURES / "E4_Other.xlsx", 5555, 100)
    finally:
        excel.Quit()
    print(f"Fixtures written to {FIXTURES}")


if __name__ == "__main__":
    main()
