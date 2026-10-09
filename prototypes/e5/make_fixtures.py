"""Create the E5 dependency fixtures in a separate hidden Excel instance.

E5_External.xlsx - target of the external-workbook reference (Sheet1!A1 = 42)
E5_Deps.xlsx     - sheets Inputs, Calc, Data (table 'Sales'), My Sheet; name MyDefinedName;
                   one scenario formula per row in Calc!B1:B14 (label in column A)
"""

from pathlib import Path

import win32com.client

HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"
XL_OPEN_XML_WORKBOOK = 51

# (id, label, formula); {external} is replaced by the external workbook's folder
SCENARIOS = [
    ("S1", "same-sheet cells", "=F1+G1"),
    ("S2", "cross-sheet cell", "=Inputs!A1*2"),
    ("S3", "range", "=SUM(F2:F101)"),
    ("S4", "absolute + relative", "=$F$1+G2"),
    ("S5", "table structured reference", "=SUM(Sales[Revenue])"),
    ("S6", "defined name", "=MyDefinedName*2"),
    ("S7", "INDIRECT", "=INDIRECT(H1)"),
    ("S8", "OFFSET", "=OFFSET(F1,0,1)"),
    ("S9", "spill reference", "=SUM(D1#)"),
    ("S10", "external workbook", "='{external}\\[E5_External.xlsx]Sheet1'!$A$1"),
    ("S11", "LET local name", "=LET(x,F1,x*G1)"),
    ("S12", "3D reference", "=SUM(Inputs:Data!A1)"),
    ("S13", "whole column", "=SUM(F:F)"),
    ("S14", "quoted sheet name", "='My Sheet'!A1+1"),
]


def main() -> None:
    FIXTURES.mkdir(exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        external = excel.Workbooks.Add()
        external.Worksheets(1).Name = "Sheet1"
        external.Worksheets(1).Range("A1").Value = 42
        external_path = FIXTURES / "E5_External.xlsx"
        external_path.unlink(missing_ok=True)
        external.SaveAs(str(external_path), FileFormat=XL_OPEN_XML_WORKBOOK)

        workbook = excel.Workbooks.Add()
        while workbook.Worksheets.Count < 4:
            workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        inputs, calc, data, my_sheet = (workbook.Worksheets(i) for i in range(1, 5))
        inputs.Name, calc.Name, data.Name, my_sheet.Name = "Inputs", "Calc", "Data", "My Sheet"

        inputs.Range("A1:A100").Formula = "=ROW()*10"
        inputs.Range("A1:A100").Value = inputs.Range("A1:A100").Value
        inputs.Range("B1").Value = 7
        workbook.Names.Add("MyDefinedName", "=Inputs!$B$1")

        data.Range("A1:B1").Value = ("Region", "Revenue")
        for row, (region, revenue) in enumerate([("N", 100), ("S", 200), ("E", 300), ("W", 400), ("C", 500)], start=2):
            data.Cells(row, 1).Value = region
            data.Cells(row, 2).Value = revenue
        table = data.ListObjects.Add(1, data.Range("A1:B6"), None, 1)
        table.Name = "Sales"

        my_sheet.Range("A1").Value = 3

        calc.Range("F1").Value = 1
        calc.Range("G1").Value = 2
        calc.Range("G2").Value = 5
        calc.Range("F2:F101").Value = 1
        calc.Range("H1").Value = "Inputs!A2"
        calc.Range("D1").Formula2 = "=SEQUENCE(5)"
        for row, (scenario, label, formula) in enumerate(SCENARIOS, start=1):
            calc.Cells(row, 1).Value = f"{scenario} {label}"
            calc.Cells(row, 2).Formula2 = formula.replace("{external}", str(FIXTURES))

        external.Close(SaveChanges=False)
        calc.Activate()
        path = FIXTURES / "E5_Deps.xlsx"
        path.unlink(missing_ok=True)
        workbook.SaveAs(str(path), FileFormat=XL_OPEN_XML_WORKBOOK)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()
    print(f"Fixtures written to {FIXTURES}")


if __name__ == "__main__":
    main()
