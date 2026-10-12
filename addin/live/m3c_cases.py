"""Clear Sheet References (M3c): the rewrite cases, shared by the offline tests and the live check.

Each case is (sheet the command runs on, formula before, formula expected after). Only references that
qualify the CURRENT sheet lose their prefix; other sheets, external workbooks, 3D ranges, string
literals, error values and similarly named sheets are left exactly as they are. Sheet names compare
case-insensitively, as in Excel.
"""

REWRITE_CASES = [
    # plain and quoted own-sheet references
    ("Sheet1", "=Sheet1!A1", "=A1"),
    ("Sheet1", "=sheet1!A1+SHEET1!B2", "=A1+B2"),
    ("Sheet1", "=SUM(Sheet1!A1:B5)", "=SUM(A1:B5)"),
    ("Sheet1", "=Sheet1!A1:Sheet1!B2", "=A1:B2"),
    ("P&L", "='P&L'!A1", "=A1"),
    ("P&L", "=SUM('P&L'!$A$1:$C$3)*2", "=SUM($A$1:$C$3)*2"),
    ("My Sheet", "='My Sheet'!A1+'my sheet'!B1", "=A1+B1"),
    ("2024-25", "='2024-25'!A1", "=A1"),
    ("O'Brien", "='O''Brien'!A1", "=A1"),
    ("Données", "=Données!A1", "=A1"),
    ("Sheet.1", "=Sheet.1!A1", "=A1"),
    ("Sheet1", "=Sheet1!MyName", "=MyName"),
    ("Sheet1", "=Sheet1!A:A", "=A:A"),
    ("Sheet1", "=INDEX(Sheet1!A1:A9,Sheet1!B1)", "=INDEX(A1:A9,B1)"),
    # other sheets and similar names: untouched
    ("Sheet1", "=Sheet2!A1", "=Sheet2!A1"),
    ("Sheet1", "=MySheet1!A1", "=MySheet1!A1"),
    ("Sheet1", "=Sheet11!A1", "=Sheet11!A1"),
    ("P&L", "='P&L 2'!A1", "='P&L 2'!A1"),
    ("P&L", "='P&L'!A1+'P&L 2'!A1", "=A1+'P&L 2'!A1"),
    ("O'Brien", "='O''Brien 2'!A1", "='O''Brien 2'!A1"),
    # external workbooks: untouched
    ("Sheet1", "=[Book.xlsx]Sheet1!A1", "=[Book.xlsx]Sheet1!A1"),
    ("Sheet1", "=[1]Sheet1!A1", "=[1]Sheet1!A1"),
    ("P&L", "='[Book.xlsx]P&L'!A1", "='[Book.xlsx]P&L'!A1"),
    ("P&L", "='C:\\Data\\[Book.xlsx]P&L'!A1", "='C:\\Data\\[Book.xlsx]P&L'!A1"),
    # ranges whose second half is prefixed: the prefix goes
    ("Sheet1", "=SUM(A1:Sheet1!B2)", "=SUM(A1:B2)"),
    ("Sheet1", "=SUM($A$1:Sheet1!$B$2)", "=SUM($A$1:$B$2)"),
    ("Sheet1", "=SUM(Sheet1!A:Sheet1!C)", "=SUM(A:C)"),
    ("Sheet1", "=SUM(Sheet1!1:Sheet1!3)", "=SUM(1:3)"),
    ("Sheet1", "=SUM(INDEX(A:A,1):Sheet1!B9)", "=SUM(INDEX(A:A,1):B9)"),
    # 3D references: untouched
    ("Sheet1", "=SUM(Sheet1:Sheet3!A1)", "=SUM(Sheet1:Sheet3!A1)"),
    ("Sheet3", "=SUM(Sheet1:Sheet3!A1)", "=SUM(Sheet1:Sheet3!A1)"),
    ("Sheet3", "=SUM(Data:Sheet3!A1)", "=SUM(Data:Sheet3!A1)"),
    ("My Sheet", "=SUM('My Sheet:Other'!A1)", "=SUM('My Sheet:Other'!A1)"),
    # string literals: untouched, inside or outside functions
    ("Sheet1", '="Sheet1!A1"', '="Sheet1!A1"'),
    ("Sheet1", '=INDIRECT("Sheet1!A1")+Sheet1!B1', '=INDIRECT("Sheet1!A1")+B1'),
    ("P&L", "=\"'P&L'!A1\"&'P&L'!B1", "=\"'P&L'!A1\"&B1"),
    ("Sheet1", '="say ""Sheet1!A1"""&Sheet1!C1', '="say ""Sheet1!A1"""&C1'),
    # error values and functions: untouched
    ("REF", "=#REF!A1", "=#REF!A1"),
    ("REF", "=REF!A1+#REF!", "=A1+#REF!"),
    ("SUM", "=SUM(A1)", "=SUM(A1)"),
    # nothing to do
    ("Sheet1", "=A1+B1", "=A1+B1"),
]
