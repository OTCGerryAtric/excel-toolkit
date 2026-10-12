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
    ("Sheet1", "=SUM(Sheet1!C5:Sheet1!B2)", "=SUM(C5:B2)"),
    ("Sheet1", "=SUM(XFD1:Sheet1!B2)", "=SUM(XFD1:B2)"),
    ("Sheet1", "=SUM(INDEX(A:A,1):Sheet1!B9)", "=SUM(INDEX(A:A,1):B9)"),
    ("Sheet1", "=SUM(Sheet1!A1:INDEX(A:A,3))", "=SUM(A1:INDEX(A:A,3))"),
    ("P&L", "=SUM('P&L'!A1:'P&L'!B2)", "=SUM(A1:B2)"),
    ("Sheet1", "=SUM(Sheet1!A1:A!B2)", "=SUM(A1:A!B2)"),
    # ... unless the first half is on another sheet: Sheet2!A1:B2 would mean Sheet2!B2
    ("Sheet1", "=SUM(Sheet2!A1:Sheet1!B2)", "=SUM(Sheet2!A1:Sheet1!B2)"),
    ("Sheet1", "=SUM(Sheet2!A1:'Sheet1'!B2)", "=SUM(Sheet2!A1:'Sheet1'!B2)"),
    ("P&L", "=SUM('P&L 2'!A1:'P&L'!B2)", "=SUM('P&L 2'!A1:'P&L'!B2)"),
    ("Sheet1", "=SUM([Book.xlsx]Sheet1!A1:Sheet1!B2)", "=SUM([Book.xlsx]Sheet1!A1:Sheet1!B2)"),
    # 3D references: untouched
    ("Sheet1", "=SUM(Sheet1:Sheet3!A1)", "=SUM(Sheet1:Sheet3!A1)"),
    ("Sheet3", "=SUM(Sheet1:Sheet3!A1)", "=SUM(Sheet1:Sheet3!A1)"),
    ("Sheet3", "=SUM(Data:Sheet3!A1)", "=SUM(Data:Sheet3!A1)"),
    ("My Sheet", "=SUM('My Sheet:Other'!A1)", "=SUM('My Sheet:Other'!A1)"),
    # ... including sheets named like columns or out-of-range cells, which Excel does not quote and reads
    # as 3D ranges (it stores A:Sheet1!B2 as Sheet1:A!B2); only a real cell (A1) starts an ordinary range
    ("Sheet1", "=SUM(A:Sheet1!B2)", "=SUM(A:Sheet1!B2)"),
    ("Sheet1", "=SUM(Sheet1:A!B2)", "=SUM(Sheet1:A!B2)"),
    ("A", "=SUM(Sheet1:A!B2)", "=SUM(Sheet1:A!B2)"),
    ("A", "=SUM(AB:A!B2)", "=SUM(AB:A!B2)"),
    ("Sheet1", "=SUM(XFE1:Sheet1!B2)", "=SUM(XFE1:Sheet1!B2)"),
    ("Sheet1", "=SUM(A0:Sheet1!B2)", "=SUM(A0:Sheet1!B2)"),
    ("Sheet1", "=SUM(1:Sheet1!B2)", "=SUM(1:Sheet1!B2)"),
    # sheets named like columns: their own references still go
    ("A", "=A!B2+a!C3", "=B2+C3"),
    ("AB", "=SUM(AB!A1:AB!B2)", "=SUM(A1:B2)"),
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

# Formulas that WOULD change but hold a range Excel could read two ways: "AB:Sheet1!B2" is stored the same
# whether it was entered as a 3D range or as AB:'Sheet1'!B2 (a name or column to a cell), and Excel reads
# it back as 3D. Rewriting such a formula means re-entering it, so it is left exactly as it is and
# reported. (sheet, formula)
AMBIGUOUS_CASES = [
    ("Sheet1", "=SUM(A:Sheet1!B2)+Sheet1!C1"),
    ("Sheet1", "=SUM(AB:Sheet1!B2)+Sheet1!C1"),
    ("A", "=SUM(AB:A!B2)+A!C1"),
    ("Sheet1", "=SUM(Sheet1:Data!B2)+Sheet1!C1"),
    ("Sheet1", "=SUM(1:Sheet1!B2)+'Sheet1'!C1"),
    ("Sheet1", "=Sheet1!AB:Other!B2"),            # removing Sheet1! would create AB:Other!B2
]
