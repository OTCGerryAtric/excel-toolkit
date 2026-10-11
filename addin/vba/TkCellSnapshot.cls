VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TkCellSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Excel Toolkit (M2): the before/after state of the formula cells changed by Convert to Values.
'
' Only cells that held a formula are recorded and written; constants are never touched, so they
' cannot be retyped. Per area it keeps the original Formula2 text, the exact result written
' (same VarType and value as the formula's result), and the original number format of text
' results (the only cells whose format is touched, briefly, so text is stored verbatim).
' Restoring writes the original formulas back to those cells only, and is verified.
Option Explicit

Private wb As Workbook
Private ws As Worksheet
Private wbFullName As String
Private wsName As String
Private targetAddress As String
Private capturedAt As Date

Private nAreas As Long
Private areas() As Range
Private formulas() As Variant      ' per area: 2D array of original Formula2 ("" for constants)
Private results() As Variant       ' per area: 2D array of the value written
Private isFormula() As Variant     ' per area: 2D Boolean array
Private textFormats() As Variant   ' per area: 2D array of NumberFormat for text results
Private allFormulas() As Boolean   ' per area: every cell held a formula
Private hasText() As Boolean       ' per area: at least one text result
Private areaFormats() As Variant   ' per area: NumberFormat before (Null when mixed)
Private nFormulas As Long

' ---------- Capture (no writes) ----------

' Returns "" or the reason the selection is refused.
Public Function Capture(ByVal target As Range) As String
    Dim a As Long, r As Long, c As Long, rows As Long, cols As Long
    Dim area As Range, cell As Range, f As Variant, v As Variant, flags() As Boolean, fmts As Variant
    Dim reason As String, whole As Variant

    Set ws = target.Worksheet
    Set wb = ws.Parent
    wbFullName = wb.FullName
    wsName = ws.Name
    targetAddress = target.Address(False, False)
    If Len(targetAddress) > 60 Then targetAddress = Left$(targetAddress, 57) & "..."
    capturedAt = Now

    nAreas = target.Areas.Count
    ReDim areas(1 To nAreas): ReDim formulas(1 To nAreas): ReDim results(1 To nAreas)
    ReDim isFormula(1 To nAreas): ReDim textFormats(1 To nAreas)
    ReDim allFormulas(1 To nAreas): ReDim hasText(1 To nAreas): ReDim areaFormats(1 To nAreas)
    nFormulas = 0

    For a = 1 To nAreas
        Set area = target.Areas(a)
        Set areas(a) = area
        rows = area.Rows.Count: cols = area.Columns.Count
        f = As2D(area.Formula2, rows, cols)
        v = As2D(area.Value2, rows, cols)
        ReDim flags(1 To rows, 1 To cols)
        ReDim fmts(1 To rows, 1 To cols)
        whole = area.HasFormula                    ' True: every cell is a formula
        allFormulas(a) = (Not IsNull(whole)) And (whole = True)
        hasText(a) = False

        For r = 1 To rows
            For c = 1 To cols
                If Left$(f(r, c), 1) = "=" Then
                    Set cell = area.Cells(r, c)
                    If allFormulas(a) Then
                        flags(r, c) = True
                    Else
                        flags(r, c) = cell.HasFormula     ' a text constant may also start with "="
                    End If
                    If flags(r, c) Then
                        reason = CheckFormulaCell(cell, CStr(f(r, c)), v(r, c))
                        If Len(reason) > 0 Then Capture = reason: Exit Function
                        nFormulas = nFormulas + 1
                        If VarType(v(r, c)) = vbString Then
                            hasText(a) = True
                            fmts(r, c) = cell.NumberFormat
                        End If
                    ElseIf IsSpilledCell(cell) Then
                        Capture = SpilledMessage(area, r, c): Exit Function
                    End If
                ElseIf Len(f(r, c)) = 0 And Not IsEmpty(v(r, c)) Then
                    ' A value with no formula text behind it may be part of a spill range.
                    If IsSpilledCell(area.Cells(r, c)) Then Capture = SpilledMessage(area, r, c): Exit Function
                End If
            Next c
        Next r

        formulas(a) = f
        results(a) = v
        isFormula(a) = flags
        textFormats(a) = fmts
        areaFormats(a) = area.NumberFormat
    Next a
End Function

Private Function CheckFormulaCell(ByVal cell As Range, ByVal formulaText As String, ByVal result As Variant) As String
    If UCase$(Left$(formulaText, 7)) = "=TABLE(" Then
        CheckFormulaCell = Where(cell, 1, 1) & " is part of a data table. Convert to Values does not support data tables yet."
    ElseIf cell.HasArray Then
        CheckFormulaCell = Where(cell, 1, 1) & " holds an array formula. Convert to Values does not support array formulas yet."
    ElseIf IsSpill(cell) Then
        CheckFormulaCell = Where(cell, 1, 1) & " is a dynamic-array formula that spills. Convert to Values does not support spill ranges yet."
    ElseIf ws.ProtectContents And cell.Locked Then
        CheckFormulaCell = Where(cell, 1, 1) & " is locked on a protected sheet. Unprotect the sheet first; nothing was changed."
    Else
        Select Case VarType(result)
        Case vbDouble, vbString, vbBoolean
        Case vbError
            Select Case ErrorCode(result)
            Case 2000, 2007, 2015, 2023, 2029, 2036, 2042     ' #NULL! #DIV/0! #VALUE! #REF! #NAME? #NUM! #N/A
            Case Else
                CheckFormulaCell = Where(cell, 1, 1) & " shows an error value that Excel cannot store as a constant."
            End Select
        Case Else
            CheckFormulaCell = Where(cell, 1, 1) & " has a result of a kind Convert to Values does not support yet."
        End Select
    End If
End Function

Private Function IsSpill(ByVal cell As Range) As Boolean
    Dim spill As Variant
    spill = SafeHasSpill(cell)
    If IsNull(spill) Then IsSpill = True Else IsSpill = CBool(spill) Or IsSpilledCell(cell)
End Function

' True for a cell inside a dynamic-array spill range (including the anchor).
Private Function IsSpilledCell(ByVal cell As Range) As Boolean
    Dim anchor As Range
    On Error GoTo NotSpilled
    Set anchor = cell.SpillParent
    IsSpilledCell = Not anchor Is Nothing
NotSpilled:
End Function

Private Function SpilledMessage(ByVal area As Range, ByVal r As Long, ByVal c As Long) As String
    SpilledMessage = "The selection includes " & Where(area, r, c) & ", which shows a value spilled from a " & _
        "dynamic-array formula. Convert to Values does not support spill ranges yet."
End Function

Public Property Get FormulaCount() As Long
    FormulaCount = nFormulas
End Property

Public Function Description() As String
    Description = "Convert to Values on '" & wsName & "'!" & targetAddress & " (" & nFormulas & _
                  " cell(s), " & Format(capturedAt, "hh:nn") & ")"
End Function

' ---------- Apply and verify ----------

' Writes each formula cell's result. Raises on failure; the caller rolls back.
Public Sub ApplyValues()
    Dim a As Long, r As Long, c As Long, written As Long, limit As Long
    Dim v As Variant, flags As Variant, fmts As Variant
    limit = Tk_TestFailAfterWrites()
    For a = 1 To nAreas
        If allFormulas(a) And Not hasText(a) Then
            ' Numbers, booleans and errors cannot be re-typed by a bulk write.
            If limit > 0 And written + areas(a).CountLarge > limit Then SimulateFailure
            If areas(a).CountLarge = 1 Then areas(a).Value2 = results(a)(1, 1) Else areas(a).Value2 = results(a)
            written = written + areas(a).CountLarge
        Else
            v = results(a): flags = isFormula(a): fmts = textFormats(a)
            For r = 1 To UBound(v, 1)
                For c = 1 To UBound(v, 2)
                    If flags(r, c) Then
                        If limit > 0 And written >= limit Then SimulateFailure
                        WriteResult areas(a).Cells(r, c), v(r, c), fmts(r, c)
                        written = written + 1
                    End If
                Next c
            Next r
        End If
    Next a
End Sub

Private Sub WriteResult(ByVal cell As Range, ByVal value As Variant, ByVal originalFormat As Variant)
    If VarType(value) = vbString Then
        If Len(value) = 0 Then
            cell.Value2 = "'"                  ' empty text stays text, not a blank cell
        ElseIf Left$(value, 1) = "'" Then
            cell.Value2 = "'" & value          ' Excel swallows one leading apostrophe even in a Text cell
        Else
            cell.NumberFormat = "@"            ' store the text verbatim: no number/date parsing
            cell.Value2 = value
            cell.NumberFormat = originalFormat
        End If
    Else
        cell.Value2 = value
    End If
End Sub

Private Sub SimulateFailure()
    Err.Raise vbObjectError + 513, "TkCellSnapshot", "Simulated failure (test hook)"
End Sub

' Returns "" when every converted cell holds exactly the formula's result, else the problem.
Public Function VerifyConverted() As String
    Dim a As Long, r As Long, c As Long, now2 As Variant, f2 As Variant, flags As Variant, v As Variant, fmts As Variant
    For a = 1 To nAreas
        now2 = As2D(areas(a).Value2, areas(a).Rows.Count, areas(a).Columns.Count)
        f2 = As2D(areas(a).Formula2, areas(a).Rows.Count, areas(a).Columns.Count)
        flags = isFormula(a): v = results(a): fmts = textFormats(a)
        For r = 1 To UBound(v, 1)
            For c = 1 To UBound(v, 2)
                If flags(r, c) Then
                    If Not SameValue(now2(r, c), v(r, c)) Then
                        VerifyConverted = Where(areas(a), r, c) & " did not receive exactly its formula's result.": Exit Function
                    End If
                    If Left$(f2(r, c), 1) = "=" Then
                        If areas(a).Cells(r, c).HasFormula Then
                            VerifyConverted = Where(areas(a), r, c) & " still holds a formula.": Exit Function
                        End If
                    End If
                    If IsNonEmptyText(v(r, c)) Then
                        If areas(a).Cells(r, c).NumberFormat <> fmts(r, c) Then
                            VerifyConverted = Where(areas(a), r, c) & " number format changed.": Exit Function
                        End If
                    End If
                End If
            Next c
        Next r
        If Not IsNull(areaFormats(a)) Then
            If IsNull(areas(a).NumberFormat) Then
                VerifyConverted = "Number formats in " & areas(a).Address(False, False) & " changed.": Exit Function
            ElseIf areas(a).NumberFormat <> areaFormats(a) Then
                VerifyConverted = "Number formats in " & areas(a).Address(False, False) & " changed.": Exit Function
            End If
        End If
    Next a
End Function

' ---------- Undo ----------

' Returns "" when the workbook and sheet are still the ones that were changed.
Public Function CheckIdentity() As String
    Dim current As String
    On Error GoTo WorkbookGone
    current = wb.FullName
    On Error GoTo 0
    If current <> wbFullName Then
        CheckIdentity = "The workbook has been renamed or saved under another name since the change (it was " & wbFullName & ")."
        Exit Function
    End If
    On Error GoTo SheetGone
    current = ws.Name
    On Error GoTo 0
    If current <> wsName Then
        CheckIdentity = "The sheet has been renamed since the change (it was '" & wsName & "')."
    End If
    Exit Function
WorkbookGone:
    CheckIdentity = "The workbook that was changed (" & wbFullName & ") is no longer open."
    Exit Function
SheetGone:
    CheckIdentity = "The sheet that was changed ('" & wsName & "') no longer exists."
End Function

' Returns "" when every converted cell still holds exactly what was written, else the first
' conflicting cell. Cells outside the converted set are neither checked nor touched.
Public Function FindConflict() As String
    Dim a As Long, r As Long, c As Long, now2 As Variant, f2 As Variant, flags As Variant, v As Variant
    For a = 1 To nAreas
        now2 = As2D(areas(a).Value2, areas(a).Rows.Count, areas(a).Columns.Count)
        f2 = As2D(areas(a).Formula2, areas(a).Rows.Count, areas(a).Columns.Count)
        flags = isFormula(a): v = results(a)
        For r = 1 To UBound(v, 1)
            For c = 1 To UBound(v, 2)
                If flags(r, c) Then
                    If Not SameValue(now2(r, c), v(r, c)) Then
                        FindConflict = Where(areas(a), r, c) & " has changed since the conversion.": Exit Function
                    End If
                    If Left$(f2(r, c), 1) = "=" Then
                        If areas(a).Cells(r, c).HasFormula Then
                            FindConflict = Where(areas(a), r, c) & " has changed since the conversion.": Exit Function
                        End If
                    End If
                    If ws.ProtectContents Then
                        If areas(a).Cells(r, c).Locked Then
                            FindConflict = "The sheet is now protected (" & Where(areas(a), r, c) & " is locked). Unprotect it first.": Exit Function
                        End If
                    End If
                End If
            Next c
        Next r
    Next a
End Function

' Writes the original formulas (and the formats of text results) back. Never raises; returns ""
' on verified success, else which cells were not restored. useHook applies the restore test hook.
Public Function RestoreFormulas(ByVal useHook As Boolean) As String
    Dim a As Long, r As Long, c As Long, flags As Variant, f As Variant, fmts As Variant, v As Variant
    Dim problems As String, done As Long, limit As Long
    If useHook Then limit = Tk_TestFailAfterRestores()
    For a = 1 To nAreas
        f = formulas(a): flags = isFormula(a): fmts = textFormats(a): v = results(a)
        If allFormulas(a) And Not hasText(a) And limit = 0 Then
            On Error Resume Next
            Err.Clear
            If areas(a).CountLarge = 1 Then areas(a).Formula2 = f(1, 1) Else areas(a).Formula2 = f
            If Err.Number <> 0 Then problems = problems & " " & areas(a).Address(False, False) & " (" & Err.Description & ")"
            On Error GoTo 0
        Else
            For r = 1 To UBound(f, 1)
                For c = 1 To UBound(f, 2)
                    If flags(r, c) Then
                        On Error Resume Next
                        Err.Clear
                        If limit > 0 And done >= limit Then Err.Raise vbObjectError + 514, "TkCellSnapshot", "Simulated restore failure (test hook)"
                        If Err.Number = 0 Then
                            areas(a).Cells(r, c).Formula2 = f(r, c)
                            If IsNonEmptyText(v(r, c)) Then areas(a).Cells(r, c).NumberFormat = fmts(r, c)
                        End If
                        If Err.Number <> 0 Then
                            problems = problems & " " & Where(areas(a), r, c)
                        Else
                            done = done + 1
                        End If
                        On Error GoTo 0
                    End If
                Next c
            Next r
        End If
    Next a
    If Len(problems) > 0 Then RestoreFormulas = "Not restored:" & problems & ".": Exit Function
    RestoreFormulas = VerifyRestored()
End Function

' Restores after a failed apply (no test hook).
Public Function RollBack() As String
    RollBack = RestoreFormulas(False)
End Function

Private Function VerifyRestored() As String
    Dim a As Long, r As Long, c As Long, f2 As Variant, flags As Variant, f As Variant, problems As String
    On Error GoTo Unreadable
    For a = 1 To nAreas
        f2 = As2D(areas(a).Formula2, areas(a).Rows.Count, areas(a).Columns.Count)
        flags = isFormula(a): f = formulas(a)
        For r = 1 To UBound(f, 1)
            For c = 1 To UBound(f, 2)
                If flags(r, c) Then
                    If CStr(f2(r, c)) <> CStr(f(r, c)) Then problems = problems & " " & Where(areas(a), r, c)
                End If
            Next c
        Next r
    Next a
    If Len(problems) > 0 Then VerifyRestored = "Formula not restored exactly:" & problems & "."
    Exit Function
Unreadable:
    VerifyRestored = "The restored cells could not be read back to verify them: " & Err.Description
End Function

' ---------- Helpers ----------

Private Function As2D(ByVal value As Variant, ByVal rows As Long, ByVal cols As Long) As Variant
    Dim one(1 To 1, 1 To 1) As Variant
    If rows = 1 And cols = 1 Then
        one(1, 1) = value
        As2D = one
    Else
        As2D = value
    End If
End Function

' VBA's And evaluates both sides, so Len() must never see an error value: test the type first.
Private Function IsNonEmptyText(ByVal value As Variant) As Boolean
    If VarType(value) = vbString Then IsNonEmptyText = (Len(value) > 0)
End Function

' Same type and value. Matching displayed text is not enough.
Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If VarType(a) <> VarType(b) Then Exit Function
    If VarType(a) = vbError Then
        SameValue = (ErrorCode(a) = ErrorCode(b))
    ElseIf VarType(a) = vbString Then
        SameValue = (StrComp(a, b, vbBinaryCompare) = 0)
    Else
        SameValue = (a = b)
    End If
End Function

Private Function ErrorCode(ByVal value As Variant) As Long
    ErrorCode = CLng(Val(Mid$(CStr(value), 7)))        ' CStr(CVErr(2042)) = "Error 2042"
End Function

Private Function Where(ByVal area As Range, ByVal r As Long, ByVal c As Long) As String
    Where = "cell " & area.Cells(r, c).Address(False, False) & " on '" & wsName & "'"
End Function
