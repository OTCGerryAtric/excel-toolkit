VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TkFormulaSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Excel Toolkit (M3c): the formulas Clear Sheet References is about to rewrite, for undo.
'
' Only formula cells are looked at (SpecialCells(xlCellTypeFormulas)), so text and other constants
' are never read or written. For each formula whose rewrite differs, the cell, its formula before
' (Formula2, exactly as Excel reports it) and the rewritten formula are kept; after writing, the
' formula as Excel stored it is kept for the conflict check. Cells in a legacy array formula (CSE)
' are refused rather than rewritten one by one. Undo-record interface: Description,
' CheckIdentity, FindConflict, Restore (see TkUndo).
Option Explicit

Private Const MAX_CHANGED As Long = 100000

Private wb As Workbook
Private ws As Worksheet
Private wbFullName As String
Private wsName As String
Private targetAddress As String
Private commandLabel As String
Private capturedAt As Date
Private sheetName As String
Private changed As Collection         ' Range (one cell) per rewritten formula
Private before As Collection          ' its formula before the command
Private rewritten As Collection       ' the formula the command writes
Private after As Collection           ' the formula as Excel stored it afterwards

' ---------- Capture (no writes) ----------

' Returns "" or the reason the command is refused. CellCount = 0 means nothing to do.
Public Function Capture(ByVal scope As Range, ByVal ownSheet As String, ByVal label As String) As String
    Dim formulaCells As Range, area As Range, values As Variant, r As Long, c As Long, oldText As String, newText As String
    Set ws = scope.Worksheet
    Set wb = ws.Parent
    wbFullName = wb.FullName
    wsName = ws.Name
    sheetName = ownSheet
    targetAddress = scope.Address(False, False)
    If Len(targetAddress) > 60 Then targetAddress = Left$(targetAddress, 57) & "..."
    commandLabel = label
    capturedAt = Now
    Set changed = New Collection
    Set before = New Collection
    Set rewritten = New Collection
    Set after = New Collection

    On Error Resume Next
    Set formulaCells = scope.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaCells Is Nothing Then Exit Function

    For Each area In formulaCells.Areas
        values = area.Formula2
        If Not IsArray(values) Then values = OneCell(values)
        For r = 1 To UBound(values, 1)
            For c = 1 To UBound(values, 2)
                oldText = CStr(values(r, c))
                If InStr(1, oldText, "!") > 0 Then
                    newText = Tk_RemoveSheetPrefix(oldText, sheetName)
                    If newText <> oldText Then
                        If Not area.Cells(r, c).HasFormula Then GoTo NextCell      ' a spilled value, not a formula
                        If area.Cells(r, c).HasArray Then
                            Capture = "The cell " & area.Cells(r, c).Address(False, False) & " is part of an array formula " & _
                                "(entered with Ctrl+Shift+Enter), which is not supported yet; nothing was changed."
                            Exit Function
                        End If
                        If changed.Count >= MAX_CHANGED Then
                            Capture = "More than " & Format(MAX_CHANGED, "#,##0") & " formulas would change; select a smaller " & _
                                "range. Nothing was changed."
                            Exit Function
                        End If
                        changed.Add area.Cells(r, c)
                        before.Add oldText
                        rewritten.Add newText
                    End If
                End If
NextCell:
            Next c
        Next r
    Next area
End Function

Private Function OneCell(ByVal value As Variant) As Variant
    Dim grid(1 To 1, 1 To 1) As Variant
    grid(1, 1) = value
    OneCell = grid
End Function

Public Property Get CellCount() As Long
    CellCount = changed.Count
End Property

Public Function Description() As String
    Description = commandLabel & " on '" & wsName & "'!" & targetAddress & " (" & Format(capturedAt, "hh:nn") & ")"
End Function

' ---------- Apply and verify ----------

' Raises on failure; the caller rolls back.
Public Sub Apply()
    Dim i As Long, limit As Long
    limit = Tk_TestFailAfterWrites()
    For i = 1 To changed.Count
        If limit > 0 And i > limit Then Err.Raise vbObjectError + 516, "TkFormulaSnapshot", "Simulated failure (test hook)"
        changed(i).Formula2 = rewritten(i)
    Next i
End Sub

' "" when every cell still holds a formula with no prefix naming this sheet; records what Excel stored.
Public Function VerifyApplied() As String
    Dim i As Long, stored As String
    Set after = New Collection
    For i = 1 To changed.Count
        If Not changed(i).HasFormula Then
            VerifyApplied = "The formula in " & changed(i).Address(False, False) & " was not written."
            Exit Function
        End If
        stored = changed(i).Formula2
        If Tk_RemoveSheetPrefix(stored, sheetName) <> stored Then
            VerifyApplied = "The formula in " & changed(i).Address(False, False) & " still refers to '" & sheetName & "' by name."
            Exit Function
        End If
        after.Add stored
    Next i
End Function

' ---------- Undo ----------

Public Function CheckIdentity() As String
    CheckIdentity = Tk_CheckTargetIdentity(wb, ws, wbFullName, wsName)
End Function

' "" when every rewritten cell still holds the formula the command left there, else what changed.
Public Function FindConflict() As String
    Dim i As Long
    If ws.ProtectContents Then
        FindConflict = "The sheet is now protected, so its formulas cannot be changed. Unprotect it first."
        Exit Function
    End If
    For i = 1 To changed.Count
        If changed(i).Formula2 <> after(i) Then
            FindConflict = "The formula in " & changed(i).Address(False, False) & " on '" & wsName & "' has changed since the command."
            Exit Function
        End If
    Next i
End Function

' Puts every original formula back. Never raises; "" on verified success.
Public Function Restore(ByVal useHook As Boolean) As String
    Dim i As Long, problems As String, limit As Long, done As Long
    If useHook Then limit = Tk_TestFailAfterRestores()
    For i = 1 To changed.Count
        On Error Resume Next
        Err.Clear
        If limit > 0 And done >= limit Then Err.Raise vbObjectError + 517, "TkFormulaSnapshot", "Simulated restore failure (test hook)"
        If Err.Number = 0 Then changed(i).Formula2 = before(i)
        If Err.Number <> 0 Then
            problems = problems & " " & changed(i).Address(False, False)
        Else
            done = done + 1
        End If
        On Error GoTo 0
    Next i
    If Len(problems) > 0 Then Restore = "Formulas not restored:" & problems & ".": Exit Function
    Restore = VerifyRestored()
End Function

Private Function VerifyRestored() As String
    Dim i As Long, problems As String
    On Error GoTo Unreadable
    For i = 1 To changed.Count
        If changed(i).Formula2 <> before(i) Then problems = problems & " " & changed(i).Address(False, False)
    Next i
    If Len(problems) > 0 Then VerifyRestored = "Formulas not restored exactly:" & problems & "."
    Exit Function
Unreadable:
    VerifyRestored = "The restored cells could not be read back to verify them: " & Err.Description
End Function
