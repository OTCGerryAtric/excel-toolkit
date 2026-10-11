Attribute VB_Name = "TkConvert"
' Excel Toolkit (M2): Convert to Values - replace the formulas in a selection with their results.
'
' Contract (docs/product-roadmap.md, section D):
'   validate -> capture -> apply -> verify (roll back on failure) -> register one-level undo.
' Clipboard-free. Only formula cells are written; constants are never touched. Every outcome is
' returned as "ok: / nothing: / refused: / rolled back: / incomplete: <text>" and success is
' never claimed unless the written cells were verified.
Option Explicit

Private Const DEFAULT_MAX_CELLS As Double = 100000   ' provisional safeguard, revisit after timings
Private maxCells As Double
Private failAfterWrites As Long                       ' test hook: simulate a failure (0 = off)
Private failAfterRestores As Long                     ' test hook: simulate a failed restore (0 = off)
Private lastTimings As String

' ---------- Settings and test hooks ----------

Public Sub Tk_SetMaxCells(ByVal limit As Double)
    maxCells = limit                      ' 0 restores the default
End Sub

Public Function Tk_MaxCells() As Double
    If maxCells > 0 Then Tk_MaxCells = maxCells Else Tk_MaxCells = DEFAULT_MAX_CELLS
End Function

Public Sub Tk_TestSetFailAfter(ByVal writes As Long, ByVal restores As Long)
    failAfterWrites = writes
    failAfterRestores = restores
End Sub

Public Function Tk_TestFailAfterWrites() As Long
    Tk_TestFailAfterWrites = failAfterWrites
End Function

Public Function Tk_TestFailAfterRestores() As Long
    Tk_TestFailAfterRestores = failAfterRestores
End Function

Public Function Tk_LastTimings() As String
    Tk_LastTimings = lastTimings
End Function

' ---------- Command ----------

Public Function Tk_ConvertRange(ByVal target As Range) As String
    Dim snap As TkCellSnapshot, reason As String, rollbackProblem As String, writing As Boolean
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double, t4 As Double, t5 As Double
    Dim saved As Variant

    On Error GoTo Failed
    t0 = Tk_Seconds()
    reason = ValidateSelection(target)
    If Len(reason) > 0 Then Tk_ConvertRange = "refused: " & reason: Exit Function

    t1 = Tk_Seconds()
    Set snap = New TkCellSnapshot
    reason = snap.Capture(target)
    If Len(reason) > 0 Then Tk_ConvertRange = "refused: " & reason: Exit Function
    If snap.FormulaCount = 0 Then Tk_ConvertRange = "nothing: There are no formulas in the selection.": Exit Function

    t2 = Tk_Seconds()
    saved = Tk_QuietExcel()
    writing = True
    snap.ApplyValues
    t3 = Tk_Seconds()
    reason = snap.VerifyConverted()
    If Len(reason) > 0 Then GoTo RollBack
    writing = False
    Tk_RestoreExcel saved
    t4 = Tk_Seconds()

    Tk_SetRecord snap, "Convert to Values"      ' replaces any earlier record; registers Ctrl+Z
    t5 = Tk_Seconds()
    lastTimings = "cells=" & snap.FormulaCount & ";validate=" & Ms(t1 - t0) & ";capture=" & Ms(t2 - t1) & _
                  ";apply=" & Ms(t3 - t2) & ";verify=" & Ms(t4 - t3) & ";register=" & Ms(t5 - t4) & _
                  ";total=" & Ms(t5 - t0)
    Tk_ConvertRange = "ok: " & snap.FormulaCount & " formula cell(s) converted to values."
    Exit Function

Failed:
    If Not writing Then
        Tk_ConvertRange = "refused: Excel reported an error before anything was changed: " & Err.Description
        Exit Function
    End If
    reason = "Excel reported an error while writing: " & Err.Description
    Resume RollBack
RollBack:
    On Error GoTo 0
    rollbackProblem = snap.RollBack()
    Tk_RestoreExcel saved
    If Len(rollbackProblem) = 0 Then
        Tk_ConvertRange = "rolled back: " & reason
    Else
        Tk_ConvertRange = "incomplete: " & reason & " Restoring the original formulas also failed: " & rollbackProblem
    End If
End Function

' ---------- Validation (no writes) ----------

' Returns "" when the selection may be converted, otherwise the reason it is refused.
Private Function ValidateSelection(ByVal target As Range) As String
    Dim area As Range, visible As Range, total As Double, i As Long, j As Long, n As Long
    Dim top() As Long, bottom() As Long, lft() As Long, rgt() As Long, merged As Variant, spill As Variant

    n = target.Areas.Count
    For Each area In target.Areas
        total = total + area.CountLarge
    Next area
    If total > Tk_MaxCells() Then
        ValidateSelection = "The selection has " & Format(total, "#,##0") & " cells; the current limit is " & _
            Format(Tk_MaxCells(), "#,##0") & ". Select fewer cells and try again."
        Exit Function
    End If

    ' Overlapping areas would be converted twice and make undo ambiguous.
    If n > 1 Then
        ReDim top(1 To n): ReDim bottom(1 To n): ReDim lft(1 To n): ReDim rgt(1 To n)
        For i = 1 To n
            Set area = target.Areas(i)
            top(i) = area.Row: bottom(i) = area.Row + area.Rows.Count - 1
            lft(i) = area.Column: rgt(i) = area.Column + area.Columns.Count - 1
        Next i
        For i = 1 To n - 1
            For j = i + 1 To n
                If top(i) <= bottom(j) And top(j) <= bottom(i) And lft(i) <= rgt(j) And lft(j) <= rgt(i) Then
                    ValidateSelection = "Parts of the selection overlap (" & target.Areas(i).Address(False, False) & _
                        " and " & target.Areas(j).Address(False, False) & "). Select each cell only once."
                    Exit Function
                End If
            Next j
        Next i
    End If

    For Each area In target.Areas
        ' Hidden or filtered-out cells inside the selection (an explicit visible-cells-only
        ' selection on a filtered sheet contains none, so it is allowed).
        If area.CountLarge = 1 Then
            If area.EntireRow.Hidden Or area.EntireColumn.Hidden Then ValidateSelection = HiddenMessage(): Exit Function
        Else
            Set visible = Nothing
            On Error Resume Next
            Set visible = area.SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            If visible Is Nothing Then ValidateSelection = HiddenMessage(): Exit Function
            If visible.CountLarge <> area.CountLarge Then ValidateSelection = HiddenMessage(): Exit Function
        End If

        merged = area.MergeCells
        If IsNull(merged) Then merged = True
        If merged Then
            ValidateSelection = "The selection includes merged cells (" & area.Address(False, False) & _
                "). Convert to Values does not support merged cells yet."
            Exit Function
        End If

        spill = SafeHasSpill(area)
        If IsNull(spill) Then spill = True
        If spill Then
            ValidateSelection = "The selection includes a dynamic-array spill range (" & area.Address(False, False) & _
                "). Convert to Values does not support spill ranges yet."
            Exit Function
        End If
    Next area
End Function

Private Function HiddenMessage() As String
    HiddenMessage = "The selection includes hidden or filtered-out cells, which would also be converted." & vbCrLf & _
        "To convert only the visible cells, select them with Alt+; (Home > Find & Select > Go To Special > " & _
        "Visible cells only) and try again."
End Function

' Range.HasSpill exists only in dynamic-array Excel; treat "not available" as no spill.
Public Function SafeHasSpill(ByVal rng As Object) As Variant
    On Error GoTo NoProperty
    SafeHasSpill = rng.HasSpill
    Exit Function
NoProperty:
    SafeHasSpill = False
End Function

' ---------- Excel state while writing ----------

' Screen updating, events (so the workbook's own change macros don't run on our writes) and
' calculation are suspended while writing and always restored.
Public Function Tk_QuietExcel() As Variant
    Dim saved(0 To 2) As Variant
    saved(0) = Application.ScreenUpdating
    saved(1) = Application.EnableEvents
    saved(2) = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    Tk_QuietExcel = saved
End Function

Public Sub Tk_RestoreExcel(ByVal saved As Variant)
    On Error Resume Next
    If Application.Calculation <> saved(2) Then Application.Calculation = saved(2)
    Application.EnableEvents = saved(1)
    Application.ScreenUpdating = saved(0)
End Sub

Private Function Ms(ByVal seconds As Double) As String
    Ms = Format(seconds * 1000, "0.0")
End Function
