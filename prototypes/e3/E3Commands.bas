Attribute VB_Name = "E3Commands"
' Phase 0 / E3 prototype: Excel Ctrl+Z integration via Application.OnUndo,
' identity handoff, and a way to make Excel busy for resilience tests.
' Throwaway code - not production.
Option Explicit

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

Private undoAddresses As Collection   ' external addresses, one per saved block
Private undoFormats As Collection     ' number format for each saved block
Private undoMarker As String

' ---------- Identity handoff (same identity as E2, returned instead of shelled) ----------

Public Function E3_Identity() As String
    E3_Identity = Application.Hwnd & "|" & GetCurrentProcessId() & "|" & _
                  ActiveWorkbook.Name & "|" & ActiveWorkbook.FullName
End Function

' ---------- Manual test hotkey ----------

Public Sub E3_Init()
    Application.OnKey "^+J", "E3_Number_Format_Toggle"
End Sub

' ---------- Formatting toggle with one-step Ctrl+Z ----------

Private Function NumberFormatCycle() As Variant
    NumberFormatCycle = Array( _
        "#,##0_-;[Red](#,##0)_-;-_-", _
        "#,##0.0_-;[Red](#,##0.0)_-;-_-", _
        "#,##0.00_-;[Red](#,##0.00)_-;-_-")
End Function

Public Sub E3_Number_Format_Toggle()
    Dim target As Range, area As Range, cell As Range
    Dim formats As Variant, current As Variant, i As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set target = Selection

    ' Save the current formats: one entry per area, or per cell when an area is mixed.
    Set undoAddresses = New Collection
    Set undoFormats = New Collection
    For Each area In target.Areas
        If IsNull(area.NumberFormat) Then
            For Each cell In area.Cells
                undoAddresses.Add cell.Address(External:=True)
                undoFormats.Add cell.NumberFormat
            Next cell
        Else
            undoAddresses.Add area.Address(External:=True)
            undoFormats.Add area.NumberFormat
        End If
    Next area

    formats = NumberFormatCycle()
    current = target.NumberFormat
    If Not IsNull(current) Then
        For i = LBound(formats) To UBound(formats) - 1
            If current = formats(i) Then
                target.NumberFormat = formats(i + 1)
                Application.OnUndo "Undo number format", "E3_Undo_Number_Format"
                Exit Sub
            End If
        Next i
    End If
    target.NumberFormat = formats(LBound(formats))
    Application.OnUndo "Undo number format", "E3_Undo_Number_Format"
End Sub

Public Sub E3_Undo_Number_Format()
    Dim i As Long
    If undoAddresses Is Nothing Then Exit Sub
    For i = 1 To undoAddresses.Count
        Application.Range(undoAddresses(i)).NumberFormat = undoFormats(i)
    Next i
    Set undoAddresses = Nothing
    Set undoFormats = Nothing
End Sub

' ---------- Ctrl+Z registration on behalf of a Python operation ----------

Public Sub E3_RegisterUndo(ByVal label As String)
    undoMarker = ""
    Application.OnUndo label, "E3_Undo_Marker"
End Sub

Public Sub E3_Undo_Marker()
    undoMarker = "undo requested"
End Sub

Public Function E3_GetUndoMarker() As String
    E3_GetUndoMarker = undoMarker
End Function

' ---------- Make Excel busy ----------

Public Sub E3_CallUndoFromVba()
    Application.Undo
End Sub

Public Sub E3_ScheduleBuiltInDialog()
    Application.OnTime Now, "E3_ShowBuiltInDialog"
End Sub

Public Sub E3_ShowBuiltInDialog()
    Application.Dialogs(xlDialogFormatNumber).Show
End Sub

Public Sub E3_ScheduleBlockingMessage()
    Application.OnTime Now, "E3_ShowBlockingMessage"
End Sub

Public Sub E3_ShowBlockingMessage()
    MsgBox "E3 test: modal dialog. Closed automatically by the test.", vbOKOnly, "E3 busy test"
End Sub
