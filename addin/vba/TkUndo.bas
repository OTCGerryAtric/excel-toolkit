Attribute VB_Name = "TkUndo"
' Excel Toolkit (M2): one-level undo for toolkit commands.
'
' Holds a single record in memory (nothing is written to disk). Ctrl+Z - registered with
' Application.OnUndo right after a command - and the Undo Toolkit Action button run the SAME
' checked restore: workbook/sheet identity, then a conflict check on the changed cells, then
' a verified restore. Excel withdraws the Ctrl+Z entry as soon as the user does anything
' undoable; the add-in cannot see that, so the button stays available while a record exists.
' Excel's earlier undo history is lost when any macro changes cells; it is not preserved.
Option Explicit

Private current As TkCellSnapshot
Private currentLabel As String
Private unusable As Boolean           ' set after an incomplete restore; the record is kept for reporting only
Private lastRestoreSeconds As Double

Public Sub Tk_SetRecord(ByVal snap As TkCellSnapshot, ByVal label As String)
    Set current = snap
    currentLabel = label
    unusable = False
    Tk_RefreshRibbon
    Application.OnUndo "Undo " & label, "Tk_UndoFromCtrlZ"     ' last: nothing may change after it
End Sub

Public Sub Tk_ClearRecord()
    Set current = Nothing
    currentLabel = ""
    unusable = False
    Tk_RefreshRibbon
End Sub

Public Function Tk_HasRecord() As Boolean
    Tk_HasRecord = Not current Is Nothing
End Function

Public Function Tk_RecordDescription() As String
    If current Is Nothing Then Exit Function
    Tk_RecordDescription = current.Description()
    If unusable Then Tk_RecordDescription = Tk_RecordDescription & " - unusable after an incomplete restore"
End Function

Public Function Tk_LastRestoreSeconds() As Double
    Tk_LastRestoreSeconds = lastRestoreSeconds
End Function

' Called by Excel when the user presses Ctrl+Z straight after a toolkit command.
Public Sub Tk_UndoFromCtrlZ()
    Tk_Notify Tk_RestoreLast()
End Sub

' The ribbon's Undo Toolkit Action.
Public Sub Tk_UndoToolkitAction()
    Tk_Notify Tk_RestoreLast()
End Sub

' Returns "ok: / nothing: / refused: / incomplete: <text>". Never raises.
Public Function Tk_RestoreLast() As String
    Dim reason As String, saved As Variant, t0 As Double
    On Error GoTo Failed
    If current Is Nothing Then Tk_RestoreLast = "nothing: There is no Excel Toolkit change to undo.": Exit Function
    If unusable Then
        Tk_RestoreLast = "refused: The last undo stopped part-way, so this change can no longer be undone safely."
        Exit Function
    End If

    reason = current.CheckIdentity()
    If Len(reason) > 0 Then
        Tk_ClearRecord                                       ' the target is gone for good
        Tk_RestoreLast = "refused: " & reason & " Nothing was changed."
        Exit Function
    End If
    reason = current.FindConflict()
    If Len(reason) > 0 Then                                  ' kept: reverting the edit makes undo possible again
        Tk_RestoreLast = "refused: " & reason & " Nothing was changed."
        Exit Function
    End If

    t0 = Tk_Seconds()
    saved = Tk_QuietExcel()
    reason = current.RestoreFormulas(True)
    Tk_RestoreExcel saved
    lastRestoreSeconds = Tk_Seconds() - t0

    If Len(reason) = 0 Then
        Tk_RestoreLast = "ok: Undid " & currentLabel & "."
        Tk_ClearRecord
    Else
        unusable = True
        Tk_RefreshRibbon
        Tk_RestoreLast = "incomplete: " & reason
    End If
    Exit Function
Failed:
    If IsArray(saved) Then Tk_RestoreExcel saved
    Tk_RestoreLast = "refused: Excel reported an error before anything was restored: " & Err.Description
End Function
