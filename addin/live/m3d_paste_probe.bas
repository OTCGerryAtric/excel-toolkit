Attribute VB_Name = "M3dPasteProbe"
' M3d experiment only - NOT part of the add-in and never built into it. Imported by
' addin/live/m3d_native_paste_experiment.py into a disposable workbook in a disposable Excel.
'
' Question: does Excel's own Paste Values command, run from VBA with CommandBars.ExecuteMso,
' keep the clipboard, give a native Ctrl+Z for the paste, and leave the earlier Undo history in
' place? Each function returns "status|details" so the harness can record what happened; none
' of them writes to cells through the object model except the two controls, which show what the
' measurements look like when the Undo history IS cleared.
Option Explicit

' Excel's built-in Paste Values on the current selection.
' "ok|<CutCopyMode before>|<CutCopyMode after>", "disabled|<CutCopyMode>" or "error|<number>|<text>".
Public Function Probe_NativePasteValues() As String
    Dim modeBefore As Long
    On Error GoTo Failed
    modeBefore = Application.CutCopyMode
    If Not Application.CommandBars.GetEnabledMso("PasteValues") Then
        Probe_NativePasteValues = "disabled|" & modeBefore
        Exit Function
    End If
    Application.CommandBars.ExecuteMso "PasteValues"
    Probe_NativePasteValues = "ok|" & modeBefore & "|" & CLng(Application.CutCopyMode)
    Exit Function
Failed:
    Probe_NativePasteValues = "error|" & Err.Number & "|" & Err.Description
End Function

' Whether Excel's Undo command is available, and the Undo list as Excel shows it (newest first),
' when the classic Undo control still exposes it. "ok|<enabled>|<entryCount>|<entry>;<entry>..."
Public Function Probe_UndoState() As String
    Dim enabled As Boolean, ctl As Object, i As Long, entries As String, entryCount As Long
    On Error Resume Next
    enabled = Application.CommandBars.GetEnabledMso("Undo")
    Set ctl = Application.CommandBars.FindControl(Id:=128)
    If Not ctl Is Nothing Then
        entryCount = -1
        entryCount = ctl.ListCount
        For i = 1 To entryCount
            entries = entries & ctl.List(i) & ";"
        Next i
    Else
        entryCount = -2
    End If
    On Error GoTo 0
    Probe_UndoState = "ok|" & enabled & "|" & entryCount & "|" & entries
End Function

' Excel's own Undo, as the Undo button (one step).
Public Function Probe_NativeUndo() As String
    On Error GoTo Failed
    If Not Application.CommandBars.GetEnabledMso("Undo") Then Probe_NativeUndo = "disabled|": Exit Function
    Application.CommandBars.ExecuteMso "Undo"
    Probe_NativeUndo = "ok|"
    Exit Function
Failed:
    Probe_NativeUndo = "error|" & Err.Number & "|" & Err.Description
End Function

' A user-interface action that adds an entry to the Undo history (the "earlier history").
Public Function Probe_NativeBold() As String
    On Error GoTo Failed
    Application.CommandBars.ExecuteMso "Bold"
    Probe_NativeBold = "ok|"
    Exit Function
Failed:
    Probe_NativeBold = "error|" & Err.Number & "|" & Err.Description
End Function

' CONTROL 1: an ordinary object-model write (expected to clear Excel's Undo history).
Public Function Probe_ControlObjectModelWrite() As String
    On Error GoTo Failed
    Selection.Cells(1, 1).Value = "written by VBA"
    Probe_ControlObjectModelWrite = "ok|"
    Exit Function
Failed:
    Probe_ControlObjectModelWrite = "error|" & Err.Number & "|" & Err.Description
End Function

' CONTROL 2: the object-model paste the legacy macro uses (expected to clear the Undo history).
Public Function Probe_ControlPasteSpecial() As String
    Dim modeBefore As Long
    On Error GoTo Failed
    modeBefore = Application.CutCopyMode
    Selection.PasteSpecial Paste:=xlPasteValues
    Probe_ControlPasteSpecial = "ok|" & modeBefore & "|" & CLng(Application.CutCopyMode)
    Exit Function
Failed:
    Probe_ControlPasteSpecial = "error|" & Err.Number & "|" & Err.Description
End Function
