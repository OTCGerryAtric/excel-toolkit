VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TkFormatSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Excel Toolkit (M3a): the number formats of a selection before a format command, for undo.
'
' Captures only the property the command changes (NumberFormat), exactly. A uniformly formatted
' area is one read; a mixed area is split recursively (columns first, then rows) until every block
' is uniform, so a whole column with a few exceptions stays cheap. Too many distinct blocks is
' refused rather than risking an incomplete undo. Undo-record interface: Description,
' CheckIdentity, FindConflict, Restore (see TkUndo).
Option Explicit

Private Const MAX_BLOCKS As Long = 50000

Private wb As Workbook
Private ws As Worksheet
Private wbFullName As String
Private wsName As String
Private targetAddress As String
Private commandLabel As String
Private capturedAt As Date
Private target As Range
Private appliedFormat As String
Private blocks As Collection          ' Range per uniformly formatted block
Private blockFormats As Collection    ' its NumberFormat before the command

' ---------- Capture (no writes) ----------

' Returns "" or the reason the command is refused.
Public Function Capture(ByVal rng As Range, ByVal newFormat As String, ByVal label As String) As String
    Dim area As Range
    Set target = rng
    Set ws = rng.Worksheet
    Set wb = ws.Parent
    wbFullName = wb.FullName
    wsName = ws.Name
    targetAddress = rng.Address(False, False)
    If Len(targetAddress) > 60 Then targetAddress = Left$(targetAddress, 57) & "..."
    commandLabel = label
    capturedAt = Now
    appliedFormat = newFormat
    Set blocks = New Collection
    Set blockFormats = New Collection
    For Each area In rng.Areas
        If Not AddBlocks(area) Then
            Capture = "The selection has too many different number formats (more than " & Format(MAX_BLOCKS, "#,##0") & _
                " separate blocks) to undo safely. Select a smaller range; nothing was changed."
            Exit Function
        End If
    Next area
End Function

Private Function AddBlocks(ByVal rng As Range) As Boolean
    Dim fmt As Variant, rows As Long, cols As Long, half As Long, ok As Boolean
    fmt = rng.NumberFormat
    If Not IsNull(fmt) Then
        If blocks.Count >= MAX_BLOCKS Then Exit Function
        blocks.Add rng
        blockFormats.Add CStr(fmt)
        AddBlocks = True
        Exit Function
    End If
    ' The second half starts inside the range (Rows/Columns(half + 1)), never Offset: offsetting a
    ' whole column by half its height runs past the last row of the sheet (error 1004).
    rows = rng.Rows.Count
    cols = rng.Columns.Count
    If cols > 1 Then
        half = cols \ 2
        ok = AddBlocks(rng.Resize(, half))
        If ok Then ok = AddBlocks(rng.Columns(half + 1).Resize(, cols - half))
    Else
        half = rows \ 2
        ok = AddBlocks(rng.Resize(half))
        If ok Then ok = AddBlocks(rng.Rows(half + 1).Resize(rows - half))
    End If
    AddBlocks = ok
End Function

Public Property Get BlockCount() As Long
    BlockCount = blocks.Count
End Property

Public Function Description() As String
    Description = commandLabel & " on '" & wsName & "'!" & targetAddress & " (" & Format(capturedAt, "hh:nn") & ")"
End Function

' ---------- Apply and verify ----------

' Raises on failure; the caller rolls back.
Public Sub Apply()
    Dim area As Range, written As Long, limit As Long
    limit = Tk_TestFailAfterWrites()
    For Each area In target.Areas
        If limit > 0 And written >= limit Then Err.Raise vbObjectError + 516, "TkFormatSnapshot", "Simulated failure (test hook)"
        area.NumberFormat = appliedFormat
        written = written + 1
    Next area
End Sub

Public Function VerifyApplied() As String
    Dim area As Range, fmt As Variant
    For Each area In target.Areas
        fmt = area.NumberFormat
        If IsNull(fmt) Then
            VerifyApplied = "The number format of " & area.Address(False, False) & " was not applied to every cell."
            Exit Function
        ElseIf CStr(fmt) <> appliedFormat Then
            VerifyApplied = "The number format of " & area.Address(False, False) & " was not applied."
            Exit Function
        End If
    Next area
End Function

' ---------- Undo ----------

Public Function CheckIdentity() As String
    CheckIdentity = Tk_CheckTargetIdentity(wb, ws, wbFullName, wsName)
End Function

' "" when every area still has the format the command applied, else what changed.
Public Function FindConflict() As String
    Dim area As Range, fmt As Variant
    If ws.ProtectContents Then
        If Not ws.Protection.AllowFormattingCells Then
            FindConflict = "The sheet is now protected and does not allow formatting. Unprotect it first."
            Exit Function
        End If
    End If
    For Each area In target.Areas
        fmt = area.NumberFormat
        If IsNull(fmt) Then
            FindConflict = "Number formats in " & area.Address(False, False) & " on '" & wsName & "' have changed since the command."
            Exit Function
        ElseIf CStr(fmt) <> appliedFormat Then
            FindConflict = "Number formats in " & area.Address(False, False) & " on '" & wsName & "' have changed since the command."
            Exit Function
        End If
    Next area
End Function

' Puts every block's original format back. Never raises; "" on verified success.
Public Function Restore(ByVal useHook As Boolean) As String
    Dim i As Long, problems As String, limit As Long, done As Long
    If useHook Then limit = Tk_TestFailAfterRestores()
    For i = 1 To blocks.Count
        On Error Resume Next
        Err.Clear
        If limit > 0 And done >= limit Then Err.Raise vbObjectError + 517, "TkFormatSnapshot", "Simulated restore failure (test hook)"
        If Err.Number = 0 Then blocks(i).NumberFormat = blockFormats(i)
        If Err.Number <> 0 Then
            problems = problems & " " & blocks(i).Address(False, False)
        Else
            done = done + 1
        End If
        On Error GoTo 0
    Next i
    If Len(problems) > 0 Then Restore = "Number formats not restored:" & problems & ".": Exit Function
    Restore = VerifyRestored()
End Function

Private Function VerifyRestored() As String
    Dim i As Long, fmt As Variant, problems As String
    On Error GoTo Unreadable
    For i = 1 To blocks.Count
        fmt = blocks(i).NumberFormat
        If IsNull(fmt) Then
            problems = problems & " " & blocks(i).Address(False, False)
        ElseIf CStr(fmt) <> blockFormats(i) Then
            problems = problems & " " & blocks(i).Address(False, False)
        End If
    Next i
    If Len(problems) > 0 Then VerifyRestored = "Number format not restored exactly:" & problems & "."
    Exit Function
Unreadable:
    VerifyRestored = "The restored cells could not be read back to verify them: " & Err.Description
End Function
