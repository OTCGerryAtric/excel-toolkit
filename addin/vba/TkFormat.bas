Attribute VB_Name = "TkFormat"
' Excel Toolkit (M3a): number-format commands ported from PERSONAL.XLSB (Ctrl+Shift+I / O / P / M).
'
' Same format strings and cycle order as the legacy macros: if the selection already has a format
' in the cycle (other than the last), move to the next one; anything else - the last format, any
' other format, or mixed formats - goes to the first. Unlike the legacy macros the change can be
' undone (Ctrl+Z straight away, or Undo Toolkit Action), and only the number format is touched.
Option Explicit

Private lastTimings As String

' Stages decide how a failure is reported (as in TkConvert): before any write it is a refusal;
' during the writes the formats are put back; after verified writes it is never "nothing changed".
Private Const STAGE_CHECKING As Long = 0
Private Const STAGE_WRITING As Long = 1
Private Const STAGE_WRITTEN As Long = 2

' ---------- The cycles (legacy formats; {GBP} is the pound sign, kept out of the ASCII source) ----------

Public Function Tk_FormatCycle(ByVal cycleName As String) As Variant
    Select Case cycleName
    Case "number"
        Tk_FormatCycle = Array("#,##0_-;[Red](#,##0)_-;-_-", "#,##0.0_-;[Red](#,##0.0)_-;-_-", _
                               "#,##0.00_-;[Red](#,##0.00)_-;-_-")
    Case "currency"
        Tk_FormatCycle = WithPound(Array("{GBP}#,##0_-;[Red]({GBP}#,##0)_-;-_-", "{GBP}#,##0.0_-;[Red]({GBP}#,##0.0)_-;-_-", _
                                         "{GBP}#,##0.00_-;[Red]({GBP}#,##0.00)_-;-_-"))
    Case "percent"
        Tk_FormatCycle = Array("0%_-;[Red](0%)_-;-_-", "0.0%_-;[Red](0.0%)_-;-_-", "0.00%_-;[Red](0.00%)_-;-_-", _
                               "0.000%_-;[Red](0.000%)_-;-_-")
    Case "date"
        Tk_FormatCycle = Array("dd-mmm-yy", "mmm-yy")
    Case Else
        Tk_FormatCycle = Empty
    End Select
End Function

Private Function WithPound(ByVal formats As Variant) As Variant
    Dim i As Long
    For i = LBound(formats) To UBound(formats)
        formats(i) = Replace(formats(i), "{GBP}", ChrW(163))
    Next i
    WithPound = formats
End Function

Public Function Tk_FormatLabel(ByVal cycleName As String) As String
    Select Case cycleName
    Case "number": Tk_FormatLabel = "Number Format"
    Case "currency": Tk_FormatLabel = "Currency Format"
    Case "percent": Tk_FormatLabel = "Percentage Format"
    Case "date": Tk_FormatLabel = "Date Format"
    End Select
End Function

' The legacy rule. `current` is the selection's NumberFormat (Null when mixed).
Public Function Tk_NextFormat(ByVal current As Variant, ByVal formats As Variant) As String
    Dim i As Long
    If Not IsNull(current) Then
        For i = LBound(formats) To UBound(formats) - 1
            If CStr(current) = formats(i) Then
                Tk_NextFormat = formats(i + 1)
                Exit Function
            End If
        Next i
    End If
    Tk_NextFormat = formats(LBound(formats))
End Function

' ---------- Commands ----------

' Ribbon and keyboard entry point: works on the selection; does nothing if it is not a range.
Public Sub Tk_RunFormat(ByVal cycleName As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_FormatRange(Selection, cycleName)
End Sub

' Returns "ok: / refused: / warning: / rolled back: / incomplete: <text>".
Public Function Tk_FormatRange(ByVal target As Range, ByVal cycleName As String) As String
    Dim formats As Variant, newFormat As String, snap As TkFormatSnapshot, reason As String
    Dim rollbackProblem As String, stage As Long, errorText As String
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double

    On Error GoTo Failed
    t0 = Tk_Seconds()
    formats = Tk_FormatCycle(cycleName)
    If Not IsArray(formats) Then Tk_FormatRange = "refused: Unknown format command '" & cycleName & "'.": Exit Function
    If target.Worksheet.ProtectContents Then
        If Not target.Worksheet.Protection.AllowFormattingCells Then
            Tk_FormatRange = "refused: The sheet is protected and does not allow formatting. Unprotect it first; nothing was changed."
            Exit Function
        End If
    End If

    newFormat = Tk_NextFormat(target.NumberFormat, formats)
    Set snap = New TkFormatSnapshot
    reason = snap.Capture(target, newFormat, Tk_FormatLabel(cycleName))
    If Len(reason) > 0 Then Tk_FormatRange = "refused: " & reason: Exit Function
    t1 = Tk_Seconds()

    stage = STAGE_WRITING
    snap.Apply
    reason = snap.VerifyApplied()
    If Len(reason) > 0 Then GoTo RollBack
    stage = STAGE_WRITTEN
    t2 = Tk_Seconds()

    Tk_SetRecord snap, Tk_FormatLabel(cycleName)        ' replaces any earlier record; registers Ctrl+Z
    t3 = Tk_Seconds()
    lastTimings = "blocks=" & snap.BlockCount & ";capture=" & Ms(t1 - t0) & ";apply+verify=" & Ms(t2 - t1) & _
                  ";register=" & Ms(t3 - t2) & ";total=" & Ms(t3 - t0)
    Tk_FormatRange = "ok: " & Tk_FormatLabel(cycleName) & " set to " & newFormat & "."
    Exit Function

Failed:
    errorText = Err.Description
    Select Case stage
    Case STAGE_CHECKING
        Tk_FormatRange = "refused: Excel reported an error before anything was changed: " & errorText
        Exit Function
    Case STAGE_WRITING
        reason = "Excel reported an error while formatting: " & errorText
        Resume RollBack
    Case Else
        If Tk_RecordIs(snap) Then
            Tk_FormatRange = "warning: The format was applied and checked, but Ctrl+Z could not be set up (" & _
                errorText & "). Use Undo Toolkit Action to reverse it."
            Exit Function
        End If
        reason = "The format was applied, but the undo record could not be stored (" & errorText & ")."
        Resume RollBack
    End Select
RollBack:
    On Error GoTo 0
    rollbackProblem = snap.Restore(False)
    If Len(rollbackProblem) = 0 Then
        Tk_FormatRange = "rolled back: " & reason
    Else
        Tk_FormatRange = "incomplete: " & reason & " Restoring the original formats also failed: " & rollbackProblem
    End If
End Function

Public Function Tk_LastFormatTimings() As String
    Tk_LastFormatTimings = lastTimings
End Function

Private Function Ms(ByVal seconds As Double) As String
    Ms = Format(seconds * 1000, "0.0")
End Function
