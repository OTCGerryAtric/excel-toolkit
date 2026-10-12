Attribute VB_Name = "TkStyle"
' Excel Toolkit (M3b): fill, outline and size commands ported from PERSONAL.XLSB (legacy keys
' Ctrl+Shift+Q / E / W / Z fills, G outline, R row height, C column width, K best fit).
'
' Same colours, borders, sizes and decisions as the legacy macros, with one documented fix: a
' fill toggled off becomes No Fill instead of white (white hid the gridlines). Unlike the legacy
' macros every change can be undone (Ctrl+Z straight away, or Undo Toolkit Action); only the
' properties the command changes are captured (TkStyleSnapshot). No keyboard shortcuts yet: keys
' move one command at a time after the command passes its live test.
Option Explicit

Private lastTimings As String

' Stages decide how a failure is reported (as in TkConvert and TkFormat).
Private Const STAGE_CHECKING As Long = 0
Private Const STAGE_WRITING As Long = 1
Private Const STAGE_WRITTEN As Long = 2

Private Const NO_FILL As Long = -1
' Legacy Grey_Ouline: thin continuous lines, theme colour 1 (xlThemeColorDark1) darkened 25%.
Private Const OUTLINE_THEME As Long = 1
Private Const OUTLINE_TINT As Double = -0.249946592608417
Private Const MAX_COLUMN_WIDTH As Double = 255
' Read-back tolerances: Excel rounds heights and widths to whole screen pixels.
Private Const ROW_TOLERANCE As Double = 0.75
Private Const COLUMN_TOLERANCE As Double = 0.25

' ---------- Legacy data and rules ----------

Public Function Tk_FillColour(ByVal fillName As String) As Long
    Select Case fillName
    Case "yellow": Tk_FillColour = RGB(255, 255, 210)
    Case "green": Tk_FillColour = RGB(145, 245, 190)
    Case "lightblue": Tk_FillColour = RGB(221, 235, 247)
    Case "grey": Tk_FillColour = RGB(242, 242, 242)
    Case Else: Tk_FillColour = -1
    End Select
End Function

Public Function Tk_StyleLabel(ByVal command As String) As String
    Select Case command
    Case "yellow": Tk_StyleLabel = "Yellow Fill"
    Case "green": Tk_StyleLabel = "Green Fill"
    Case "lightblue": Tk_StyleLabel = "Light Blue Fill"
    Case "grey": Tk_StyleLabel = "Grey Fill"
    Case "outline": Tk_StyleLabel = "Grey Outline"
    Case "rowheight": Tk_StyleLabel = "Row Height"
    Case "colwidth": Tk_StyleLabel = "Column Width"
    Case "bestfit": Tk_StyleLabel = "Best Fit Width"
    End Select
End Function

' The legacy rules. `current` is the selection's property: Null when mixed, which never matches.
Public Function Tk_NextRowHeight(ByVal current As Variant) As Double
    If Not IsNull(current) Then
        If current = 15 Then Tk_NextRowHeight = 4.5: Exit Function
    End If
    Tk_NextRowHeight = 15
End Function

Public Function Tk_NextColumnWidth(ByVal current As Variant) As Double
    If Not IsNull(current) Then
        If current = 2 Then Tk_NextColumnWidth = 8.43: Exit Function
        If current = 8.43 Then Tk_NextColumnWidth = 12: Exit Function
    End If
    Tk_NextColumnWidth = 2
End Function

' Legacy Best_Fit_Column_Width: CColWidth As Integer = Round(width + 5) / 5, width = CColWidth * 5.
' The assignment to Integer rounds a second time (banker's rounding), exactly as the macro does.
Public Function Tk_BestFitWidth(ByVal activeWidth As Double) As Double
    Dim steps As Integer
    steps = Round(activeWidth + 5) / 5
    Tk_BestFitWidth = steps * 5
End Function

' Every edge of every cell in a range: its outside edges, plus the inside edges it has.
Public Function Tk_CellEdges(ByVal rng As Range) As Variant
    If rng.Columns.Count > 1 And rng.Rows.Count > 1 Then
        Tk_CellEdges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
    ElseIf rng.Columns.Count > 1 Then
        Tk_CellEdges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical)
    ElseIf rng.Rows.Count > 1 Then
        Tk_CellEdges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal)
    Else
        Tk_CellEdges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    End If
End Function

' "" when a protected sheet still allows this kind of change, else why not.
Public Function Tk_StyleProtectionProblem(ByVal ws As Worksheet, ByVal kind As String) As String
    If Not ws.ProtectContents Then Exit Function
    Select Case kind
    Case "rowheight"
        If Not ws.Protection.AllowFormattingRows Then _
            Tk_StyleProtectionProblem = "The sheet is protected and does not allow changing row heights. Unprotect it first."
    Case "colwidth"
        If Not ws.Protection.AllowFormattingColumns Then _
            Tk_StyleProtectionProblem = "The sheet is protected and does not allow changing column widths. Unprotect it first."
    Case Else
        If Not ws.Protection.AllowFormattingCells Then _
            Tk_StyleProtectionProblem = "The sheet is protected and does not allow formatting. Unprotect it first."
    End Select
End Function

Private Function KindOf(ByVal command As String) As String
    Select Case command
    Case "outline": KindOf = "borders"
    Case "rowheight": KindOf = "rowheight"
    Case "colwidth", "bestfit": KindOf = "colwidth"
    Case Else: KindOf = "fill"
    End Select
End Function

' ---------- Commands ----------

' Ribbon entry points: work on the selection; do nothing if it is not a range.
Public Sub Tk_RunFill(ByVal fillName As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_FillRange(Selection, fillName)
End Sub

Public Sub Tk_RunOutline()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_OutlineRange(Selection)
End Sub

Public Sub Tk_RunRowHeight()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_RowHeightRange(Selection)
End Sub

Public Sub Tk_RunColumnWidth()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_ColumnWidthRange(Selection)
End Sub

Public Sub Tk_RunBestFit()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_BestFitRange(Selection, ActiveCell)
End Sub

' Each returns "ok: / refused: / warning: / rolled back: / incomplete: <text>".
Public Function Tk_FillRange(ByVal target As Range, ByVal fillName As String) As String
    Tk_FillRange = RunStyle(target, fillName, Nothing)
End Function

Public Function Tk_OutlineRange(ByVal target As Range) As String
    Tk_OutlineRange = RunStyle(target, "outline", Nothing)
End Function

Public Function Tk_RowHeightRange(ByVal target As Range) As String
    Tk_RowHeightRange = RunStyle(target, "rowheight", Nothing)
End Function

Public Function Tk_ColumnWidthRange(ByVal target As Range) As String
    Tk_ColumnWidthRange = RunStyle(target, "colwidth", Nothing)
End Function

' `reference` is the cell whose column width the new width is based on (the active cell).
Public Function Tk_BestFitRange(ByVal target As Range, ByVal reference As Range) As String
    Tk_BestFitRange = RunStyle(target, "bestfit", reference)
End Function

Private Function RunStyle(ByVal target As Range, ByVal command As String, ByVal reference As Range) As String
    Dim snap As TkStyleSnapshot, kind As String, newValue As Variant, okText As String, reason As String
    Dim rollbackProblem As String, stage As Long, errorText As String, saved As Variant
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double

    On Error GoTo Failed
    t0 = Tk_Seconds()
    kind = KindOf(command)
    reason = Tk_StyleProtectionProblem(target.Worksheet, kind)
    If Len(reason) = 0 Then reason = Decide(target, command, reference, newValue, okText)
    If Len(reason) > 0 Then RunStyle = "refused: " & reason & " Nothing was changed.": Exit Function

    Set snap = New TkStyleSnapshot
    reason = snap.Capture(target, kind, Tk_StyleLabel(command))
    If Len(reason) > 0 Then RunStyle = "refused: " & reason: Exit Function
    t1 = Tk_Seconds()

    saved = Tk_QuietExcel()
    stage = STAGE_WRITING
    ApplyChange target, command, newValue
    reason = VerifyChange(target, command, newValue)
    If Len(reason) = 0 Then reason = snap.RecordApplied()
    If Len(reason) > 0 Then GoTo RollBack
    stage = STAGE_WRITTEN
    Tk_RestoreExcel saved
    t2 = Tk_Seconds()

    Tk_SetRecord snap, Tk_StyleLabel(command)        ' replaces any earlier record; registers Ctrl+Z
    t3 = Tk_Seconds()
    lastTimings = "blocks=" & snap.BlockCount & ";capture=" & Ms(t1 - t0) & ";apply+verify=" & Ms(t2 - t1) & _
                  ";register=" & Ms(t3 - t2) & ";total=" & Ms(t3 - t0)
    RunStyle = "ok: " & okText
    Exit Function

Failed:
    errorText = Err.Description
    Select Case stage
    Case STAGE_CHECKING
        RunStyle = "refused: Excel reported an error before anything was changed: " & errorText
        Exit Function
    Case STAGE_WRITING
        reason = "Excel reported an error while formatting: " & errorText
        Resume RollBack
    Case Else
        If Tk_RecordIs(snap) Then
            RunStyle = "warning: The change was applied and checked, but Ctrl+Z could not be set up (" & _
                errorText & "). Use Undo Toolkit Action to reverse it."
            Exit Function
        End If
        reason = "The change was applied, but the undo record could not be stored (" & errorText & ")."
        Resume RollBack
    End Select
RollBack:
    On Error GoTo 0
    rollbackProblem = snap.Restore(False)
    If IsArray(saved) Then Tk_RestoreExcel saved
    If Len(rollbackProblem) = 0 Then
        RunStyle = "rolled back: " & reason
    Else
        RunStyle = "incomplete: " & reason & " Putting the original formatting back also failed: " & rollbackProblem
    End If
End Function

' The legacy decision for the command. Returns "" (newValue and okText set) or why it is refused.
Private Function Decide(ByVal target As Range, ByVal command As String, ByVal reference As Range, _
                        ByRef newValue As Variant, ByRef okText As String) As String
    Dim current As Variant, colour As Long
    Select Case command
    Case "outline"
        ' Legacy: add the outline only when the top edge has no line; otherwise (a line, or a
        ' mixed top edge) remove all six edges.
        current = target.Borders(xlEdgeTop).LineStyle
        newValue = False
        If Not IsNull(current) Then newValue = (current = xlNone)
        If newValue Then okText = "Grey outline added." Else okText = "Outline and inside borders removed."
    Case "rowheight"
        newValue = Tk_NextRowHeight(target.RowHeight)
        okText = "Row height set to " & newValue & "."
    Case "colwidth"
        newValue = Tk_NextColumnWidth(target.ColumnWidth)
        okText = "Column width set to " & newValue & "."
    Case "bestfit"
        If reference Is Nothing Then Decide = "There is no active cell to take the column width from.": Exit Function
        newValue = Tk_BestFitWidth(reference.ColumnWidth)
        If newValue > MAX_COLUMN_WIDTH Then
            Decide = "The active cell's column is " & reference.ColumnWidth & " wide; best fit would set " & newValue & _
                ", more than Excel's maximum column width (255)."
            Exit Function
        End If
        okText = "Column width set to " & newValue & " (from the active cell's width, " & reference.ColumnWidth & ")."
    Case Else
        colour = Tk_FillColour(command)
        If colour < 0 Then Decide = "Unknown command '" & command & "'.": Exit Function
        current = target.Interior.Color
        newValue = colour
        okText = Tk_StyleLabel(command) & " applied."
        If Not IsNull(current) Then
            If current = colour Then
                newValue = NO_FILL                       ' fix: legacy set white here
                okText = Tk_StyleLabel(command) & " removed (No Fill)."
            End If
        End If
    End Select
End Function

' Raises on failure; the caller rolls back.
Private Sub ApplyChange(ByVal target As Range, ByVal command As String, ByVal newValue As Variant)
    Dim area As Range, edge As Variant, written As Long, limit As Long
    limit = Tk_TestFailAfterWrites()
    If command = "outline" Then
        ' Exactly the legacy sequence, on the whole selection, one edge at a time.
        For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
            If limit > 0 And written >= limit Then Err.Raise vbObjectError + 516, "TkStyle", "Simulated failure (test hook)"
            If newValue Then
                With target.Borders(edge)
                    .LineStyle = xlContinuous
                    .ThemeColor = OUTLINE_THEME
                    .TintAndShade = OUTLINE_TINT
                    .Weight = xlThin
                End With
            Else
                target.Borders(edge).LineStyle = xlNone
            End If
            written = written + 1
        Next edge
        Exit Sub
    End If
    For Each area In target.Areas
        If limit > 0 And written >= limit Then Err.Raise vbObjectError + 516, "TkStyle", "Simulated failure (test hook)"
        Select Case command
        Case "rowheight": area.RowHeight = newValue
        Case "colwidth", "bestfit": area.ColumnWidth = newValue
        Case Else
            If newValue = NO_FILL Then area.Interior.Pattern = xlNone Else area.Interior.Color = newValue
        End Select
        written = written + 1
    Next area
End Sub

Private Function VerifyChange(ByVal target As Range, ByVal command As String, ByVal newValue As Variant) As String
    Dim area As Range, edge As Variant, fault As Boolean
    For Each area In target.Areas
        Select Case command
        Case "outline"
            For Each edge In Tk_CellEdges(area)
                If newValue Then
                    fault = Differs(area.Borders(edge).LineStyle, xlContinuous) Or Differs(area.Borders(edge).Weight, xlThin)
                Else
                    fault = Differs(area.Borders(edge).LineStyle, xlNone)
                End If
                If fault Then Exit For
            Next edge
        Case "rowheight"
            fault = OffBy(area.RowHeight, newValue, ROW_TOLERANCE)
        Case "colwidth", "bestfit"
            fault = OffBy(area.ColumnWidth, newValue, COLUMN_TOLERANCE)
        Case Else
            If newValue = NO_FILL Then
                fault = Differs(area.Interior.Pattern, xlNone)
            Else
                fault = Differs(area.Interior.Color, newValue)
            End If
        End Select
        If fault Then
            VerifyChange = "The change to " & area.Address(False, False) & " could not be confirmed."
            Exit Function
        End If
    Next area
End Function

Private Function Differs(ByVal got As Variant, ByVal expected As Variant) As Boolean
    If IsNull(got) Then Differs = True Else Differs = (got <> expected)
End Function

Private Function OffBy(ByVal got As Variant, ByVal expected As Double, ByVal tolerance As Double) As Boolean
    If IsNull(got) Then OffBy = True Else OffBy = (Abs(got - expected) > tolerance)
End Function

Public Function Tk_LastStyleTimings() As String
    Tk_LastStyleTimings = lastTimings
End Function

Private Function Ms(ByVal seconds As Double) As String
    Ms = Format(seconds * 1000, "0.0")
End Function
