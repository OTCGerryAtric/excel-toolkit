VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TkStyleSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Excel Toolkit (M3b): the fill, borders, row heights or column widths a command is about to
' change, for undo.
'
' Captures only the properties the command changes, exactly (M3 plan, "Undo for formatting"):
'   fill      - pattern, colour (RGB, or theme colour + tint) and, for patterned fills, the
'               pattern colour. Gradient fills are refused.
'   borders   - each edge's line style, weight and colour (automatic, RGB, or theme colour +
'               tint) for every cell of the selection, and the facing edge of the cells around
'               it, which share the selection's outside edges. Those are restored last.
'   rowheight - the height of every affected ROW (whole rows, not just the selected cells).
'               Hidden rows are refused: their height before hiding cannot be read.
'   colwidth  - the width of every affected COLUMN; hidden columns are refused likewise.
' A uniform range is one read; a mixed one is split recursively (columns first, then rows) into
' uniform blocks, as in TkFormatSnapshot; more than MAX_BLOCKS blocks is refused. Undo-record
' interface: Description, CheckIdentity, FindConflict, Restore (see TkUndo).
'
' Theme colours: Excel's ThemeColor raises an error for a colour that is not a theme colour. For
' one cell that means "RGB"; for several cells it could also hide a mixture of RGB and theme
' colours of the same shade, so such a range is split until each block is one cell (exact, but
' slower for large RGB-coloured areas - measured in the live run).
Option Explicit

Private Const MAX_BLOCKS As Long = 50000
Private Const NO_THEME As String = "rgb"

Private wb As Workbook
Private ws As Worksheet
Private wbFullName As String
Private wsName As String
Private targetAddress As String
Private commandLabel As String
Private capturedAt As Date
Private kind As String                ' "fill", "borders", "rowheight" or "colwidth"
Private blocks As Collection          ' Range per uniform block: the selection first, then the cells around it
Private blockEdges As Collection      ' borders: the edge indexes each block's state covers (else Empty)
Private before As Collection          ' each block's state before the command
Private after As Collection           ' and after it; the conflict check compares against this
Private refusal As String

' ---------- Capture (no writes) ----------

' Returns "" or the reason the command is refused.
Public Function Capture(ByVal rng As Range, ByVal kindName As String, ByVal label As String) As String
    Dim area As Range, parts As Collection, part As Variant
    Set ws = rng.Worksheet
    Set wb = ws.Parent
    wbFullName = wb.FullName
    wsName = ws.Name
    targetAddress = rng.Address(False, False)
    If Len(targetAddress) > 60 Then targetAddress = Left$(targetAddress, 57) & "..."
    commandLabel = label
    capturedAt = Now
    kind = kindName
    Set blocks = New Collection
    Set blockEdges = New Collection
    Set before = New Collection
    Set after = New Collection

    ' Array(range, edges): edges is "all" (every edge of every cell), an Array of edge indexes,
    ' or Empty for kinds that are not borders.
    Set parts = New Collection
    For Each area In rng.Areas
        Select Case kind
        Case "rowheight"            ' column A of the affected rows: its RowHeight is the rows' height
            parts.Add Array(ws.Range(ws.Cells(area.Row, 1), ws.Cells(area.Row + area.Rows.Count - 1, 1)), Empty)
        Case "colwidth"             ' row 1 of the affected columns
            parts.Add Array(ws.Range(ws.Cells(1, area.Column), ws.Cells(1, area.Column + area.Columns.Count - 1)), Empty)
        Case "borders"
            parts.Add Array(area, "all")
        Case Else
            parts.Add Array(area, Empty)
        End Select
    Next area
    If kind = "borders" Then
        For Each area In rng.Areas
            AddSurroundingEdges parts, area
        Next area
    End If

    For Each part In parts
        If Not AddBlocks(part(0), part(1)) Then
            If Len(refusal) = 0 Then
                refusal = "The selection has too many differently formatted blocks (more than " & _
                    Format(MAX_BLOCKS, "#,##0") & ") to undo safely. Select a smaller range; nothing was changed."
            End If
            Capture = refusal
            Exit Function
        End If
    Next part
End Function

' The cells around an area share its outside edges (the cell above shows the area's top edge as
' its bottom edge), so their facing edge is captured too.
Private Sub AddSurroundingEdges(ByVal parts As Collection, ByVal area As Range)
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    r1 = area.Row: r2 = r1 + area.Rows.Count - 1
    c1 = area.Column: c2 = c1 + area.Columns.Count - 1
    If r1 > 1 Then parts.Add Array(ws.Range(ws.Cells(r1 - 1, c1), ws.Cells(r1 - 1, c2)), Array(xlEdgeBottom))
    If r2 < ws.Rows.Count Then parts.Add Array(ws.Range(ws.Cells(r2 + 1, c1), ws.Cells(r2 + 1, c2)), Array(xlEdgeTop))
    If c1 > 1 Then parts.Add Array(ws.Range(ws.Cells(r1, c1 - 1), ws.Cells(r2, c1 - 1)), Array(xlEdgeRight))
    If c2 < ws.Columns.Count Then parts.Add Array(ws.Range(ws.Cells(r1, c2 + 1), ws.Cells(r2, c2 + 1)), Array(xlEdgeLeft))
End Sub

Private Function AddBlocks(ByVal rng As Range, ByVal edgeSpec As Variant) As Boolean
    Dim edges As Variant, state As Variant, rows As Long, cols As Long, half As Long, ok As Boolean
    edges = EdgesFor(rng, edgeSpec)
    state = ReadState(rng, edges, False)
    If IsArray(state) Then
        If blocks.Count >= MAX_BLOCKS Then Exit Function
        blocks.Add rng
        blockEdges.Add edges
        before.Add state
        AddBlocks = True
        Exit Function
    End If
    If Not IsNull(state) Then refusal = CStr(state): Exit Function         ' unsupported: refused
    If rng.CountLarge = 1 Then
        refusal = "The formatting of " & rng.Address(False, False) & " on '" & wsName & _
            "' could not be read exactly, so the change could not be undone; nothing was changed."
        Exit Function
    End If
    ' The second half starts inside the range, never Offset (see TkFormatSnapshot).
    rows = rng.Rows.Count
    cols = rng.Columns.Count
    If cols > 1 Then
        half = cols \ 2
        ok = AddBlocks(rng.Resize(, half), edgeSpec)
        If ok Then ok = AddBlocks(rng.Columns(half + 1).Resize(, cols - half), edgeSpec)
    Else
        half = rows \ 2
        ok = AddBlocks(rng.Resize(half), edgeSpec)
        If ok Then ok = AddBlocks(rng.Rows(half + 1).Resize(rows - half), edgeSpec)
    End If
    AddBlocks = ok
End Function

Private Function EdgesFor(ByVal rng As Range, ByVal edgeSpec As Variant) As Variant
    If IsArray(edgeSpec) Then
        EdgesFor = edgeSpec
    ElseIf IsEmpty(edgeSpec) Then
        EdgesFor = Empty
    Else
        EdgesFor = Tk_CellEdges(rng)
    End If
End Function

Public Property Get BlockCount() As Long
    BlockCount = blocks.Count
End Property

Public Function Description() As String
    Description = commandLabel & " on '" & wsName & "'!" & targetAddress & " (" & Format(capturedAt, "hh:nn") & ")"
End Function

' ---------- Reading and writing one block ----------

' An Array when the range is uniform; Null when it must be split; a String when it cannot be
' captured (the command is refused). `lenient` (checks after the command only) treats an
' unreadable theme colour as RGB instead of splitting.
Private Function ReadState(ByVal rng As Range, ByVal edges As Variant, ByVal lenient As Boolean) As Variant
    Dim cellCount As Double
    If lenient Then cellCount = 1 Else cellCount = rng.CountLarge
    Select Case kind
    Case "fill": ReadState = FillState(rng, cellCount)
    Case "borders": ReadState = BorderState(rng, edges, cellCount)
    Case "rowheight": ReadState = SizeState(rng, rng.RowHeight, "rows")
    Case "colwidth": ReadState = SizeState(rng, rng.ColumnWidth, "columns")
    Case Else: ReadState = "Unknown formatting kind '" & kind & "'; nothing was changed."
    End Select
End Function

Private Function FillState(ByVal rng As Range, ByVal cellCount As Double) As Variant
    Dim fillFormat As Interior, pattern As Variant, colour As Variant, theme As Variant, tint As Variant
    Dim patternIndex As Variant, patternColour As Variant, patternTheme As Variant, patternTint As Variant
    Set fillFormat = rng.Interior
    pattern = fillFormat.Pattern
    If IsNull(pattern) Then FillState = Null: Exit Function
    If pattern = xlNone Then FillState = Array(xlNone): Exit Function
    If pattern = xlPatternLinearGradient Or pattern = xlPatternRectangularGradient Then
        FillState = "The selection includes a gradient fill (" & rng.Address(False, False) & "), which cannot be " & _
            "restored exactly, so it is not supported. Nothing was changed."
        Exit Function
    End If
    colour = fillFormat.Color
    theme = ThemeOf(fillFormat, False, cellCount)
    tint = fillFormat.TintAndShade
    If IsNull(colour) Or IsNull(theme) Or IsNull(tint) Then FillState = Null: Exit Function
    If pattern = xlSolid Then FillState = Array(pattern, colour, theme, tint): Exit Function

    patternIndex = fillFormat.PatternColorIndex
    If IsNull(patternIndex) Then FillState = Null: Exit Function
    If patternIndex = xlAutomatic Then FillState = Array(pattern, colour, theme, tint, True): Exit Function
    patternColour = fillFormat.PatternColor
    patternTheme = ThemeOf(fillFormat, True, cellCount)
    patternTint = fillFormat.PatternTintAndShade
    If IsNull(patternColour) Or IsNull(patternTheme) Or IsNull(patternTint) Then FillState = Null: Exit Function
    FillState = Array(pattern, colour, theme, tint, False, patternColour, patternTheme, patternTint)
End Function

Private Function BorderState(ByVal rng As Range, ByVal edges As Variant, ByVal cellCount As Double) As Variant
    Dim states() As Variant, i As Long, one As Variant
    ReDim states(LBound(edges) To UBound(edges))
    For i = LBound(edges) To UBound(edges)
        one = EdgeState(rng.Borders(edges(i)), cellCount)
        If Not IsArray(one) Then BorderState = one: Exit Function
        states(i) = one
    Next i
    BorderState = states
End Function

Private Function EdgeState(ByVal edge As Border, ByVal cellCount As Double) As Variant
    Dim lineKind As Variant, weight As Variant, colourIndex As Variant, colour As Variant, theme As Variant, tint As Variant
    lineKind = edge.LineStyle
    If IsNull(lineKind) Then EdgeState = Null: Exit Function
    If lineKind = xlNone Then EdgeState = Array(xlNone): Exit Function
    weight = edge.Weight
    colourIndex = edge.ColorIndex
    If IsNull(weight) Or IsNull(colourIndex) Then EdgeState = Null: Exit Function
    If colourIndex = xlAutomatic Then EdgeState = Array(lineKind, weight, True): Exit Function
    colour = edge.Color
    theme = ThemeOf(edge, False, cellCount)
    tint = edge.TintAndShade
    If IsNull(colour) Or IsNull(theme) Or IsNull(tint) Then EdgeState = Null: Exit Function
    EdgeState = Array(lineKind, weight, False, colour, theme, tint)
End Function

Private Function SizeState(ByVal rng As Range, ByVal size As Variant, ByVal what As String) As Variant
    Dim where As String
    If IsNull(size) Then SizeState = Null: Exit Function
    If size = 0 Then
        If what = "rows" Then where = rng.EntireRow.Address(False, False) Else where = rng.EntireColumn.Address(False, False)
        SizeState = "The selection includes hidden " & what & " (" & where & "). Changing their size would unhide " & _
            "them, and the size they had before they were hidden cannot be read, so it could not be undone. " & _
            "Unhide them or leave them out of the selection; nothing was changed."
        Exit Function
    End If
    SizeState = Array(CDbl(size))
End Function

' The theme colour number (1-12), NO_THEME for an RGB colour, or Null when it cannot be told for
' this range (mixed, or an error on several cells - see the header).
Private Function ThemeOf(ByVal source As Object, ByVal ofPattern As Boolean, ByVal cellCount As Double) As Variant
    Dim theme As Variant
    On Error GoTo Unreadable
    If ofPattern Then theme = source.PatternThemeColor Else theme = source.ThemeColor
    On Error GoTo 0
    If IsNull(theme) Then ThemeOf = Null: Exit Function
    If IsNumeric(theme) Then
        If theme >= xlThemeColorDark1 And theme <= xlThemeColorFollowedHyperlink Then ThemeOf = CLng(theme): Exit Function
    End If
    ThemeOf = NO_THEME
    Exit Function
Unreadable:
    If cellCount = 1 Then ThemeOf = NO_THEME Else ThemeOf = Null
End Function

Private Function IsRgb(ByVal theme As Variant) As Boolean
    IsRgb = (VarType(theme) = vbString)
End Function

Private Sub WriteState(ByVal rng As Range, ByVal edges As Variant, ByVal state As Variant)
    Dim i As Long
    Select Case kind
    Case "fill": WriteFill rng.Interior, state
    Case "borders"
        For i = LBound(edges) To UBound(edges)
            WriteEdge rng.Borders(edges(i)), state(i)
        Next i
    Case "rowheight": rng.RowHeight = state(0)
    Case "colwidth": rng.ColumnWidth = state(0)
    End Select
End Sub

Private Sub WriteFill(ByVal fillFormat As Interior, ByVal state As Variant)
    If state(0) = xlNone Then fillFormat.Pattern = xlNone: Exit Sub
    fillFormat.Pattern = state(0)
    WriteColour fillFormat, False, state(1), state(2), state(3)
    If UBound(state) < 4 Then Exit Sub                       ' solid fill: no pattern colour
    If state(4) Then
        fillFormat.PatternColorIndex = xlAutomatic
    Else
        WriteColour fillFormat, True, state(5), state(6), state(7)
    End If
End Sub

Private Sub WriteEdge(ByVal edge As Border, ByVal state As Variant)
    If state(0) = xlNone Then edge.LineStyle = xlNone: Exit Sub
    edge.LineStyle = state(0)
    edge.Weight = state(1)
    If state(2) Then
        edge.ColorIndex = xlAutomatic
    Else
        WriteColour edge, False, state(3), state(4), state(5)
    End If
End Sub

' An RGB or theme colour and its tint, on an Interior (main or pattern colour) or a Border.
Private Sub WriteColour(ByVal target As Object, ByVal ofPattern As Boolean, ByVal colour As Variant, _
                        ByVal theme As Variant, ByVal tint As Variant)
    If ofPattern Then
        If IsRgb(theme) Then target.PatternColor = colour Else target.PatternThemeColor = theme
        target.PatternTintAndShade = tint
    Else
        If IsRgb(theme) Then target.Color = colour Else target.ThemeColor = theme
        target.TintAndShade = tint
    End If
End Sub

Private Function SameState(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim i As Long
    If IsArray(a) <> IsArray(b) Then Exit Function
    If IsArray(a) Then
        If LBound(a) <> LBound(b) Or UBound(a) <> UBound(b) Then Exit Function
        For i = LBound(a) To UBound(a)
            If Not SameState(a(i), b(i)) Then Exit Function
        Next i
        SameState = True
        Exit Function
    End If
    If IsNull(a) Or IsNull(b) Then Exit Function
    If VarType(a) = vbString Or VarType(b) = vbString Then
        If VarType(a) = VarType(b) Then SameState = (a = b)
        Exit Function
    End If
    SameState = (a = b)
End Function

' ---------- After the command ----------

' Reads every block back after the command (the conflict check compares against it). "" or why
' the result cannot be checked later - the caller then rolls back.
Public Function RecordApplied() As String
    Dim i As Long, state As Variant
    Set after = New Collection
    For i = 1 To blocks.Count
        state = ReadState(blocks(i), blockEdges(i), True)
        If Not IsArray(state) Then
            RecordApplied = "The new formatting of " & blocks(i).Address(False, False) & " could not be read back to check it."
            Exit Function
        End If
        after.Add state
    Next i
End Function

' ---------- Undo ----------

Public Function CheckIdentity() As String
    CheckIdentity = Tk_CheckTargetIdentity(wb, ws, wbFullName, wsName)
End Function

' "" when every block still holds what the command left there, else what changed.
Public Function FindConflict() As String
    Dim i As Long
    FindConflict = Tk_StyleProtectionProblem(ws, kind)
    If Len(FindConflict) > 0 Then Exit Function
    For i = 1 To blocks.Count
        If Not SameState(ReadState(blocks(i), blockEdges(i), True), after(i)) Then
            FindConflict = KindPhrase() & " in " & blocks(i).Address(False, False) & " on '" & wsName & _
                "' changed since the command."
            Exit Function
        End If
    Next i
End Function

Private Function KindPhrase() As String
    Select Case kind
    Case "fill": KindPhrase = "The fill"
    Case "borders": KindPhrase = "The borders"
    Case "rowheight": KindPhrase = "Row heights"
    Case Else: KindPhrase = "Column widths"
    End Select
End Function

' Puts every block's original state back, the selection first and the cells around it last.
' Never raises; "" on verified success.
Public Function Restore(ByVal useHook As Boolean) As String
    Dim i As Long, problems As String, limit As Long, done As Long
    If useHook Then limit = Tk_TestFailAfterRestores()
    For i = 1 To blocks.Count
        On Error Resume Next
        Err.Clear
        If limit > 0 And done >= limit Then Err.Raise vbObjectError + 517, "TkStyleSnapshot", "Simulated restore failure (test hook)"
        If Err.Number = 0 Then WriteState blocks(i), blockEdges(i), before(i)
        If Err.Number <> 0 Then
            problems = problems & " " & blocks(i).Address(False, False)
        Else
            done = done + 1
        End If
        On Error GoTo 0
    Next i
    If Len(problems) > 0 Then Restore = "Formatting not restored:" & problems & ".": Exit Function
    Restore = VerifyRestored()
End Function

Private Function VerifyRestored() As String
    Dim i As Long, problems As String
    On Error GoTo Unreadable
    For i = 1 To blocks.Count
        If Not SameState(ReadState(blocks(i), blockEdges(i), False), before(i)) Then
            problems = problems & " " & blocks(i).Address(False, False)
        End If
    Next i
    If Len(problems) > 0 Then VerifyRestored = "Formatting not restored exactly:" & problems & "."
    Exit Function
Unreadable:
    VerifyRestored = "The restored cells could not be read back to verify them: " & Err.Description
End Function
