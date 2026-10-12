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
'   borders   - every edge of every selected cell: line style, weight and colour (automatic, RGB,
'               or theme colour + tint). Adjacent cells share an edge (writing one side writes
'               the other), so each edge is held once: each cell's top and left edge, then the
'               bottom edge of the last row and the right edge of the last column - which are
'               also the facing edges of the cells around the selection.
'   rowheight - the height of every affected ROW (whole rows, not just the selected cells).
'               Hidden rows are refused: their height before hiding cannot be read.
'   colwidth  - the width of every affected COLUMN; hidden columns are refused likewise.
' Fills and sizes: a uniform range is one read; a mixed one is split recursively (columns first,
' then rows) into uniform blocks, as in TkFormatSnapshot. Borders cannot be read that way - for
' several cells Excel reports the FIRST cell's edge, never "mixed" (live run, 2026-10-12) - so
' every cell is read and equal neighbours are joined into runs and rectangles. More than
' MAX_BLOCKS blocks, or more than MAX_BORDER_CELLS cells for borders, is refused. Undo-record
' interface: Description, CheckIdentity, FindConflict, Restore (see TkUndo).
'
' Theme colours: ThemeColor reads 0 for an RGB colour and Null for a mix of theme and RGB
' colours (live run, 2026-10-12), so a uniform RGB area is one block. An error (seen only on
' automatic border colours, which are recognised first) is taken as RGB for one cell and
' splits a larger range.
Option Explicit

Private Const MAX_BLOCKS As Long = 50000
Private Const MAX_BORDER_CELLS As Long = 50000
Private Const NO_THEME As String = "rgb"

Private wb As Workbook
Private ws As Worksheet
Private wbFullName As String
Private wsName As String
Private targetAddress As String
Private commandLabel As String
Private capturedAt As Date
Private kind As String                ' "fill", "borders", "rowheight" or "colwidth"
Private blocks As Collection          ' Range per uniform block, in restore order
Private blockEdges As Collection      ' borders: the cell edges each block's state covers (else Empty)
Private before As Collection          ' each block's state before the command
Private after As Collection           ' and after it; the conflict check compares against this
Private refusal As String

' ---------- Capture (no writes) ----------

' Returns "" or the reason the command is refused.
Public Function Capture(ByVal rng As Range, ByVal kindName As String, ByVal label As String) As String
    Dim area As Range, part As Range, ok As Boolean, cellCount As Double
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

    If kind = "borders" Then
        For Each area In rng.Areas
            cellCount = cellCount + area.Cells.CountLarge
        Next area
        If cellCount > MAX_BORDER_CELLS Then
            Capture = "Grey Outline can be undone for up to " & Format(MAX_BORDER_CELLS, "#,##0") & " cells at a time " & _
                "(every cell's borders are recorded); the selection has " & Format(cellCount, "#,##0") & _
                ". Select fewer cells; nothing was changed."
            Exit Function
        End If
    End If

    For Each area In rng.Areas
        Select Case kind
        Case "rowheight"            ' column A of the affected rows: its RowHeight is the rows' height
            Set part = ws.Range(ws.Cells(area.Row, 1), ws.Cells(area.Row + area.Rows.Count - 1, 1))
            ok = AddBlocks(part)
        Case "colwidth"             ' row 1 of the affected columns
            Set part = ws.Range(ws.Cells(1, area.Column), ws.Cells(1, area.Column + area.Columns.Count - 1))
            ok = AddBlocks(part)
        Case "borders"
            ok = AddBorderBlocks(area)
        Case Else
            ok = AddBlocks(area)
        End Select
        If Not ok Then
            If Len(refusal) = 0 Then
                refusal = "The selection has too many differently formatted blocks (more than " & _
                    Format(MAX_BLOCKS, "#,##0") & ") to undo safely. Select a smaller range; nothing was changed."
            End If
            Capture = refusal
            Exit Function
        End If
    Next area
End Function

' Part of a range by corner cells. Never Offset (a whole column offset by half runs off the
' sheet) and never Rows(n)/Columns(n), whose result counts rows/columns rather than cells.
Private Function SubRange(ByVal rng As Range, ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long) As Range
    Set SubRange = ws.Range(rng.Cells(r1, c1), rng.Cells(r2, c2))
End Function

' Fills and sizes.
Private Function AddBlocks(ByVal rng As Range) As Boolean
    Dim state As Variant, rows As Long, cols As Long, half As Long, ok As Boolean
    state = ReadState(rng, Empty, False)
    If IsArray(state) Then
        AddBlocks = AddBlock(rng, Empty, state, False)
        Exit Function
    End If
    If Not IsNull(state) Then refusal = CStr(state): Exit Function         ' unsupported: refused
    If rng.Cells.CountLarge = 1 Then
        refusal = "The formatting of " & rng.Address(False, False) & " on '" & wsName & _
            "' could not be read exactly, so the change could not be undone; nothing was changed."
        Exit Function
    End If
    rows = rng.Rows.Count
    cols = rng.Columns.Count
    If cols > 1 Then
        half = cols \ 2
        ok = AddBlocks(SubRange(rng, 1, 1, rows, half))
        If ok Then ok = AddBlocks(SubRange(rng, 1, half + 1, rows, cols))
    Else
        half = rows \ 2
        ok = AddBlocks(SubRange(rng, 1, 1, half, 1))
        If ok Then ok = AddBlocks(SubRange(rng, half + 1, 1, rows, 1))
    End If
    AddBlocks = ok
End Function

' Borders: each row's top and left edges, then the last row's bottom edges and the last
' column's right edges, as runs of equal cells.
Private Function AddBorderBlocks(ByVal area As Range) As Boolean
    Dim r As Long, rows As Long, cols As Long
    rows = area.Rows.Count
    cols = area.Columns.Count
    For r = 1 To rows
        If Not AddRuns(SubRange(area, r, 1, r, cols), Array(xlEdgeTop, xlEdgeLeft)) Then Exit Function
    Next r
    If Not AddRuns(SubRange(area, rows, 1, rows, cols), Array(xlEdgeBottom)) Then Exit Function
    AddBorderBlocks = AddRuns(SubRange(area, 1, cols, rows, cols), Array(xlEdgeRight))
End Function

' Splits a single row or column into runs of cells whose edges are equal.
Private Function AddRuns(ByVal strip As Range, ByVal cellEdges As Variant) As Boolean
    Dim i As Long, n As Long, first As Long, state As Variant, runState As Variant
    n = strip.Cells.Count
    For i = 1 To n
        state = CellBorderState(strip.Cells(i), cellEdges)
        If Not IsArray(state) Then
            refusal = "The borders of " & strip.Cells(i).Address(False, False) & " on '" & wsName & _
                "' could not be read exactly, so the change could not be undone; nothing was changed."
            Exit Function
        End If
        If i = 1 Then
            first = 1
            runState = state
        ElseIf Not SameState(state, runState) Then
            If Not AddBlock(ws.Range(strip.Cells(first), strip.Cells(i - 1)), cellEdges, runState, True) Then Exit Function
            first = i
            runState = state
        End If
    Next i
    AddRuns = AddBlock(ws.Range(strip.Cells(first), strip.Cells(n)), cellEdges, runState, True)
End Function

' Adds a block; with `canJoin`, a one-row run directly below an equal run of the same columns
' extends it into a rectangle (so a uniformly bordered area restores in a few writes).
Private Function AddBlock(ByVal rng As Range, ByVal edges As Variant, ByVal state As Variant, ByVal canJoin As Boolean) As Boolean
    Dim last As Range, n As Long
    n = blocks.Count
    If canJoin And n > 0 Then
        Set last = blocks(n)
        If last.Column = rng.Column And last.Columns.Count = rng.Columns.Count And rng.Rows.Count = 1 Then
            If last.Row + last.Rows.Count = rng.Row Then
                If SameState(blockEdges(n), edges) And SameState(before(n), state) Then
                    Set rng = ws.Range(last.Cells(1, 1), rng.Cells(1, rng.Columns.Count))
                    blocks.Remove n
                    blockEdges.Remove n
                    before.Remove n
                End If
            End If
        End If
    End If
    If blocks.Count >= MAX_BLOCKS Then Exit Function
    blocks.Add rng
    blockEdges.Add edges
    before.Add state
    AddBlock = True
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
    If lenient Then cellCount = 1 Else cellCount = rng.Cells.CountLarge
    Select Case kind
    Case "fill": ReadState = FillState(rng, cellCount)
    Case "borders": ReadState = BorderState(rng, edges)
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

' The cell edges shared by every cell of the range, or Null. Read cell by cell: for several
' cells Excel reports the first cell's edge, not "mixed".
Private Function BorderState(ByVal rng As Range, ByVal cellEdges As Variant) As Variant
    Dim cell As Range, first As Variant, one As Variant
    For Each cell In rng.Cells
        one = CellBorderState(cell, cellEdges)
        If Not IsArray(one) Then BorderState = Null: Exit Function
        If IsEmpty(first) Then
            first = one
        ElseIf Not SameState(first, one) Then
            BorderState = Null
            Exit Function
        End If
    Next cell
    BorderState = first
End Function

Private Function CellBorderState(ByVal cell As Range, ByVal cellEdges As Variant) As Variant
    Dim states() As Variant, i As Long, one As Variant
    ReDim states(LBound(cellEdges) To UBound(cellEdges))
    For i = LBound(cellEdges) To UBound(cellEdges)
        one = EdgeState(cell.Borders(cellEdges(i)), 1)
        If Not IsArray(one) Then CellBorderState = Null: Exit Function
        states(i) = one
    Next i
    CellBorderState = states
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
        ' A cell edge written on a block: tops (and the rows inside), lefts (and the columns
        ' inside), or the bottom/right edge of a one-row/one-column strip.
        For i = LBound(edges) To UBound(edges)
            Select Case edges(i)
            Case xlEdgeTop
                WriteEdge rng.Borders(xlEdgeTop), state(i)
                If rng.Rows.Count > 1 Then WriteEdge rng.Borders(xlInsideHorizontal), state(i)
            Case xlEdgeLeft
                WriteEdge rng.Borders(xlEdgeLeft), state(i)
                If rng.Columns.Count > 1 Then WriteEdge rng.Borders(xlInsideVertical), state(i)
            Case Else
                WriteEdge rng.Borders(edges(i)), state(i)
            End Select
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
        state = AfterState(i)
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
    Dim saved As Variant
    FindConflict = Tk_StyleProtectionProblem(ws, kind)
    If Len(FindConflict) > 0 Then Exit Function
    saved = Tk_QuietExcel()                   ' cell-by-cell border reads are several times faster
    FindConflict = FirstConflict()
    Tk_RestoreExcel saved
End Function

' Never raises.
Private Function FirstConflict() As String
    Dim i As Long
    On Error GoTo Unreadable
    For i = 1 To blocks.Count
        If Not SameState(AfterState(i), after(i)) Then
            FirstConflict = KindPhrase() & " in " & blocks(i).Address(False, False) & " on '" & wsName & _
                "' changed since the command."
            Exit Function
        End If
    Next i
    Exit Function
Unreadable:
    FirstConflict = "The changed cells could not be read to check them: " & Err.Description
End Function

' A block's state after the command, for the conflict check. Fills and sizes stay uniform within
' each block; borders need not (on a multi-area selection the legacy outline is applied to the
' whole selection at once), so they are recorded cell by cell.
Private Function AfterState(ByVal i As Long) As Variant
    Dim cell As Range, states() As Variant, n As Long
    If kind <> "borders" Then AfterState = ReadState(blocks(i), blockEdges(i), True): Exit Function
    ReDim states(1 To blocks(i).Cells.Count)
    For Each cell In blocks(i).Cells
        n = n + 1
        states(n) = CellBorderState(cell, blockEdges(i))
        If Not IsArray(states(n)) Then AfterState = Null: Exit Function
    Next cell
    AfterState = states
End Function

Private Function KindPhrase() As String
    Select Case kind
    Case "fill": KindPhrase = "The fill"
    Case "borders": KindPhrase = "The borders"
    Case "rowheight": KindPhrase = "Row heights"
    Case Else: KindPhrase = "Column widths"
    End Select
End Function

' Puts every block's original state back.
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
