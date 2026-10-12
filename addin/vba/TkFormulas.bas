Attribute VB_Name = "TkFormulas"
' Excel Toolkit (M3c): Clear Sheet References and Find Constants, ported from PERSONAL.XLSB
' (legacy keys Ctrl+Shift+F and, formerly, C; Find Constants moves to H when shortcuts migrate).
'
' Clear Sheet References removes the current sheet's own name from references in FORMULAS
' ('P&L'!A1 -> A1 on the sheet P&L), which does not change what they refer to. The legacy macro did
' a plain text Replace across every cell; this version reads each formula and only removes a
' sheet prefix that names this sheet, so text cells, string literals ("P&L!A1"), other sheets,
' similarly named sheets (P&L 2), external workbooks ([Book.xlsx]P&L!A1), 3D ranges (Sheet1:Sheet3!A1)
' and error values (#REF!) are never changed. Undo restores every changed formula exactly.
'
' Find Constants selects the typed-in numbers in the selection - the same as the legacy macro. It
' only changes the selection, so it needs no undo.
'
' Scope: the selection; a single selected cell means the whole sheet (as Excel's Find & Replace
' and the legacy macro, which always worked on the whole sheet).
Option Explicit

Private lastTimings As String

Private Const STAGE_CHECKING As Long = 0
Private Const STAGE_WRITING As Long = 1
Private Const STAGE_WRITTEN As Long = 2

' ---------- The rewrite (pure; no Excel calls) ----------

' The formula with every reference prefix naming `sheetName` removed (see Tk_RewriteSheetPrefix).
Public Function Tk_RemoveSheetPrefix(ByVal formula As String, ByVal sheetName As String) As String
    Dim ambiguous As Boolean
    Tk_RemoveSheetPrefix = Tk_RewriteSheetPrefix(formula, sheetName, ambiguous)
End Function

' The formula with every reference prefix naming `sheetName` removed. Never changes anything inside a
' string literal, a quoted name of another sheet or workbook, a 3D range or an error value, nor the
' second half of a range whose first half is on another sheet. A formula that would change but holds a
' range Excel could read two ways (see HasAmbiguousColon) is returned unchanged, with `ambiguous` set.
Public Function Tk_RewriteSheetPrefix(ByVal formula As String, ByVal sheetName As String, _
                                      ByRef ambiguous As Boolean) As String
    Dim result As String
    ambiguous = False
    result = RemovePrefixes(formula, sheetName)
    If result <> formula Then
        If HasAmbiguousColon(formula) Or HasAmbiguousColon(result) Then
            ambiguous = True
            result = formula
        End If
    End If
    Tk_RewriteSheetPrefix = result
End Function

Private Function RemovePrefixes(ByVal formula As String, ByVal sheetName As String) As String
    Dim result As String, i As Long, j As Long, n As Long, ch As String
    n = Len(formula)
    i = 1
    Do While i <= n
        ch = Mid$(formula, i, 1)
        If ch = """" Then                                       ' string literal: copied as it is
            j = QuoteEnd(formula, i, """")
            result = result & Mid$(formula, i, j - i + 1)
            i = j + 1
        ElseIf ch = "'" Then                                    ' quoted name: 'P&L'!, 'O''Brien'!
            j = QuoteEnd(formula, i, "'")
            If j < n Then
                If Mid$(formula, j + 1, 1) = "!" Then
                    If OwnPrefix(formula, i, Replace(Mid$(formula, i + 1, j - i - 1), "''", "'"), sheetName) Then
                        i = j + 2                                ' drop 'name'!
                        GoTo NextToken
                    End If
                End If
            End If
            result = result & Mid$(formula, i, j - i + 1)
            i = j + 1
        ElseIf IsNameChar(ch) Then                              ' a word: function, name, cell or sheet
            j = WordEnd(formula, i)
            If j < n Then
                If Mid$(formula, j + 1, 1) = "!" Then
                    If OwnPrefix(formula, i, Mid$(formula, i, j - i + 1), sheetName) Then
                        i = j + 2                                ' drop name!
                        GoTo NextToken
                    End If
                End If
            End If
            result = result & Mid$(formula, i, j - i + 1)
            i = j + 1
        Else
            result = result & ch
            i = i + 1
        End If
NextToken:
    Loop
    RemovePrefixes = result
End Function

' Whether the sheet prefix starting at `i` (its name is `nameText`, followed by "!") can go: it names
' this sheet, it does not follow "]" (an external workbook) or "#" (an error value such as #REF!), and
' if it follows ":" the first half of that range is a cell on this sheet or ends with ")" - otherwise
' the ":" ends a 3D range (Sheet1:Sheet3!A1, A:Sheet1!B2) or the first half is on another sheet
' (Sheet2!A1:Sheet1!B2 would become Sheet2!A1:B2).
Private Function OwnPrefix(ByVal formula As String, ByVal i As Long, ByVal nameText As String, _
                           ByVal sheetName As String) As Boolean
    Dim prevCh As String
    If StrComp(nameText, sheetName, vbTextCompare) <> 0 Then Exit Function
    If i > 1 Then prevCh = Mid$(formula, i - 1, 1)
    If prevCh = "]" Or prevCh = "#" Then Exit Function
    If prevCh = ":" Then
        If Not RangeHalfBefore(formula, i - 1, sheetName) Then Exit Function
    End If
    OwnPrefix = True
End Function

' Whether what ends just before the ":" at `colon` is the first half of an ordinary range on this
' sheet: ")" (INDEX(...):), or a cell reference ($A$1, A1) with no prefix or this sheet's prefix.
' Excel reads a column, row or any other word there (A:Sheet1!B2, 1:Sheet1!B2) as a 3D range.
Private Function RangeHalfBefore(ByVal formula As String, ByVal colon As Long, ByVal sheetName As String) As Boolean
    Dim k As Long
    If colon <= 1 Then Exit Function
    If Mid$(formula, colon - 1, 1) = ")" Then RangeHalfBefore = True: Exit Function
    k = TokenStart(formula, colon - 1)
    If Not IsCellReference(Mid$(formula, k, colon - k)) Then Exit Function
    If k > 1 Then
        If Mid$(formula, k - 1, 1) = "!" Then                   ' prefixed first half: must be this sheet
            RangeHalfBefore = (StrComp(PrefixNameBefore(formula, k - 1), sheetName, vbTextCompare) = 0)
            Exit Function
        End If
    End If
    RangeHalfBefore = True
End Function

' Whether the formula holds a ":" followed by a sheet prefix whose first half is neither ")", a cell
' reference nor itself prefixed: Excel stores AB:Sheet1!B2 the same way whether it was entered as a
' 3D range or not, and reads it back as a 3D range, so such a formula is never rewritten.
Private Function HasAmbiguousColon(ByVal formula As String) As Boolean
    Dim i As Long, n As Long, ch As String, k As Long
    n = Len(formula)
    i = 1
    Do While i <= n
        ch = Mid$(formula, i, 1)
        If ch = """" Or ch = "'" Then
            i = QuoteEnd(formula, i, ch) + 1
        ElseIf ch = ":" Then
            If PrefixAt(formula, i + 1) Then
                If i = 1 Then HasAmbiguousColon = True: Exit Function
                If Mid$(formula, i - 1, 1) <> ")" Then
                    k = TokenStart(formula, i - 1)
                    If k = i Then HasAmbiguousColon = True: Exit Function
                    If Not IsCellReference(Mid$(formula, k, i - k)) Then
                        If k = 1 Then HasAmbiguousColon = True: Exit Function
                        If Mid$(formula, k - 1, 1) <> "!" Then HasAmbiguousColon = True: Exit Function
                    End If
                End If
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' Whether a sheet prefix (name! or 'name'!) starts at `p`.
Private Function PrefixAt(ByVal formula As String, ByVal p As Long) As Boolean
    Dim j As Long, n As Long
    n = Len(formula)
    If p > n Then Exit Function
    If Mid$(formula, p, 1) = "'" Then
        j = QuoteEnd(formula, p, "'")
    ElseIf IsNameChar(Mid$(formula, p, 1)) Then
        j = WordEnd(formula, p)
    Else
        Exit Function
    End If
    If j < n Then PrefixAt = (Mid$(formula, j + 1, 1) = "!")
End Function

' The sheet name of the prefix whose "!" is at `bang` (unescaped if quoted); "" for an external
' reference or when there is none.
Private Function PrefixNameBefore(ByVal formula As String, ByVal bang As Long) As String
    Dim k As Long, closing As Long
    If bang <= 1 Then Exit Function
    If Mid$(formula, bang - 1, 1) = "'" Then
        closing = bang - 1
        k = closing - 1
        Do While k >= 1
            If Mid$(formula, k, 1) = "'" Then
                If k > 1 Then
                    If Mid$(formula, k - 1, 1) = "'" Then k = k - 2: GoTo NextChar
                End If
                Exit Do
            End If
            k = k - 1
NextChar:
        Loop
        If k < 1 Then Exit Function
        PrefixNameBefore = Replace(Mid$(formula, k + 1, closing - k - 1), "''", "'")
    Else
        k = bang - 1
        Do While k >= 1
            If Not IsNameChar(Mid$(formula, k, 1)) Then Exit Do
            k = k - 1
        Loop
        If k >= 1 Then
            If Mid$(formula, k, 1) = "]" Then Exit Function
        End If
        PrefixNameBefore = Mid$(formula, k + 1, bang - k - 1)
    End If
End Function

' Start of the reference or word ending at `last` (name characters and "$"); last + 1 if none.
Private Function TokenStart(ByVal formula As String, ByVal last As Long) As Long
    Dim k As Long, ch As String
    k = last
    Do While k >= 1
        ch = Mid$(formula, k, 1)
        If ch <> "$" And Not IsNameChar(ch) Then Exit Do
        k = k - 1
    Loop
    TokenStart = k + 1
End Function

' End of the word starting at `i`.
Private Function WordEnd(ByVal formula As String, ByVal i As Long) As Long
    Dim j As Long, n As Long
    n = Len(formula)
    j = i
    Do While j < n
        If Not IsNameChar(Mid$(formula, j + 1, 1)) Then Exit Do
        j = j + 1
    Loop
    WordEnd = j
End Function

' Whether `token` is a cell reference within the sheet: [$]A-XFD[$]1-1048576.
Private Function IsCellReference(ByVal token As String) As Boolean
    Dim p As Long, n As Long, colNumber As Long, rowText As String, ch As String
    n = Len(token)
    p = 1
    If p <= n Then
        If Mid$(token, p, 1) = "$" Then p = p + 1
    End If
    Do While p <= n
        ch = UCase$(Mid$(token, p, 1))
        If Not ch Like "[A-Z]" Then Exit Do
        If colNumber > 16384 Then Exit Function
        colNumber = colNumber * 26 + (Asc(ch) - 64)
        p = p + 1
    Loop
    If colNumber < 1 Or colNumber > 16384 Then Exit Function
    If p <= n Then
        If Mid$(token, p, 1) = "$" Then p = p + 1
    End If
    rowText = Mid$(token, p)
    If Len(rowText) < 1 Or Len(rowText) > 7 Then Exit Function
    If rowText Like "*[!0-9]*" Then Exit Function
    If Left$(rowText, 1) = "0" Then Exit Function
    IsCellReference = (CLng(rowText) <= 1048576)
End Function

' Position of the quote that closes the one at `start` (a doubled quote is an escaped quote), or
' the end of the formula if it is never closed.
Private Function QuoteEnd(ByVal formula As String, ByVal start As Long, ByVal quote As String) As Long
    Dim k As Long, n As Long
    n = Len(formula)
    k = start + 1
    Do While k <= n
        If Mid$(formula, k, 1) = quote Then
            If k < n Then
                If Mid$(formula, k + 1, 1) = quote Then
                    k = k + 2
                Else
                    QuoteEnd = k
                    Exit Function
                End If
            Else
                QuoteEnd = k
                Exit Function
            End If
        Else
            k = k + 1
        End If
    Loop
    QuoteEnd = n
End Function

' Characters of an unquoted sheet name, defined name, function name or cell reference.
Private Function IsNameChar(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    IsNameChar = (ch Like "[A-Za-z0-9_.]") Or code > 127
End Function

' ---------- Clear Sheet References ----------

Public Sub Tk_RunClearSheetReferences()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_ClearSheetReferences(Selection)
End Sub

' Returns "ok: / nothing: / refused: / warning: / rolled back: / incomplete: <text>".
Public Function Tk_ClearSheetReferences(ByVal target As Range) As String
    Dim snap As TkFormulaSnapshot, scope As Range, reason As String, rollbackProblem As String
    Dim stage As Long, errorText As String, saved As Variant, t0 As Double, t1 As Double, t2 As Double, t3 As Double

    On Error GoTo Failed
    t0 = Tk_Seconds()
    If target.Worksheet.ProtectContents Then
        Tk_ClearSheetReferences = "refused: The sheet is protected, so its formulas cannot be changed. Unprotect it first; nothing was changed."
        Exit Function
    End If
    If target.CountLarge = 1 Then Set scope = target.Worksheet.UsedRange Else Set scope = target

    Set snap = New TkFormulaSnapshot
    reason = snap.Capture(scope, target.Worksheet.Name, "Clear Sheet References")
    If Len(reason) > 0 Then Tk_ClearSheetReferences = "refused: " & reason: Exit Function
    If snap.CellCount = 0 Then
        If snap.AmbiguousCount > 0 Then
            Tk_ClearSheetReferences = "nothing: Nothing was changed. " & AmbiguousText(snap)
        Else
            Tk_ClearSheetReferences = "nothing: No formula in " & ScopeText(target) & " refers to '" & _
                target.Worksheet.Name & "' by name."
        End If
        Exit Function
    End If
    t1 = Tk_Seconds()

    saved = Tk_QuietExcel()
    stage = STAGE_WRITING
    snap.Apply
    reason = snap.VerifyApplied()
    If Len(reason) > 0 Then GoTo RollBack
    stage = STAGE_WRITTEN
    Tk_RestoreExcel saved
    t2 = Tk_Seconds()

    Tk_SetRecord snap, "Clear Sheet References"          ' replaces any earlier record; registers Ctrl+Z
    t3 = Tk_Seconds()
    lastTimings = "cells=" & snap.CellCount & ";capture=" & Ms(t1 - t0) & ";apply+verify=" & Ms(t2 - t1) & _
                  ";register=" & Ms(t3 - t2) & ";total=" & Ms(t3 - t0)
    Tk_ClearSheetReferences = "ok: Removed '" & target.Worksheet.Name & "'! from " & snap.CellCount & " formula(s) in " & _
        ScopeText(target) & "."
    If snap.AmbiguousCount > 0 Then
        Tk_ClearSheetReferences = "warning: " & Mid$(Tk_ClearSheetReferences, 5) & " " & AmbiguousText(snap)
    End If
    Exit Function

Failed:
    errorText = Err.Description
    Select Case stage
    Case STAGE_CHECKING
        Tk_ClearSheetReferences = "refused: Excel reported an error before anything was changed: " & errorText
        Exit Function
    Case STAGE_WRITING
        reason = "Excel reported an error while rewriting formulas: " & errorText
        Resume RollBack
    Case Else
        If Tk_RecordIs(snap) Then
            Tk_ClearSheetReferences = "warning: The formulas were rewritten and checked, but Ctrl+Z could not be set up (" & _
                errorText & "). Use Undo Toolkit Action to reverse it."
            Exit Function
        End If
        reason = "The formulas were rewritten, but the undo record could not be stored (" & errorText & ")."
        Resume RollBack
    End Select
RollBack:
    On Error GoTo 0
    rollbackProblem = snap.Restore(False)
    If IsArray(saved) Then Tk_RestoreExcel saved
    If Len(rollbackProblem) = 0 Then
        Tk_ClearSheetReferences = "rolled back: " & reason
    Else
        Tk_ClearSheetReferences = "incomplete: " & reason & " Putting the original formulas back also failed: " & rollbackProblem
    End If
End Function

Private Function AmbiguousText(ByVal snap As TkFormulaSnapshot) As String
    AmbiguousText = snap.AmbiguousCount & " formula(s) naming this sheet were left unchanged (the first is in " & _
        snap.FirstAmbiguous & ") because they contain a range Excel could read two ways, such as A:Sheet1!B2 " & _
        "(a range across sheets, or a column to a cell). Check those by hand."
End Function

Private Function ScopeText(ByVal target As Range) As String
    If target.CountLarge = 1 Then ScopeText = "the sheet" Else ScopeText = "the selection"
End Function

Public Function Tk_LastFormulaTimings() As String
    Tk_LastFormulaTimings = lastTimings
End Function

' ---------- Find Constants (selection only) ----------

Public Sub Tk_RunFindConstants()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Tk_Notify Tk_FindConstants(Selection)
End Sub

' Selects the typed-in numbers in `target` (a single cell: the whole sheet, as Excel does).
Public Function Tk_FindConstants(ByVal target As Range) As String
    Dim found As Range
    On Error Resume Next
    Set found = target.SpecialCells(xlCellTypeConstants, xlNumbers)
    On Error GoTo 0
    If found Is Nothing Then
        Tk_FindConstants = "nothing: There are no typed-in numbers in " & ScopeText(target) & "."
        Exit Function
    End If
    found.Worksheet.Activate
    found.Select
    Tk_FindConstants = "ok: Selected " & Format(found.CountLarge, "#,##0") & " cell(s) with typed-in numbers."
End Function

Private Function Ms(ByVal seconds As Double) As String
    Ms = Format(seconds * 1000, "0.0")
End Function
