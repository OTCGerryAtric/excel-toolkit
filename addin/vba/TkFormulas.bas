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

' The formula with every reference prefix naming `sheetName` removed. Never changes anything
' inside a string literal, a quoted name of another sheet or workbook, a 3D range or an error value.
Public Function Tk_RemoveSheetPrefix(ByVal formula As String, ByVal sheetName As String) As String
    Dim result As String, i As Long, j As Long, n As Long, ch As String, prevCh As String, nameText As String
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
                    nameText = Replace(Mid$(formula, i + 1, j - i - 1), "''", "'")
                    If StrComp(nameText, sheetName, vbTextCompare) = 0 Then
                        i = j + 2                                ' drop 'name'!
                        GoTo NextToken
                    End If
                End If
            End If
            result = result & Mid$(formula, i, j - i + 1)
            i = j + 1
        ElseIf IsNameChar(ch) Then                              ' a word: function, name, cell or sheet
            j = i
            Do While j < n
                If Not IsNameChar(Mid$(formula, j + 1, 1)) Then Exit Do
                j = j + 1
            Loop
            If i > 1 Then prevCh = Mid$(formula, i - 1, 1) Else prevCh = ""
            If IsOwnSheetWord(formula, i, j, n, prevCh, sheetName) Then
                i = j + 2                                        ' drop name!
            Else
                result = result & Mid$(formula, i, j - i + 1)
                i = j + 1
            End If
        Else
            result = result & ch
            i = i + 1
        End If
NextToken:
    Loop
    Tk_RemoveSheetPrefix = result
End Function

' A word followed by "!" that names this sheet - unless it follows "]" (an external workbook),
' "#" (an error value such as #REF!) or a ":" that ends a 3D range (Sheet1:Sheet3!A1). A ":"
' after a cell reference or ")" is an ordinary range (Sheet1!A1:Sheet1!B2), so the prefix goes.
Private Function IsOwnSheetWord(ByVal formula As String, ByVal i As Long, ByVal j As Long, ByVal n As Long, _
                                ByVal prevCh As String, ByVal sheetName As String) As Boolean
    If j >= n Then Exit Function
    If Mid$(formula, j + 1, 1) <> "!" Then Exit Function
    If prevCh = "]" Or prevCh = "#" Then Exit Function
    If prevCh = ":" Then
        If Not RangeBeforeColon(formula, i - 1) Then Exit Function
    End If
    IsOwnSheetWord = (StrComp(Mid$(formula, i, j - i + 1), sheetName, vbTextCompare) = 0)
End Function

' Whether what ends just before the ":" at `colon` is a cell, column or row reference ($A$1, A1, A,
' 1) or a closing bracket - rather than a sheet name starting a 3D range.
Private Function RangeBeforeColon(ByVal formula As String, ByVal colon As Long) As Boolean
    Dim k As Long, word As String, ch As String, letters As Long, digits As Long, p As Long
    k = colon - 1
    If k < 1 Then Exit Function
    If Mid$(formula, k, 1) = ")" Then RangeBeforeColon = True: Exit Function
    Do While k >= 1
        ch = Mid$(formula, k, 1)
        If ch = "$" Then
            k = k - 1
        ElseIf IsNameChar(ch) Then
            word = ch & word
            k = k - 1
        Else
            Exit Do
        End If
    Loop
    If Len(word) = 0 Then Exit Function
    For p = 1 To Len(word)                                       ' letters, then digits, nothing else
        ch = Mid$(word, p, 1)
        If ch Like "[A-Za-z]" Then
            If digits > 0 Then Exit Function
            letters = letters + 1
        ElseIf ch Like "#" Then
            digits = digits + 1
        Else
            Exit Function
        End If
    Next p
    RangeBeforeColon = (letters <= 3 And digits <= 7)
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
        Tk_ClearSheetReferences = "nothing: No formula in " & ScopeText(target) & " refers to '" & target.Worksheet.Name & _
            "' by name."
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
