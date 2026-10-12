Attribute VB_Name = "TkMain"
' Excel Toolkit (M2): ribbon callbacks, the command registry and user messages.
'
' The registry is the single source of each command's label, tooltip and shortcut, so the
' ribbon cannot drift from it. Commands never launch Python or any other executable.
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
#End If

Public Const TK_TITLE As String = "Excel Toolkit"
Public Const TOOLKIT_VERSION As String = "0.3.0-m3a"

Private ribbon As IRibbonUI
Private quietMode As Boolean          ' test harness: record messages instead of showing them
Private lastMessage As String
Private auditReady As Boolean

' ---------- Command registry ----------

' Array(label, screentip, supertip, shortcut, modifies workbook, undo class) for a ribbon id.
Private Function Registry(ByVal id As String) As Variant
    Select Case id
    Case "tkConvertToValues"
        Registry = Array("Convert to Values", "Convert to Values", _
            "Replaces the formulas in the selected cells with their current results. Constants and " & _
            "number formats are left unchanged. Undo with Ctrl+Z straight away, or later with Undo " & _
            "Toolkit Action. Like any macro, this clears Excel's earlier undo history.", _
            "", True, "cells")
    Case "tkUndo"
        Registry = Array("Undo Toolkit Action", "Undo Toolkit Action", _
            "Reverses the last Excel Toolkit change, if the changed cells have not been edited since.", _
            "", True, "")
    Case "tkFmtNumber"
        Registry = Array("Number (0 / 0.0 / 0.00)", "Number format", _
            "Cycles the selection through #,##0 -> #,##0.0 -> #,##0.00 (negatives in red brackets, zero as a " & _
            "dash) - the same formats as the existing macro. Only the number format changes. Undo with Ctrl+Z " & _
            "straight away, or with Undo Toolkit Action.", _
            "Ctrl+Shift+I", True, "format")
    Case "tkFmtCurrency"
        Registry = Array("Currency " & ChrW(163) & " (0 / 0.0 / 0.00)", "Currency format", _
            "Cycles the selection through the " & ChrW(163) & " formats with 0, 1 and 2 decimals - the same formats " & _
            "as the existing macro. Undo with Ctrl+Z straight away, or with Undo Toolkit Action.", _
            "Ctrl+Shift+O", True, "format")
    Case "tkFmtPercent"
        Registry = Array("Percentage (0% to 0.000%)", "Percentage format", _
            "Cycles the selection through percentages with 0, 1, 2 and 3 decimals - the same formats as the " & _
            "existing macro. Undo with Ctrl+Z straight away, or with Undo Toolkit Action.", _
            "Ctrl+Shift+P", True, "format")
    Case "tkFmtDate"
        Registry = Array("Date (dd-mmm-yy / mmm-yy)", "Date format", _
            "Switches the selection between dd-mmm-yy and mmm-yy - the same formats as the existing macro. " & _
            "Undo with Ctrl+Z straight away, or with Undo Toolkit Action.", _
            "Ctrl+Shift+M", True, "format")
    Case "tkShortcuts"
        Registry = Array("Toolkit shortcuts", "Toolkit keyboard shortcuts", _
            "When on (this Excel session only), Ctrl+Shift+I, O, P and M run the Excel Toolkit number-format " & _
            "commands, which can be undone. When off, those keys do what they did before. Off by default.", _
            "", False, "none")
    Case "tkAudit"
        Registry = Array("Audit Formulas", "Audit Formulas", _
            "Checks the active sheet for formulas that break the pattern of their neighbours and lists " & _
            "them in a separate workbook. Read-only: the audited workbook is not changed. Needs the " & _
            "audit worker to be running.", _
            "", False, "none")
    Case Else
        Registry = Array(id, id, "", "", False, "none")
    End Select
End Function

Public Function Tk_CommandIds() As Variant
    Tk_CommandIds = Array("tkConvertToValues", "tkUndo", "tkFmtNumber", "tkFmtCurrency", "tkFmtPercent", _
                          "tkFmtDate", "tkShortcuts", "tkAudit")
End Function

' ---------- Ribbon callbacks (customUI14.xml) ----------

Public Sub Tk_RibbonLoaded(ByVal ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Public Sub Tk_RefreshRibbon()
    On Error Resume Next                  ' the ribbon reference is lost if VBA state was reset
    If Not ribbon Is Nothing Then ribbon.InvalidateControl "tkUndo"
End Sub

Public Sub Tk_RefreshShortcutToggle()
    On Error Resume Next
    If Not ribbon Is Nothing Then ribbon.InvalidateControl "tkShortcuts"
End Sub

Public Sub Tk_GetLabel(control As IRibbonControl, ByRef returnedVal)
    Dim info As Variant
    info = Registry(control.Id)
    returnedVal = info(0)
End Sub

Public Sub Tk_GetScreentip(control As IRibbonControl, ByRef returnedVal)
    Dim info As Variant
    info = Registry(control.Id)
    If Len(info(3)) > 0 Then returnedVal = info(1) & " (" & info(3) & " with Toolkit shortcuts on)" Else returnedVal = info(1)
End Sub

Public Sub Tk_GetSupertip(control As IRibbonControl, ByRef returnedVal)
    If control.Id = "tkUndo" Then
        If Tk_HasRecord() Then
            returnedVal = "Reverses " & Tk_RecordDescription() & ". Refused if those cells have changed since."
        Else
            returnedVal = "Nothing to undo."
        End If
    Else
        Dim info As Variant
        info = Registry(control.Id)
        returnedVal = info(2)
    End If
End Sub

Public Sub Tk_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    If control.Id = "tkUndo" Then returnedVal = Tk_HasRecord() Else returnedVal = True
End Sub

Public Sub Tk_OnAction(control As IRibbonControl)
    Select Case control.Id
    Case "tkConvertToValues": Tk_ConvertToValues
    Case "tkUndo": Tk_UndoToolkitAction
    Case "tkFmtNumber": Tk_RunFormat "number"
    Case "tkFmtCurrency": Tk_RunFormat "currency"
    Case "tkFmtPercent": Tk_RunFormat "percent"
    Case "tkFmtDate": Tk_RunFormat "date"
    Case "tkAudit": Tk_AuditFormulas
    End Select
End Sub

Public Sub Tk_GetPressed(control As IRibbonControl, ByRef returnedVal)
    If control.Id = "tkShortcuts" Then returnedVal = Tk_ShortcutsEnabled() Else returnedVal = False
End Sub

Public Sub Tk_OnToggle(control As IRibbonControl, pressed As Boolean)
    If control.Id = "tkShortcuts" Then Tk_SetShortcuts pressed
End Sub

' ---------- Commands (also runnable from Alt+F8) ----------

Public Sub Tk_ConvertToValues()
    If TypeName(Selection) <> "Range" Then
        Tk_Notify "refused: Select the cells to convert first."
        Exit Sub
    End If
    Tk_Notify Tk_ConvertRange(Selection)
End Sub

Public Sub Tk_AuditFormulas()
    If Not auditReady Then
        PhaseB_Configure Environ("LOCALAPPDATA") & "\ExcelToolkit"
        auditReady = True
    End If
    PhaseB_AuditActiveSheet
End Sub

' ---------- Messages ----------

' Shows a command outcome ("ok: ...", "nothing: ...", "refused: ...", "warning: ...",
' "rolled back: ...", "incomplete: ..."). Success is silent, like Excel's own paste.
Public Sub Tk_Notify(ByVal outcome As String)
    Dim kind As String, text As String, p As Long
    lastMessage = outcome
    If quietMode Then Exit Sub
    p = InStr(outcome, ": ")
    If p = 0 Then kind = outcome Else kind = Left$(outcome, p - 1): text = Mid$(outcome, p + 2)
    Select Case kind
    Case "ok"
    Case "nothing": MsgBox text, vbInformation, TK_TITLE
    Case "refused": MsgBox text, vbExclamation, TK_TITLE
    Case "warning": MsgBox text, vbExclamation, TK_TITLE
    Case "rolled back": MsgBox "The command could not finish and nothing was changed." & vbCrLf & vbCrLf & text, vbExclamation, TK_TITLE
    Case "incomplete": MsgBox "The command stopped part-way and could not fully restore the cells." & vbCrLf & vbCrLf & text, vbCritical, TK_TITLE
    Case Else: MsgBox outcome, vbInformation, TK_TITLE
    End Select
End Sub

Public Sub Tk_SetQuiet(ByVal quiet As Boolean)
    quietMode = quiet
End Sub

Public Function Tk_LastMessage() As String
    Tk_LastMessage = lastMessage
End Function

Public Function Tk_Version() As String
    Tk_Version = TOOLKIT_VERSION
End Function

' ---------- Timing ----------

Public Function Tk_Seconds() As Double
    Dim counter As Currency, frequency As Currency
    QueryPerformanceCounter counter
    QueryPerformanceFrequency frequency
    Tk_Seconds = counter / frequency
End Function
