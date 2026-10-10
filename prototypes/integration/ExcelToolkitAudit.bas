Attribute VB_Name = "ExcelToolkitAudit"
' ============================================================================
' SCAFFOLDING ONLY - NOT EXECUTED, NOT INSTALLED (offline integration prep).
'
' Minimal VBA front end for the end-to-end formula-audit demo:
'   1. one command to run an audit (bind to a ribbon button or Ctrl+Shift+J),
'   2. hand the LIVE workbook (incl. unsaved edits) to the Python engine,
'   3. show the findings in a separate results workbook (audited book untouched),
'   4. jump to a finding's cell.
'
' Deliberately basic: no task pane, no polished ribbon. The handoff identity matches
' the explicit, proven E2/E4 contract (window handle | pid | workbook name | path) so
' Python attaches to THIS Excel instance, reads the in-memory formulas via COM, builds a
' snapshot (worksheet_source.snapshot_from_com), and runs the SAME detector used offline.
'
' The engine path below is a placeholder to be wired in the authorised test environment;
' nothing here launches Python or COM on this machine.
' ============================================================================
Option Explicit

Private Const ENGINE_CMD As String = "<engine>"   ' set in the authorised test environment
Private Const RESULT_SEP As String = vbTab

' Entry point: run an audit on the active worksheet of the active workbook.
Public Sub Toolkit_AuditFormulas()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    ' Build the explicit handoff so Python attaches to THIS instance (see E2).
    Dim handoff As String
    handoff = Application.Hwnd & "|" & 0 & "|" & wb.Name & "|" & wb.FullName

    ' --- wiring to be completed in the authorised environment ---------------
    ' Python reads live formulas (unsaved included) and returns one finding per line:
    '   sheet <TAB> cell <TAB> kind <TAB> confidence <TAB> message
    ' Example (one-shot):  findings = RunEngine(ENGINE_CMD, handoff)
    Dim findings As String
    findings = ""   ' placeholder; no engine is launched here
    ' ------------------------------------------------------------------------

    ShowFindings wb, findings
End Sub

' Display findings in a SEPARATE workbook so the audited workbook is never modified.
Private Sub ShowFindings(auditedBook As Workbook, findings As String)
    Dim out As Workbook: Set out = Workbooks.Add
    Dim ws As Worksheet: Set ws = out.Worksheets(1)
    ws.Name = "Audit Findings"
    ws.Range("A1:D1").Value = Array("Sheet", "Cell", "Confidence", "Potential inconsistency (review)")

    Dim rows() As String, parts() As String, i As Long, r As Long
    If Len(findings) = 0 Then Exit Sub
    rows = Split(findings, vbLf)
    r = 2
    For i = LBound(rows) To UBound(rows)
        If Len(Trim$(rows(i))) > 0 Then
            parts = Split(rows(i), RESULT_SEP)
            ws.Cells(r, 1).Value = parts(0)
            ws.Cells(r, 2).Value = parts(1)
            ws.Cells(r, 3).Value = parts(3)
            ws.Cells(r, 4).Value = parts(4)
            ' A link back to the suspect cell in the audited workbook.
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 2), Address:="", _
                SubAddress:="'" & parts(0) & "'!" & parts(1), _
                TextToDisplay:=parts(1)
            r = r + 1
        End If
    Next i
    ws.Columns("A:D").AutoFit
End Sub

' Navigate to a finding: Goto 'Sheet'!Cell in the audited workbook (read-only; no change).
Public Sub Toolkit_GotoFinding(ByVal sheetName As String, ByVal cellRef As String)
    On Error Resume Next
    Application.Goto Reference:=ActiveWorkbook.Worksheets(sheetName).Range(cellRef), Scroll:=True
End Sub
