Attribute VB_Name = "PhaseBBridge"
' Phase B front end (OFFLINE-PREPARED - reviewed, NOT yet executed against live Excel).
'
' Talks to an ALREADY-RUNNING Python worker over E4's local named pipe and receives findings
' back through a COM callback (PhaseB_OnResult). It NEVER launches Python or any executable:
' there is no Shell, no one-shot, no auto-start. If no worker is listening, it says so and
' stops - the worker is started by hand from the development environment.
'
' Import PhaseBBridge.bas (this module) and PhaseBResultsHook.cls into the open development
' workbook's VBA project. No .xlam, no installer, no registry or trust changes.
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private Const RESULTS_FIRST_DATA_ROW As Long = 2
' Visible columns: A Cell | B Confidence | C Message. Hidden navigation columns: E Workbook |
' F Sheet | G Cell (read by the double-click handler to jump to the ORIGINAL workbook).
Private Const COL_NAV_WORKBOOK As Long = 5
Private Const COL_NAV_SHEET As Long = 6
Private Const COL_NAV_CELL As Long = 7

Private stateDir As String
Private requests As Object            ' id -> Scripting.Dictionary record
Private currentRequestId As String
Private requestCounter As Long
Private resultsWorkbookName As String
Private resultsSheetName As String
Private hook As PhaseBResultsHook     ' Application-level event sink for double-click

' ---------- Setup ----------

Public Sub PhaseB_Configure(ByVal state As String)
    stateDir = state
    Set requests = CreateObject("Scripting.Dictionary")
    currentRequestId = ""
    requestCounter = 0
    Set hook = New PhaseBResultsHook
    Set hook.App = Application        ' enables double-click navigation on the results sheet
    Randomize
End Sub

Private Function PipeName() As String
    ' Same per-user, local, inbound-only pipe the E4 worker creates (reused verbatim).
    PipeName = "\\.\pipe\excel_toolkit_e4_" & Environ("USERNAME")
End Function

Private Function JsonEscape(ByVal text As String) As String
    JsonEscape = Replace(Replace(text, "\", "\\"), """", "\""")
End Function

Private Function NewToken() As String
    ' A per-request correlation nonce so a stale callback cannot be mistaken for this result.
    NewToken = Hex$(CLng(Rnd * &H7FFFFFFF)) & Hex$(CLng(Timer * 1000) Mod &H7FFFFFFF)
End Function

' ---------- Request (write-only to the pipe; never launches anything) ----------

Public Sub PhaseB_AuditActiveSheet()
    Dim wb As Workbook, ws As Worksheet, id As String, token As String
    Dim record As Object, json As String
    On Error GoTo Failed

    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    If wb Is Nothing Or ws Is Nothing Then
        MsgBox "Open the workbook to audit and select its sheet first.", vbExclamation, "Excel Toolkit"
        Exit Sub
    End If

    requestCounter = requestCounter + 1
    id = GetCurrentProcessId() & "-" & requestCounter & "-" & Format(Now, "hhnnss")
    token = NewToken()

    Set record = CreateObject("Scripting.Dictionary")
    record("status") = "pending"
    record("token") = token
    record("pid") = GetCurrentProcessId()
    record("wb") = wb.Name
    record("full") = wb.FullName
    record("sheet") = ws.Name
    record("error") = ""
    requests.Add id, record
    currentRequestId = id             ' only callbacks for THIS id are current; older are stale

    json = "{""op"":""audit"",""id"":""" & id & """,""token"":""" & token & _
           """,""hwnd"":" & Application.Hwnd & ",""pid"":" & GetCurrentProcessId() & _
           ",""workbook"":""" & JsonEscape(wb.Name) & """,""full_name"":""" & JsonEscape(wb.FullName) & _
           """,""sheet"":""" & JsonEscape(ws.Name) & _
           """,""macro_workbook"":""" & JsonEscape(ThisWorkbook.Name) & """}"   ' qualifies the callback

    If Not TryWritePipe(json) Then
        record("status") = "failed"
        record("error") = "no worker listening"
        MsgBox "The audit worker is not running." & vbCrLf & vbCrLf & _
               "Start it manually from the development environment, then try again. " & _
               "This add-in will not start Python itself.", vbExclamation, "Excel Toolkit"
        Exit Sub
    End If
    ' The result arrives asynchronously through PhaseB_OnResult.
    Exit Sub
Failed:
    MsgBox "Could not send the audit request: " & Err.Description, vbExclamation, "Excel Toolkit"
End Sub

Private Function TryWritePipe(ByVal json As String) As Boolean
    Dim fileNo As Integer
    On Error GoTo Failed
    fileNo = FreeFile
    Open PipeName() For Binary Access Write As #fileNo   ' fails cleanly if no worker is listening
    Put #fileNo, , json & vbLf
    Close #fileNo
    TryWritePipe = True
    Exit Function
Failed:
    TryWritePipe = False
End Function

' ---------- Result callback (called by the Python worker via Application.Run) ----------

Public Function PhaseB_OnResult(ByVal id As String, ByVal token As String, ByVal pid As Long, _
                                ByVal workbookName As String, ByVal fullName As String, _
                                ByVal status As String, ByVal errorText As String, _
                                ByVal count As Long, ByVal findingsJson As String) As String
    Dim record As Object, verdict As String
    If requests Is Nothing Then Set requests = CreateObject("Scripting.Dictionary")

    verdict = ClassifyCallback(id, token, pid, workbookName, fullName)
    If verdict <> "accepted" Then
        PhaseB_OnResult = verdict           ' stale / unknown / identity or token mismatch -> ignored
        Exit Function
    End If

    Set record = requests(id)
    If status <> "ok" Then                  ' an error is NEVER shown as an empty successful audit
        record("status") = "failed"
        record("error") = errorText
        MsgBox "Audit failed (request " & id & "):" & vbCrLf & errorText, vbExclamation, "Excel Toolkit"
        PhaseB_OnResult = "accepted: reported failure"
        Exit Function
    End If

    record("status") = "done"
    record("count") = count
    RenderResults record, findingsJson
    PhaseB_OnResult = "accepted"
End Function

Private Function ClassifyCallback(ByVal id As String, ByVal token As String, ByVal pid As Long, _
                                  ByVal workbookName As String, ByVal fullName As String) As String
    ' Mirror of phaseb_contract.classify_callback (verified offline in test_phaseb.py).
    Dim rec As Object
    If Not requests.Exists(id) Then ClassifyCallback = "rejected: unknown request": Exit Function
    Set rec = requests(id)
    If rec("status") <> "pending" Then ClassifyCallback = "rejected: already " & rec("status"): Exit Function
    If id <> currentRequestId Then ClassifyCallback = "rejected: stale (superseded by a newer request)": Exit Function
    If pid <> rec("pid") Or workbookName <> rec("wb") Or fullName <> rec("full") Then
        ClassifyCallback = "rejected: identity mismatch": Exit Function
    End If
    If token <> rec("token") Then ClassifyCallback = "rejected: correlation token mismatch": Exit Function
    ClassifyCallback = "accepted"
End Function

' ---------- Results display (a SEPARATE workbook; the audited book is never touched) --------

Private Sub RenderResults(ByVal record As Object, ByVal findingsJson As String)
    Dim results As Workbook, sheet As Worksheet, items As Collection, obj As Object
    Dim r As Long, i As Long
    Set results = Workbooks.Add
    Set sheet = results.Worksheets(1)
    sheet.Name = "Audit findings"
    resultsWorkbookName = results.Name
    resultsSheetName = sheet.Name

    sheet.Cells(1, 1).Value = "Cell"
    sheet.Cells(1, 2).Value = "Confidence"
    sheet.Cells(1, 3).Value = "Potential inconsistency (please review)"
    sheet.Rows(1).Font.Bold = True

    Set items = ParseFlatFindings(findingsJson)
    r = RESULTS_FIRST_DATA_ROW
    For i = 1 To items.Count
        Set obj = items(i)
        sheet.Cells(r, 1).Value = DictGet(obj, "cell")
        sheet.Cells(r, 2).Value = DictGet(obj, "confidence")
        sheet.Cells(r, 3).Value = DictGet(obj, "message")        ' written as a value, never parsed as a formula
        sheet.Cells(r, COL_NAV_WORKBOOK).Value = DictGet(obj, "workbook")   ' hidden navigation data
        sheet.Cells(r, COL_NAV_SHEET).Value = DictGet(obj, "sheet")
        sheet.Cells(r, COL_NAV_CELL).Value = DictGet(obj, "cell")
        r = r + 1
    Next i
    If items.Count = 0 Then sheet.Cells(RESULTS_FIRST_DATA_ROW, 1).Value = "No inconsistencies detected."
    sheet.Columns("E:G").Hidden = True
    sheet.Columns("A:C").AutoFit
End Sub

Private Function DictGet(ByVal obj As Object, ByVal key As String) As String
    If obj.Exists(key) Then DictGet = CStr(obj(key)) Else DictGet = ""
End Function

' ---------- Double-click navigation (called by PhaseBResultsHook) ----------

Public Function PhaseB_HandleDoubleClick(ByVal Sh As Object, ByVal Target As Range) As Boolean
    ' True if this double-click was a results-row navigation (so the caller cancels edit mode).
    PhaseB_HandleDoubleClick = False
    If resultsWorkbookName = "" Then Exit Function
    If Sh.Parent.Name <> resultsWorkbookName Then Exit Function
    If Sh.Name <> resultsSheetName Then Exit Function
    If Target.Row < RESULTS_FIRST_DATA_ROW Then Exit Function

    Dim wbName As String, shName As String, cell As String
    wbName = CStr(Sh.Cells(Target.Row, COL_NAV_WORKBOOK).Value)
    shName = CStr(Sh.Cells(Target.Row, COL_NAV_SHEET).Value)
    cell = CStr(Sh.Cells(Target.Row, COL_NAV_CELL).Value)
    If wbName = "" Or shName = "" Or cell = "" Then Exit Function

    PhaseB_GotoFinding wbName, shName, cell
    PhaseB_HandleDoubleClick = True
End Function

Public Sub PhaseB_GotoFinding(ByVal workbookName As String, ByVal sheetName As String, ByVal cell As String)
    ' Resolve the ORIGINAL audited workbook via the object model (apostrophes/spaces safe),
    ' never the results workbook, and select the suspect cell. Read-only: navigation only.
    Dim wb As Workbook
    On Error GoTo Failed
    Set wb = Workbooks(workbookName)
    Application.Goto wb.Worksheets(sheetName).Range(cell), True
    Exit Sub
Failed:
    MsgBox "Could not navigate to " & workbookName & " / " & sheetName & " / " & cell & _
           " - is the audited workbook still open?", vbInformation, "Excel Toolkit"
End Sub

' ---------- Minimal JSON reader for the FLAT findings array ----------
' Handles exactly what display_json emits: an array of objects whose values are JSON strings
' or booleans (no nested arrays/objects). The wire format is pinned by offline tests
' (test_phaseb.py) so this consumer has a fixed contract to meet. Reviewed, not yet executed.

Private Function ParseFlatFindings(ByVal s As String) As Collection
    Dim items As New Collection, i As Long, n As Long, ch As String
    n = Len(s)
    i = InStr(s, "[")
    If i = 0 Then Set ParseFlatFindings = items: Exit Function
    i = i + 1
    Do While i <= n
        SkipWsOrComma s, i
        If i > n Then Exit Do
        ch = Mid$(s, i, 1)
        If ch = "]" Then Exit Do
        If ch = "{" Then
            items.Add ParseObject(s, i)
        Else
            i = i + 1
        End If
    Loop
    Set ParseFlatFindings = items
End Function

Private Function ParseObject(ByRef s As String, ByRef i As Long) As Object
    Dim obj As Object, key As String, ch As String, numStart As Long
    Set obj = CreateObject("Scripting.Dictionary")
    i = i + 1                               ' past '{'
    Do While i <= Len(s)
        SkipWs s, i
        ch = Mid$(s, i, 1)
        If ch = "}" Then i = i + 1: Exit Do
        If ch = "," Then i = i + 1: GoTo ContinuePair
        key = ParseString(s, i)
        SkipWs s, i
        If Mid$(s, i, 1) = ":" Then i = i + 1
        SkipWs s, i
        ch = Mid$(s, i, 1)
        If ch = """" Then
            obj(key) = ParseString(s, i)
        ElseIf ch = "t" Then
            obj(key) = True: i = i + 4
        ElseIf ch = "f" Then
            obj(key) = False: i = i + 5
        ElseIf ch = "n" Then
            obj(key) = "": i = i + 4         ' null -> empty
        Else
            numStart = i
            Do While i <= Len(s) And InStr("0123456789+-.eE", Mid$(s, i, 1)) > 0
                i = i + 1
            Loop
            obj(key) = Val(Mid$(s, numStart, i - numStart))
        End If
ContinuePair:
    Loop
    Set ParseObject = obj
End Function

Private Function ParseString(ByRef s As String, ByRef i As Long) As String
    Dim res As String, ch As String, nx As String
    i = i + 1                               ' past opening quote
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" Then
            nx = Mid$(s, i + 1, 1)
            Select Case nx
                Case """": res = res & """"
                Case "\": res = res & "\"
                Case "/": res = res & "/"
                Case "n": res = res & vbLf
                Case "r": res = res & vbCr
                Case "t": res = res & vbTab
                Case "b": res = res & Chr$(8)
                Case "f": res = res & Chr$(12)
                Case "u": res = res & ChrW$(CLng("&H" & Mid$(s, i + 2, 4))): i = i + 4
                Case Else: res = res & nx
            End Select
            i = i + 2
        ElseIf ch = """" Then
            i = i + 1
            Exit Do
        Else
            res = res & ch
            i = i + 1
        End If
    Loop
    ParseString = res
End Function

Private Sub SkipWs(ByRef s As String, ByRef i As Long)
    Do While i <= Len(s)
        Select Case Mid$(s, i, 1)
            Case " ", vbCr, vbLf, vbTab: i = i + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub SkipWsOrComma(ByRef s As String, ByRef i As Long)
    Do While i <= Len(s)
        Select Case Mid$(s, i, 1)
            Case " ", vbCr, vbLf, vbTab, ",": i = i + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub
