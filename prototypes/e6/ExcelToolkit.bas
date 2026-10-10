Attribute VB_Name = "ExcelToolkit"
' Phase 0 / E6 prototype add-in: packaged bridge to the Python engine.
' Finds the engine next to this add-in; worker first, one-shot fallback.
' Throwaway code - not production.
Option Explicit

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal access As Long, ByVal inherit As Long, ByVal pid As Long) As LongPtr
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal handle As LongPtr, ByRef code As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Private Const VERSION_TEXT As String = "{{VERSION}}"
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const STILL_ACTIVE As Long = 259
Private Const REQUEST_TIMEOUT_SECONDS As Double = 60
Private Const WORKER_SILENT_SECONDS As Double = 15
Private Const AUDIT_KEY As String = "^+J"

Private requests As Object
Private requestCounter As Long
Private watchdogScheduled As Boolean
Private autoOpenRan As Boolean
Private quietMode As Boolean
Private lastShownId As String

' ---------- Lifecycle ----------

Public Sub Auto_Open()
    autoOpenRan = True
    Application.OnKey AUDIT_KEY, "Toolkit_AuditSelection"
End Sub

Public Sub Auto_Close()
    Application.OnKey AUDIT_KEY
End Sub

Public Function Toolkit_Version() As String
    Toolkit_Version = VERSION_TEXT
End Function

Public Function Toolkit_Status() As String
    Toolkit_Status = "version=" & VERSION_TEXT & "|auto_open=" & autoOpenRan & "|engine=" & EnginePath() & _
                     "|engine_exists=" & (Dir(EnginePath()) <> "") & "|state=" & StateDir()
End Function

Public Sub Toolkit_SetQuiet(ByVal quiet As Boolean)
    quietMode = quiet
End Sub

Private Function EnginePath() As String
    EnginePath = ThisWorkbook.Path & "\engine\ExcelToolkitEngine.exe"
End Function

Private Function StateDir() As String
    StateDir = Environ("LOCALAPPDATA") & "\ExcelToolkit\state"
    If Dir(Environ("LOCALAPPDATA") & "\ExcelToolkit", vbDirectory) = "" Then MkDir Environ("LOCALAPPDATA") & "\ExcelToolkit"
    If Dir(StateDir, vbDirectory) = "" Then MkDir StateDir
End Function

Private Function Seconds() As Double
    Dim counter As Currency, frequency As Currency
    QueryPerformanceCounter counter
    QueryPerformanceFrequency frequency
    Seconds = counter / frequency
End Function

Private Function JsonEscape(ByVal text As String) As String
    JsonEscape = Replace(Replace(text, "\", "\\"), """", "\""")
End Function

Private Function PipeName() As String
    PipeName = "\\.\pipe\excel_toolkit_" & Environ("USERNAME")
End Function

' ---------- User command ----------

Public Sub Toolkit_AuditSelection()
    Dim id As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    id = Toolkit_Request("auto", "classify", Selection.Worksheet.Name, Selection.Address(False, False), Selection.Worksheet.Parent.Name)
    If Left$(id, 5) = "ERROR" And Not quietMode Then MsgBox id, vbExclamation, "Excel Toolkit"
End Sub

' ---------- Requests ----------

' mode: "auto" (worker, falling back to one-shot), "worker", "oneshot", or "none" (register only, for tests)
Public Function Toolkit_Request(ByVal mode As String, ByVal op As String, ByVal sheetName As String, _
                                ByVal address As String, ByVal workbookName As String) As String
    Dim id As String, json As String, record As Object, target As Workbook
    On Error GoTo Failed
    If requests Is Nothing Then Set requests = CreateObject("Scripting.Dictionary")
    Set target = Workbooks(workbookName)
    requestCounter = requestCounter + 1
    id = GetCurrentProcessId() & "-" & requestCounter & "-" & Format(Now, "hhnnss")

    Set record = CreateObject("Scripting.Dictionary")
    record("status") = "pending"
    record("start") = Seconds()
    record("wb") = target.Name
    record("full") = target.FullName
    record("route") = mode
    record("error") = ""
    json = "{""id"":""" & id & """,""op"":""" & op & """,""sheet"":""" & JsonEscape(sheetName) & _
           """,""address"":""" & address & """,""hwnd"":" & Application.Hwnd & _
           ",""pid"":" & GetCurrentProcessId() & ",""workbook"":""" & JsonEscape(target.Name) & _
           """,""full_name"":""" & JsonEscape(target.FullName) & """}"
    record("json") = json
    requests.Add id, record

    Select Case mode
        Case "worker", "auto"
            If SendToWorker(json, record) Then
                record("route") = "worker"
            ElseIf mode = "auto" Then
                SendOneShot id, json
                record("route") = "oneshot (worker unavailable)"
            Else
                record("status") = "failed"
                record("error") = "the Python worker could not be started"
            End If
        Case "oneshot"
            SendOneShot id, json
    End Select
    ScheduleWatchdog
    Toolkit_Request = id
    Exit Function
Failed:
    Toolkit_Request = "ERROR " & Err.Number & ": " & Err.Description
End Function

Private Sub SendOneShot(ByVal id As String, ByVal json As String)
    Dim path As String, fileNo As Integer
    path = StateDir() & "\request_" & id & ".json"
    fileNo = FreeFile
    Open path For Output As #fileNo
    Print #fileNo, json
    Close #fileNo
    Shell """" & EnginePath() & """ oneshot """ & path & """", vbHide
End Sub

Private Function TryWritePipe(ByVal json As String) As Boolean
    Dim fileNo As Integer
    On Error GoTo Failed
    fileNo = FreeFile
    Open PipeName() For Binary Access Write As #fileNo
    Put #fileNo, , json & vbLf
    Close #fileNo
    TryWritePipe = True
    Exit Function
Failed:
    TryWritePipe = False
End Function

Private Function SendToWorker(ByVal json As String, ByVal record As Object) As Boolean
    Dim deadline As Double
    If TryWritePipe(json) Then
        SendToWorker = True
        Exit Function
    End If
    If Dir(EnginePath()) = "" Then Exit Function
    If Not WorkerAlive() Then
        Shell """" & EnginePath() & """ worker """ & StateDir() & """", vbHide
        record("worker_started") = True
    End If
    deadline = Seconds() + 10
    Do While Seconds() < deadline
        Sleep 20
        If TryWritePipe(json) Then
            SendToWorker = True
            Exit Function
        End If
    Loop
End Function

Private Function WorkerAlive() As Boolean
    Dim fileNo As Integer, text As String, handle As LongPtr, code As Long
    On Error GoTo NotAlive
    If Dir(StateDir() & "\worker.pid") = "" Then Exit Function
    fileNo = FreeFile
    Open StateDir() & "\worker.pid" For Input As #fileNo
    Line Input #fileNo, text
    Close #fileNo
    handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, CLng(text))
    If handle = 0 Then Exit Function
    GetExitCodeProcess handle, code
    CloseHandle handle
    WorkerAlive = (code = STILL_ACTIVE)
    Exit Function
NotAlive:
    WorkerAlive = False
End Function

' ---------- Result callback ----------

Public Function Toolkit_OnResult(ByVal id As String, ByVal pid As Long, ByVal workbookName As String, _
                                 ByVal fullName As String, ByVal status As String, ByVal errorText As String, _
                                 ByVal probe As Variant, ByVal count As Long, ByVal findings As Variant) As String
    Dim record As Object, i As Long, kinds As Object, key As Variant, summary As String, c As Long
    If requests Is Nothing Then Set requests = CreateObject("Scripting.Dictionary")
    If Not requests.Exists(id) Then
        Toolkit_OnResult = "rejected: unknown request"
        Exit Function
    End If
    Set record = requests(id)
    If record("status") <> "pending" Then
        Toolkit_OnResult = "rejected: request is " & record("status")
        Exit Function
    End If
    If pid <> GetCurrentProcessId() Or workbookName <> record("wb") Or fullName <> record("full") Then
        Toolkit_OnResult = "rejected: identity mismatch"
        Exit Function
    End If
    record("ms") = (Seconds() - record("start")) * 1000
    record("status") = status
    record("error") = errorText
    record("probe") = probe
    record("count") = count
    If IsArray(findings) Then
        Set kinds = CreateObject("Scripting.Dictionary")
        c = LBound(findings, 2)
        For i = LBound(findings, 1) To UBound(findings, 1)
            kinds(findings(i, c + 2)) = kinds(findings(i, c + 2)) + 1
        Next i
        For Each key In kinds.Keys
            summary = summary & key & ": " & kinds(key) & vbLf
        Next key
    End If
    record("summary") = summary
    If Not quietMode Then
        lastShownId = id
        Application.OnTime Now, "Toolkit_ShowLastResult"   ' never show UI inside the callback
    End If
    Toolkit_OnResult = "accepted"
End Function

Public Sub Toolkit_ShowLastResult()
    Dim record As Object
    If Not requests.Exists(lastShownId) Then Exit Sub
    Set record = requests(lastShownId)
    If record("status") = "ok" Then
        MsgBox record("count") & " cells classified in " & record("wb") & " (" & Format(record("ms"), "0") & " ms, " & _
               record("route") & ")" & vbLf & vbLf & record("summary"), vbInformation, "Excel Toolkit " & VERSION_TEXT
    Else
        MsgBox "The audit did not complete: " & record("error"), vbExclamation, "Excel Toolkit " & VERSION_TEXT
    End If
End Sub

Public Function Toolkit_GetResult(ByVal id As String) As String
    Dim record As Object
    If requests Is Nothing Then Set requests = CreateObject("Scripting.Dictionary")
    If Not requests.Exists(id) Then
        Toolkit_GetResult = "unknown"
        Exit Function
    End If
    Set record = requests(id)
    Toolkit_GetResult = record("status") & "|" & record("count") & "|" & Format(record("ms"), "0.000") & "|" & _
                        record("probe") & "|" & record("error") & "|" & record("route") & "|" & record("wb")
End Function

' ---------- Watchdog: falls back to one-shot if the worker dies ----------

Private Sub ScheduleWatchdog()
    If watchdogScheduled Then Exit Sub
    watchdogScheduled = True
    Application.OnTime Now + TimeSerial(0, 0, 1), "Toolkit_Watchdog"
End Sub

Public Sub Toolkit_Watchdog()
    Dim key As Variant, record As Object, anyPending As Boolean
    watchdogScheduled = False
    For Each key In requests.Keys
        Set record = requests(key)
        If record("status") = "pending" Then
            If record("route") = "worker" And Not WorkerAlive() Then
                record("route") = "oneshot (worker stopped)"
                SendOneShot CStr(key), record("json")
                anyPending = True
            ElseIf record("route") = "worker" And Seconds() - record("start") > WORKER_SILENT_SECONDS Then
                record("route") = "oneshot (worker silent)"
                SendOneShot CStr(key), record("json")
                anyPending = True
            ElseIf Seconds() - record("start") > REQUEST_TIMEOUT_SECONDS Then
                record("status") = "failed"
                record("error") = "no answer within " & REQUEST_TIMEOUT_SECONDS & " s"
            Else
                anyPending = True
            End If
        End If
    Next key
    If anyPending Then ScheduleWatchdog
End Sub
