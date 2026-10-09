Attribute VB_Name = "E4Bridge"
' Phase 0 / E4 prototype: send an audit request to Python (one-shot or persistent worker)
' and receive a structured result back through a callback. Throwaway code - not production.
Option Explicit

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal access As Long, ByVal inherit As Long, ByVal pid As Long) As LongPtr
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal handle As LongPtr, ByRef code As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const STILL_ACTIVE As Long = 259
Private Const REQUEST_TIMEOUT_SECONDS As Double = 60

Private pythonExe As String
Private scriptsDir As String
Private stateDir As String
Private requests As Object      ' request id -> Scripting.Dictionary
Private rejectedCount As Long
Private requestCounter As Long
Private watchdogScheduled As Boolean

' ---------- Setup ----------

Public Sub E4_Configure(ByVal python As String, ByVal scripts As String, ByVal state As String)
    pythonExe = python
    scriptsDir = scripts
    stateDir = state
    Set requests = CreateObject("Scripting.Dictionary")
    rejectedCount = 0
End Sub

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
    PipeName = "\\.\pipe\excel_toolkit_e4_" & Environ("USERNAME")
End Function

' ---------- Request ----------

' Returns the request id at once; the result arrives later through E4_OnResult.
Public Function E4_Request(ByVal mode As String, ByVal op As String, ByVal sheetName As String, ByVal address As String, _
                           Optional ByVal workbookName As String = "") As String
    ' The workbook is named explicitly: ActiveWorkbook can lag behind a programmatic Activate.
    Dim id As String, json As String, path As String, fileNo As Integer
    Dim record As Object, target As Workbook
    On Error GoTo Failed

    If workbookName = "" Then Set target = ActiveWorkbook Else Set target = Workbooks(workbookName)
    requestCounter = requestCounter + 1
    id = GetCurrentProcessId() & "-" & requestCounter & "-" & Format(Now, "hhnnss")

    Set record = CreateObject("Scripting.Dictionary")
    record("status") = "pending"
    record("mode") = mode
    record("start") = Seconds()
    record("wb") = target.Name
    record("full") = target.FullName
    record("error") = ""
    requests.Add id, record

    json = "{""id"":""" & id & """,""op"":""" & op & """,""sheet"":""" & JsonEscape(sheetName) & _
           """,""address"":""" & address & """,""hwnd"":" & Application.Hwnd & _
           ",""pid"":" & GetCurrentProcessId() & ",""workbook"":""" & JsonEscape(target.Name) & _
           """,""full_name"":""" & JsonEscape(target.FullName) & """}"

    If mode = "oneshot" Then
        path = stateDir & "\request_" & id & ".json"
        fileNo = FreeFile
        Open path For Output As #fileNo
        Print #fileNo, json
        Close #fileNo
        Shell """" & pythonExe & """ """ & scriptsDir & "\e4_oneshot.py"" """ & path & """", vbHide
    ElseIf mode = "worker" Then
        If Not SendToWorker(json, record) Then
            record("status") = "failed"
            record("error") = "the Python worker could not be started"
        End If
    End If  ' any other mode: register the request only (used to test result validation)

    ScheduleWatchdog
    E4_Request = id
    Exit Function
Failed:
    E4_Request = "ERROR " & Err.Number & ": " & Err.Description
End Function

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
    ' No worker listening (or between connections): start one and retry for up to 10 s.
    If Not WorkerAlive() Then
        Shell """" & pythonExe & """ """ & scriptsDir & "\e4_worker.py"" """ & stateDir & """", vbHide
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
    SendToWorker = False
End Function

Private Function WorkerAlive() As Boolean
    Dim fileNo As Integer, text As String, handle As LongPtr, code As Long
    On Error GoTo NotAlive
    If Dir(stateDir & "\worker.pid") = "" Then Exit Function
    fileNo = FreeFile
    Open stateDir & "\worker.pid" For Input As #fileNo
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

' ---------- Result callback (called by Python) ----------

Public Function E4_OnResult(ByVal id As String, ByVal pid As Long, ByVal workbookName As String, _
                            ByVal fullName As String, ByVal status As String, ByVal errorText As String, _
                            ByVal probe As Variant, ByVal count As Long, ByVal findings As Variant) As String
    Dim record As Object, first As Long, last As Long
    If requests Is Nothing Then Set requests = CreateObject("Scripting.Dictionary")
    If Not requests.Exists(id) Then
        rejectedCount = rejectedCount + 1
        E4_OnResult = "rejected: unknown request"
        Exit Function
    End If
    Set record = requests(id)
    If record("status") <> "pending" Then
        rejectedCount = rejectedCount + 1
        E4_OnResult = "rejected: request is " & record("status")
        Exit Function
    End If
    If pid <> GetCurrentProcessId() Or workbookName <> record("wb") Or fullName <> record("full") Then
        rejectedCount = rejectedCount + 1
        E4_OnResult = "rejected: identity mismatch"
        Exit Function
    End If

    record("ms") = (Seconds() - record("start")) * 1000
    record("status") = status
    record("error") = errorText
    record("probe") = probe
    record("count") = count
    record("rows") = 0
    If IsArray(findings) Then
        first = LBound(findings, 1)
        last = UBound(findings, 1)
        record("rows") = last - first + 1
        record("first") = findings(first, LBound(findings, 2)) & "!" & findings(first, LBound(findings, 2) + 1) & "=" & findings(first, LBound(findings, 2) + 2)
        record("last") = findings(last, LBound(findings, 2)) & "!" & findings(last, LBound(findings, 2) + 1) & "=" & findings(last, LBound(findings, 2) + 2)
    End If
    E4_OnResult = "accepted"
End Function

Public Function E4_GetResult(ByVal id As String) As String
    Dim record As Object
    If Not requests.Exists(id) Then
        E4_GetResult = "unknown"
        Exit Function
    End If
    Set record = requests(id)
    E4_GetResult = record("status") & "|" & record("count") & "|" & record("rows") & "|" & _
                   Format(record("ms"), "0.000") & "|" & record("first") & "|" & record("last") & "|" & _
                   record("probe") & "|" & record("error") & "|" & record("worker_started") & "|" & record("wb")
End Function

Public Function E4_GetRejectedCount() As Long
    E4_GetRejectedCount = rejectedCount
End Function

' ---------- Watchdog ----------

Private Sub ScheduleWatchdog()
    If watchdogScheduled Then Exit Sub
    watchdogScheduled = True
    Application.OnTime Now + TimeSerial(0, 0, 1), "E4_Watchdog"
End Sub

Public Sub E4_Watchdog()
    Dim key As Variant, record As Object, anyPending As Boolean
    watchdogScheduled = False
    For Each key In requests.Keys
        Set record = requests(key)
        If record("status") = "pending" Then
            If record("mode") = "worker" And Not WorkerAlive() Then
                record("status") = "failed"
                record("error") = "the Python worker stopped before answering"
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

' ---------- In-process write benchmark (disposable workbook only) ----------

Public Function E4_BenchFormat(ByVal workbookName As String, ByVal sheetName As String, _
                               ByVal address As String, ByVal numberFormat As String) As Double
    Dim started As Double
    started = Seconds()
    Workbooks(workbookName).Worksheets(sheetName).Range(address).NumberFormat = numberFormat
    E4_BenchFormat = (Seconds() - started) * 1000
End Function
