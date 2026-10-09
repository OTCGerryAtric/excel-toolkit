Attribute VB_Name = "E1Test"
' Phase 0 / E1 prototype: number-format cycle in an .xlam on a test hotkey,
' plus an equivalence and timing harness against PERSONAL.XLSB!Number_Format_Toggle.
' Throwaway code - not production.
Option Explicit

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long

Private Const TEST_KEY As String = "^+J"
Private Const LEGACY_MACRO As String = "PERSONAL.XLSB!Number_Format_Toggle"
Private Const NEW_MACRO As String = "E1Test.xlam!E1_Number_Format_Toggle"
Private Const TIMED_RUNS As Long = 200
Private Const WARMUP_RUNS As Long = 10

' ---------- Hotkey binding ----------

Public Sub Auto_Open()
    E1_Init
End Sub

Public Sub Auto_Close()
    Application.OnKey TEST_KEY
End Sub

Public Sub E1_Init()
    Application.OnKey TEST_KEY, "E1_Number_Format_Toggle"
End Sub

' ---------- The command ----------

Private Function NumberFormatCycle() As Variant
    NumberFormatCycle = Array( _
        "#,##0_-;[Red](#,##0)_-;-_-", _
        "#,##0.0_-;[Red](#,##0.0)_-;-_-", _
        "#,##0.00_-;[Red](#,##0.00)_-;-_-")
End Function

Public Sub E1_Number_Format_Toggle()
    If TypeName(Selection) <> "Range" Then Exit Sub
    CycleNumberFormat Selection, NumberFormatCycle()
End Sub

' Same semantics as the VBA original: if the whole selection has a format in the
' cycle, move to the next one; otherwise (including mixed formats) start at the first.
Private Sub CycleNumberFormat(ByVal target As Range, ByVal formats As Variant)
    Dim current As Variant
    Dim i As Long
    current = target.NumberFormat
    If Not IsNull(current) Then
        For i = LBound(formats) To UBound(formats) - 1
            If current = formats(i) Then
                target.NumberFormat = formats(i + 1)
                Exit Sub
            End If
        Next i
    End If
    target.NumberFormat = formats(LBound(formats))
End Sub

' ---------- Harness ----------

Private Function Seconds() As Double
    Dim counter As Currency, frequency As Currency
    QueryPerformanceCounter counter
    QueryPerformanceFrequency frequency
    Seconds = counter / frequency
End Function

Private Function TestSelections() As Variant
    TestSelections = Array("B2", "B2:D10", "B2:B5,D7:E9")
End Function

Private Function StartStates() As Variant
    Dim formats As Variant
    formats = NumberFormatCycle()
    StartStates = Array(formats(0), formats(1), formats(2), "General", "0.00", "<mixed>")
End Function

Private Sub ApplyStartState(ByVal target As Range, ByVal state As String)
    If state = "<mixed>" Then
        target.NumberFormat = NumberFormatCycle()(0)
        target.Areas(1).Cells(1, 1).NumberFormat = "General"
    Else
        target.NumberFormat = state
    End If
End Sub

Private Function FormatOf(ByVal target As Range) As String
    If IsNull(target.NumberFormat) Then
        FormatOf = "<mixed>"
    Else
        FormatOf = target.NumberFormat
    End If
End Function

Private Function Csv(ByVal text As String) As String
    Csv = """" & Replace(text, """", """""") & """"
End Function

' Runs on a new scratch workbook, closed without saving. Writes two CSVs to outDir.
Public Function E1_RunAll(ByVal outDir As String) As String
    Dim scratch As Workbook
    Dim sheet As Worksheet
    Set scratch = Workbooks.Add
    Set sheet = scratch.Worksheets(1)
    sheet.Range("A1:H50").Value = 1234.567
    sheet.Activate
    WriteEquivalence sheet, outDir & "\e1_equivalence.csv"
    WriteTimings sheet, outDir & "\e1_timings.csv"
    scratch.Close SaveChanges:=False
    E1_RunAll = "done"
End Function

Private Sub WriteEquivalence(ByVal sheet As Worksheet, ByVal path As String)
    Dim fileNo As Integer
    Dim sel As Variant, state As Variant
    Dim target As Range
    Dim legacyResult As String, newResult As String
    Dim legacySeq As String, newSeq As String
    Dim press As Long

    fileNo = FreeFile
    Open path For Output As #fileNo
    Print #fileNo, "test,selection,start_state,legacy_result,new_result,match"

    For Each sel In TestSelections()
        Set target = sheet.Range(sel)
        For Each state In StartStates()
            ApplyStartState target, CStr(state)
            target.Select
            Application.Run LEGACY_MACRO
            legacyResult = FormatOf(target)

            ApplyStartState target, CStr(state)
            target.Select
            Application.Run NEW_MACRO
            newResult = FormatOf(target)

            Print #fileNo, "single_press," & Csv(CStr(sel)) & "," & Csv(CStr(state)) & "," & _
                Csv(legacyResult) & "," & Csv(newResult) & "," & (legacyResult = newResult)
        Next state

        ' Five presses from General: checks the full cycle order and wrap-around.
        legacySeq = "": newSeq = ""
        target.NumberFormat = "General": target.Select
        For press = 1 To 5
            Application.Run LEGACY_MACRO
            legacySeq = legacySeq & IIf(press > 1, " > ", "") & FormatOf(target)
        Next press
        target.NumberFormat = "General": target.Select
        For press = 1 To 5
            Application.Run NEW_MACRO
            newSeq = newSeq & IIf(press > 1, " > ", "") & FormatOf(target)
        Next press
        Print #fileNo, "cycle_5_presses," & Csv(CStr(sel)) & "," & Csv("General") & "," & _
            Csv(legacySeq) & "," & Csv(newSeq) & "," & (legacySeq = newSeq)
    Next sel

    Close #fileNo
End Sub

Private Sub WriteTimings(ByVal sheet As Worksheet, ByVal path As String)
    Dim fileNo As Integer
    Dim sel As Variant, macroName As Variant
    Dim target As Range
    Dim run As Long
    Dim started As Double

    fileNo = FreeFile
    Open path For Output As #fileNo
    Print #fileNo, "macro,selection,run,ms"

    For Each sel In TestSelections()
        Set target = sheet.Range(sel)
        target.Select
        For Each macroName In Array(LEGACY_MACRO, NEW_MACRO)
            For run = 1 To WARMUP_RUNS
                Application.Run macroName
            Next run
            For run = 1 To TIMED_RUNS
                started = Seconds()
                Application.Run macroName
                Print #fileNo, Csv(CStr(macroName)) & "," & Csv(CStr(sel)) & "," & run & "," & _
                    Format((Seconds() - started) * 1000, "0.0000")
            Next run
        Next macroName
    Next sel

    Close #fileNo
End Sub
