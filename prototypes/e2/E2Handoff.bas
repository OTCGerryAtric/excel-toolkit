Attribute VB_Name = "E2Handoff"
' Phase 0 / E2 prototype: hand the exact Excel instance and workbook identity to Python.
' Throwaway code - not production.
Option Explicit

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

' Launches the reader with this instance's main window handle, process id and the
' active workbook's name. Returns immediately; the reader writes its JSON to outFile.
Public Function E2_Handoff(ByVal pythonExe As String, ByVal script As String, _
                           ByVal backend As String, ByVal outFile As String) As String
    Dim command As String
    command = """" & pythonExe & """ """ & script & """" & _
              " --backend " & backend & _
              " --hwnd " & Application.Hwnd & _
              " --pid " & GetCurrentProcessId() & _
              " --workbook """ & ActiveWorkbook.Name & """" & _
              " --out """ & outFile & """"
    Shell command, vbHide
    E2_Handoff = command
End Function
