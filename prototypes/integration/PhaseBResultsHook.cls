VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhaseBResultsHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Phase B double-click navigation (OFFLINE-PREPARED - reviewed, NOT yet executed).
'
' An Application-level event sink so a double-click on a row of the Phase B results sheet
' navigates to the suspect cell in the ORIGINAL audited workbook. Using Application events
' avoids injecting code into the dynamically created results workbook, so there is NO use of
' the VBA project object model and NO trust/security change. Owned by PhaseBBridge (set in
' PhaseB_Configure: Set hook = New PhaseBResultsHook: Set hook.App = Application).
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Only the Phase B results sheet reacts; every other double-click behaves normally.
    If PhaseB_HandleDoubleClick(Sh, Target) Then Cancel = True   ' cancel edit mode; we navigated
End Sub
