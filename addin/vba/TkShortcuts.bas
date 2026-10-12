Attribute VB_Name = "TkShortcuts"
' Excel Toolkit (M3a): optional keyboard shortcuts through Application.OnKey - OFF by default.
'
' Session only; PERSONAL.XLSB is never edited. Switching off calls Application.OnKey <key> with no
' procedure, which returns each key to its normal behaviour. Whether that - and closing the add-in,
' and restarting Excel - restores the PERSONAL.XLSB macro shortcuts is NOT assumed: it is proven
' live (M3a shortcut-precedence test) before this toggle is relied on.
Option Explicit

Private enabled As Boolean

' Array(OnKey code, procedure) for each migrated command.
Private Function Bindings() As Variant
    Bindings = Array(Array("^+I", "Tk_KeyNumberFormat"), Array("^+O", "Tk_KeyCurrencyFormat"), _
                     Array("^+P", "Tk_KeyPercentFormat"), Array("^+M", "Tk_KeyDateFormat"))
End Function

Public Function Tk_ShortcutsEnabled() As Boolean
    Tk_ShortcutsEnabled = enabled
End Function

Public Sub Tk_SetShortcuts(ByVal turnOn As Boolean)
    Dim binding As Variant
    For Each binding In Bindings()
        If turnOn Then
            Application.OnKey binding(0), "'" & ThisWorkbook.Name & "'!" & binding(1)
        Else
            Application.OnKey binding(0)
        End If
    Next binding
    enabled = turnOn
    Tk_RefreshShortcutToggle
End Sub

' Excel runs Auto_Close when the add-in workbook closes (including when Excel exits): never leave
' a key pointing at a closed add-in.
Public Sub Auto_Close()
    If enabled Then Tk_SetShortcuts False
End Sub

' ---------- Key handlers ----------

Public Sub Tk_KeyNumberFormat()
    Tk_RunFormat "number"
End Sub

Public Sub Tk_KeyCurrencyFormat()
    Tk_RunFormat "currency"
End Sub

Public Sub Tk_KeyPercentFormat()
    Tk_RunFormat "percent"
End Sub

Public Sub Tk_KeyDateFormat()
    Tk_RunFormat "date"
End Sub
