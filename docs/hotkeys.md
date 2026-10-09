# Hotkey register

Default shortcuts for the add-in. Colleagues may remap these in their own configuration.

## Current VBA shortcuts (PERSONAL.XLSB, via Macro Options)

Read from `PERSONAL.XLSB` on 2026-10-09. All are Ctrl+Shift+letter.

| Key | Macro | Status |
|---|---|---|
| Ctrl+Shift+I | Number_Format_Toggle | VBA |
| Ctrl+Shift+O | Currency_Format_Toggle | VBA |
| Ctrl+Shift+P | Percentage_Format_Toggle | VBA |
| Ctrl+Shift+M | Month_Date_Format | VBA |
| Ctrl+Shift+R | Row_Height_Toggle | VBA |
| Ctrl+Shift+C | Column_Width_Toggle | VBA |
| Ctrl+Shift+Q | Yellow_Cell_Toggle | VBA |
| Ctrl+Shift+E | Green_Cell_Toggle | VBA |
| Ctrl+Shift+W | Light_Blue_Cell_Toggle | VBA |
| Ctrl+Shift+Z | Grey_Cell_Toggle | VBA |
| Ctrl+Shift+V | Paste_As_Values | VBA |
| Ctrl+Shift+K | Best_Fit_Column_Width | VBA |
| Ctrl+Shift+G | Grey_Ouline | VBA |
| Ctrl+Shift+F | Clear_Sheet_Formula_References | VBA |
| Ctrl+Shift+N | Delete_Names | VBA |
| Ctrl+Shift+X | ReportExternalLinks (Module2) | VBA |

## Add-in shortcuts

| Key | Command | Notes |
|---|---|---|
| Ctrl+Shift+H | Find constants | Moved from VBA (was clashing with Column_Width_Toggle on Ctrl+Shift+C). Decided 2026-10-09. |

## Decisions

- 2026-10-09: Find Constants moves to the add-in on Ctrl+Shift+H. Its Ctrl+Shift+C shortcut in PERSONAL.XLSB should be cleared so Column_Width_Toggle owns Ctrl+Shift+C.
