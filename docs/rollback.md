# Rollback instructions

How to return Excel to its pre-project state. Baseline captured 2026-10-09, before PyXLL was installed.

## What was captured (`baseline/`)

| File | Contents |
|---|---|
| `PERSONAL.XLSB.bak` | Copy of `%APPDATA%\Microsoft\Excel\XLSTART\PERSONAL.XLSB` (32,763 bytes, last modified 2026-06-15). Not committed to git. |
| `vba/PERSONAL_XLSB_olevba.txt` | Extracted VBA source of PERSONAL.XLSB (Module1, Module2, ThisWorkbook, Sheet1), including hotkey attributes. |
| `registry/excel_options.reg` | `HKCU\Software\Microsoft\Office\16.0\Excel\Options` — includes the `OPEN` key loading Arixcel Explorer. |
| `registry/excel_com_addins.reg` | `HKCU\Software\Microsoft\Office\Excel\Addins` — COM add-in registrations. |

`HKCU\...\Excel\Add-in Manager` did not exist at capture time, so it was not exported.

Excel add-ins loaded at baseline: Arixcel Explorer only (`OPEN = /R "...\Arixcel Explorer\ArixcelExplorer.xlam"`).

## Removing the project add-in

PyXLL was ruled out on 2026-10-09 before it registered anything with Excel (verified: no PyXLL registry entries; only Arixcel in the `OPEN` keys).

To remove the project's `.xlam` once it exists:
1. Excel → File → Options → Add-ins → Manage: Excel Add-ins → Go → untick it.
2. Close Excel and delete the `.xlam` from its install folder.
3. Restart Excel and check Arixcel is still listed.

## If Excel settings are damaged

1. Close Excel.
2. Double-click `baseline/registry/excel_options.reg` (and `excel_com_addins.reg` if COM add-ins are affected) to restore the keys.
   Note: importing a .reg file merges keys; it does not delete values added since. If PyXLL left an extra `OPENn` value, delete it in regedit under `HKCU\Software\Microsoft\Office\16.0\Excel\Options`.
3. Restart Excel.

## If PERSONAL.XLSB is lost or damaged

1. Close Excel.
2. Copy `baseline/PERSONAL.XLSB.bak` to `%APPDATA%\Microsoft\Excel\XLSTART\PERSONAL.XLSB`.
3. Restart Excel. Hotkeys come back with the file (they are stored as macro attributes).
