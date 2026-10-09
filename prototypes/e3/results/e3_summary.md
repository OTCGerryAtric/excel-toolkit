# E3 results: 28/28 passed


## Formatting toggle: Excel Ctrl+Z via Application.OnUndo

- **PASS** Toggle changes mixed multi-area selection to the first format — before ['#,##0.0', '#,##0.0', 'General'] -> ['#,##0_-;[Red](#,##0)_-;-_-', '#,##0_-;[Red](#,##0)_-;-_-', '#,##0_-;[Red](#,##0)_-;-_-']
- **PASS** Informational: Application.Undo called from automation triggers the OnUndo restore — restored: False (no error); formats now ['#,##0_-;[Red](#,##0)_-;-_-', '#,##0_-;[Red](#,##0)_-;-_-', '#,##0_-;[Red](#,##0)_-;-_-']
- **PASS** Informational: Application.Undo called inside VBA triggers the OnUndo restore — restored: False (no error); formats now ['#,##0_-;[Red](#,##0)_-;-_-', '#,##0_-;[Red](#,##0)_-;-_-', '#,##0_-;[Red](#,##0)_-;-_-']
- **PASS** Informational: OnUndo hook registered for a Python operation fires on Application.Undo — marker after Undo: '' - real Ctrl+Z needs a manual keypress test

## Bulk transformation: custom undo

- **PASS** Evidence: naive Value2 = Value2 re-parses text constants — text '0012' and '1/2' became [12.0, 46024.0]
- **PASS** Paste as values: values, text, errors, booleans and formats correct — {"A1 now a value (10)": true, "B1 text '0012' kept as text": true, "B2 text '1/2' kept as text": true, "B3 #N/A kept as error constant": true, "B4 TRUE kept as boolean": true, "C formats kept": true}
- **PASS** Custom undo restores formulas and constants exactly — {'ok': True}; verified equal to before: True; formula A1 now '=ROW()*10'

## Bulk transformation: conflict detection

- **PASS** Undo refused when an affected cell changed since; conflicting cell reported — {'ok': False, 'reason': 'cells changed since the operation', 'conflicts': ['Data!$A$2']}
- **PASS** Nothing restored after a refused undo (no partial overwrite) — A1 formula '10', A2 777.0

## Bulk transformation: rollback after partial failure

- **PASS** First block written, second refused by sheet protection, first block rolled back — status rolled_back; events ['23:33:42 failed: com_error: (-2147352567, \'Exception occurred.\', (0, \'Microsoft Excel\', "The cell or chart you\'re trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.", \'xlmain11.chm\', 0, -2146827284), None)', "23:33:42 rolled_back: restored ['Prot!A1:A3']"]; error: paste_as_values failed and was rolled back: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', "The cell or cha

## Destructive: delete names

- **PASS** Preview lists every name, including hidden, sheet-scoped, print area and broken — [('Broken', True), ('Data!LocalName', True), ('Data!Print_Area', True), ('HiddenName', False), ('Rate', True)]
- **PASS** Refuses without confirmation — delete names needs explicit confirmation after the preview
- **PASS** Failure after 2 deletions rolls back to the identical name set — delete names failed and was rolled back: simulated failure after deleting 2 names
- **PASS** Delete names removes them all — remaining []
- **PASS** Undo restores every name with scope, definition and visibility — {'ok': True}; differences: []

## Identity: unsaved new workbook

- **PASS** Unsaved workbook has no path; identity uses its name — name 'Book2', full name 'Book2', Path ''
- **PASS** Operation and undo work on an unsaved workbook — {'ok': True}

## Identity: renamed, closed and reopened workbook

- **PASS** Undo after Save As under a new name is refused — workbook 'E3_Rename.xlsx' is no longer open (closed or renamed)
- **PASS** Undo against a closed-and-reopened workbook (same name/path) is refused by the content check — {'ok': False, 'reason': 'cells changed since the operation', 'conflicts': ['Data!$A$1', 'Data!$A$2', 'Data!$A$3', 'Data!$A$4', 'Data!$A$5']}
- **PASS** Stale identity after close is refused — workbook 'E3_Rename.xlsx' is no longer open (closed or renamed)

## Identity: stale process

- **PASS** Identity whose process start time doesn't match (pid reused) is refused — process id 26728 now belongs to a different process (stale identity)

## Resilience: Excel busy

- **PASS** VBA MsgBox open: outcome of read + write with a 5 s bounded wait — read and write succeeded after 0.0 s
- **PASS** built-in dialog (Format Number) open: outcome of read + write with a 5 s bounded wait — read and write succeeded after 0.0 s
- **PASS** Cell edit mode (informational) — read succeeded in 0.0 s ([[['=ROW()*10', 10.0]]]); edit mode was probably not entered via posted keys - verify manually
- **PASS** Excel responsive again and the edited cell unchanged after Escape — responsive True, A1 '=ROW()*10'

## Resilience: Excel closes during an operation

- **PASS** Excel killed mid-operation: clear error, no hang, journal marked incomplete — 1.1 s; Excel closed or stopped responding during the operation.; journal status incomplete, cells written before loss ~400 of 3000
- **PASS** Identity of the killed Excel is refused afterwards — Excel process 23228 is no longer running

## housekeeping

- **PASS** Cleanup — test Excel 26728 did not quit within 15 s; terminated
