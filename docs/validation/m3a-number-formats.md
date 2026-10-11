# M3a validation — number-format commands and shortcut precedence

Status: **implemented and offline-tested; live tests NOT yet run** (need separate confirmation).
Plan: `docs/product-roadmap.md` (M3, Gate 1). Source: `addin/`.

## What M3a adds

| Part | File |
|---|---|
| Number (I), Currency £ (O), Percentage (P) cycles and Date toggle (M), legacy formats and rule | `addin/vba/TkFormat.bas` |
| Number-format undo record: uniform areas read once, mixed areas split recursively (columns, then rows) into uniform blocks; more than 50,000 blocks refused | `addin/vba/TkFormatSnapshot.cls` |
| Optional shortcuts via `Application.OnKey`, **off by default**, session only, switched off on `Auto_Close` | `addin/vba/TkShortcuts.bas` |
| Ribbon: Format ▸ Number formats menu; Settings ▸ Toolkit shortcuts toggle | `addin/customUI/customUI14.xml` |
| Undo record generalised (any snapshot with Description / CheckIdentity / FindConflict / Restore); shared identity check | `addin/vba/TkUndo.bas`; `TkCellSnapshot.Restore` added |

The legacy rule is reproduced exactly: from a cycle format other than the last, go to the next; from
anything else (the last format, any other format, mixed) go to the first. Only `NumberFormat` changes;
the decision uses the selection's `NumberFormat`, as the legacy macros do.

## Offline evidence

`tests/addin/test_addin_offline.py` (59 tests): the four toolkit cycles reproduce the legacy macros'
transitions and default, parsed from `baseline/vba` (a deliberate mutation is detected); both snapshot
classes provide the undo-record interface that `TkUndo` calls late-bound; shortcuts are off by default,
never switched on by code, bound to Ctrl+Shift+I/O/P/M with zero-argument handlers, and switched off on
`Auto_Close`; ribbon callbacks (including the toggle's `getPressed` / `onAction(control, pressed)`)
resolve with the right signatures; registry, ribbon and dispatch agree; M2's hygiene checks (ASCII, CRLF,
names, declarations, no clipboard or process launch). Not provable offline: compilation and behaviour.

## Live tests (pending confirmation)

1. **Build** in a separate hidden Excel; **VBA Compile** check in a disposable instance.
2. **Scripted** (`addin/live/m3a_live_check.py --perf`, disposable Excel; legacy macros imported from
   `baseline/vba` into the disposable workbook): each command side by side with its legacy macro on
   single cell / block / multi-area selections from every cycle format, General, `0.00` and mixed
   formats, plus 5-press sequences; exact undo including mixed formats; values untouched; conflict and
   identity refusal; protected sheet (refused unless formatting is allowed); forced apply and restore
   failures; too-many-blocks refusal; toggle on/off; timings (uniform 10,000; alternating 10,000; three
   uniform columns; whole column uniform; whole column with 10 exceptions) with undo timed separately.
3. **Manual** (owner's Excel, File → Open the development `.xlam`, new blank workbook):

| # | Step | Expected | Result |
|---|---|---|---|
| 1 | Ribbon | Format ▸ Number formats menu with four commands and tooltips; Settings ▸ Toolkit shortcuts (not pressed) | |
| 2 | Put numbers in A1:A3 with **different** formats; select; Number formats ▸ Number | All become `#,##0` style; Undo Toolkit Action enabled | |
| 3 | Press **Ctrl+Z** | Excel offers "Undo Number Format"; each cell gets back its **own** original format | |
| 4 | Use Currency, Percentage and Date once each; Ctrl+Z after each | Each applies and undoes | |

**Shortcut-precedence test** (the toolkit is identifiable: it enables Undo Toolkit Action and offers
"Undo Number Format" on Ctrl+Z; the PERSONAL.XLSB macro does neither):

| # | Step | Expected | Result |
|---|---|---|---|
| S1 | Toggle **off** (default). Select a cell, press **Ctrl+Shift+I** | PERSONAL.XLSB macro runs (format changes; Undo Toolkit Action stays as it was; Ctrl+Z cannot undo it) | |
| S2 | Toggle **on**. Press Ctrl+Shift+I | Toolkit runs (Undo Toolkit Action enabled; Ctrl+Z offers "Undo Number Format") | |
| S3 | Toggle **off** again. Press Ctrl+Shift+I | **PERSONAL.XLSB macro again** | |
| S4 | Toggle on; close the add-in by closing Excel (**Don't Save**); reopen Excel *without* the add-in; press Ctrl+Shift+I | PERSONAL.XLSB macro | |
| S5 | Repeat S1–S3 for Ctrl+Shift+O, P, M (quick check) | Same pattern | |
| S6 | Ctrl+Shift+V with the toggle on and off | Always the PERSONAL.XLSB macro (the toolkit does not bind V in M3a) | |
| S7 | Close Excel, **Don't Save** everywhere | PERSONAL.XLSB unchanged (hash); no security events | |

If S3 or S4 does not give the PERSONAL.XLSB macro back, the toggle is not adopted and a safer mechanism
is proposed. Excel's own (native) Ctrl+Shift+V behaviour, without PERSONAL.XLSB, is recorded separately
in M3d.

## Results

_Not yet run._
