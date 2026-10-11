# M3a validation — number-format commands and shortcut precedence

Status: **build, VBA compilation and automated live checks passed (19/19, 2026-10-11/12). Manual checklist and shortcut-precedence test NOT yet run** (awaiting review).
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

## Results — automated live run

Separate hidden Excel processes for build, compile check and scripted checks; nothing installed.
PERSONAL.XLSB SHA-256 unchanged; 0 Defender detections; SentinelOne logged one routine burst of 14
event-143 entries, none naming Excel, Python or the add-in; no Excel process left.

**VBA compilation:** clean at the first build (VBE Debug > Compile; no dialogs).

**First run 13/14 — investigated before re-running:**
1. *Add-in defect.* A mixed **whole column** failed with error 1004 while capturing (the command
   refused safely; nothing changed): the recursive split took the second half with `Offset(half)`,
   which for a whole column runs past the last row of the sheet. Fixed to `Rows(half + 1).Resize(...)`
   (and the column equivalent). The performance cases are now pass/fail checks, not just timings.
2. *Harness defect.* `Protect(AllowFormattingCells=True)` passed by keyword through late-bound COM is
   silently ignored (confirmed: property stays False; positional form sets True). The add-in had
   correctly refused a sheet that really did not allow formatting. Harness now protects positionally
   and asserts the fixture.

**Final: 19/19** (`addin/live/results/m3a_live_20261011_235821.json`):

| Area | Result |
|---|---|
| Legacy equivalence | Toolkit matches the legacy macro **cell by cell** for all four commands × three selections (single cell, block, multi-area) × every start state (each cycle format, General, `0.00`, mixed) — 72 cases — plus 5-press sequences |
| Exact undo | Every one of those changes undone exactly, including mixed formats; 500 randomly mixed cells restored exactly |
| Values | Untouched by format commands |
| Conflict / identity | Undo refused after a format edit (nothing restored) and after a sheet rename |
| Protection | Refused when formatting is not allowed (nothing changed); applied and undone when it is |
| Failures | Forced failure after the first area rolled back exactly; forced restore failure reported as incomplete, record unusable |
| Safeguard | 60,000 alternating formats (> 50,000 blocks) refused, nothing changed |
| Shortcut toggle | Switches on and off without error (key behaviour is the manual test) |

**Performance** (Number format; ms; apply = click to undo registered):

| Case | Blocks | Apply | Undo |
|---|---|---|---|
| Uniform 10,000 cells | 1 | **3.6** | 15 |
| Mixed 10,000 (alternating rows) | 10,000 | 202 | 1,279 |
| Mixed 30,000 (3 uniform columns) | 3 | 5.8 | 18 |
| Whole column, uniform | 1 | 3.9 | 13 |
| Whole column, 10 exceptions | 175 | 12.5 | 221 |

Uniform selections meet the 50 ms aspiration by a wide margin. Worst-case mixed formatting (every row
different) is usable to apply (0.2 s) but slow to undo (1.3 s, one write per block) — a candidate for
the separate optimisation track, not a correctness issue.
