# M3a validation — number-format commands and shortcut precedence

Status: **build, VBA compilation and automated live checks passed (19/19, 2026-10-11/12), reviewed and
accepted at `eabaadf`. Manual checklist and shortcut-precedence test approved to run (2026-10-12); not
yet run.** Not approved for merge.
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
3. **Manual** (owner's normal Excel, so PERSONAL.XLSB loads as usual; File → Open the development
   `addin/build/ExcelToolkit.xlam`; new blank workbook). Nothing is installed and no setting changes.
   PERSONAL.XLSB SHA-256 before the session: `A4F5900D…BB81EB7D`.

| # | Step | Expected | Result |
|---|---|---|---|
| 1 | Ribbon, before any command | Format ▸ Number formats menu with four commands and tooltips; Settings ▸ Toolkit shortcuts not pressed; **Undo Toolkit Action greyed out** | |
| 2 | Numbers in A1:A3 with **different** formats; select; Number formats ▸ Number | All become `#,##0` style; Undo Toolkit Action **enabled** | |
| 3 | Press **Ctrl+Z** | Each cell gets back its **own** original format; Undo Toolkit Action **greyed out** again | |
| 4 | Currency, Percentage and Date once each: undo the first two with Ctrl+Z, the third with the **Undo Toolkit Action** button | Each applies and undoes; the button greys out after each undo | |
| 5 | Number on A1; change A1's format by hand (Ctrl+1); click Undo Toolkit Action | Refused with a message; A1 keeps the hand-set format | |

**Shortcut-precedence test.** The toolkit is identifiable: it enables Undo Toolkit Action and Ctrl+Z
reverses it; the PERSONAL.XLSB macro does neither. Each step uses all four keys (Ctrl+Shift+I, O, P, M)
on a cell holding a number.

| # | Step | Expected | Result |
|---|---|---|---|
| S0 | **Before opening the add-in**: press each key | PERSONAL.XLSB macros (baseline) | |
| S1 | Add-in open, toggle **off** (default): each key | PERSONAL.XLSB macros | |
| S2 | Toggle **on**: each key | Toolkit commands | |
| S3 | Toggle **off** again: each key | **PERSONAL.XLSB macros again** | |
| S4 | Ctrl+Shift+V with the toggle on, then off | Always the PERSONAL.XLSB macro (the toolkit does not bind V in M3a) | |
| S5 | Toggle **on**, then close **only the add-in**: Alt+F11, Ctrl+G, type `Workbooks("ExcelToolkit.xlam").Close False`, Enter, Alt+Q. The Excel Toolkit tab disappears. Press each key | PERSONAL.XLSB macros; **no "Cannot run the macro" message** | |
| S6 | Close Excel (**Don't Save** everywhere); start Excel normally, without the add-in; press each key | PERSONAL.XLSB macros | |
| S7 | Close Excel, Don't Save | PERSONAL.XLSB unchanged (hash); no security events | |

Known risk in S5: Excel documents that `Auto_Close` does **not** run when a workbook is closed by code, so
the toggle's switch-off may not run and the keys may still point at the closed add-in. Closing by code is
the only way to close a File → Open add-in without closing Excel, so the step is kept as written. Excel
closing (S6) does run `Auto_Close`, and `OnKey` assignments never survive a restart.

**Stop rule.** If S3, S5 or S6 does not give the PERSONAL.XLSB macro back, record the exact behaviour
(including any message text), stop the shortcut migration and report. Do not edit PERSONAL.XLSB, repair
anything automatically or change shortcut registrations. The formatting commands can still be accepted
through the ribbon. Excel's own (native) Ctrl+Shift+V behaviour, without PERSONAL.XLSB, is recorded
separately in M3d.

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
