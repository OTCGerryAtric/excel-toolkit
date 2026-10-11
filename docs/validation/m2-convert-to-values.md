# M2 validation — ribbon, Convert to Values, one-level undo

Status: **build, automated live checks (40/40) and manual ribbon / real Ctrl+Z checklist (10/10) passed, 2026-10-11.** Open issue: text conversion speed (below). Not merged; awaiting review.
Specification: `docs/product-roadmap.md` §D. Source: `addin/`.

## What exists

| Part | File |
|---|---|
| Ribbon (one tab, Quick Tools: Convert to Values, Undo Toolkit Action, Audit Formulas) | `addin/customUI/customUI14.xml` |
| Command registry, ribbon callbacks, messages | `addin/vba/TkMain.bas` |
| Convert to Values: validation, apply, verify, rollback, timings | `addin/vba/TkConvert.bas` |
| Captured state of the converted formula cells; restore and conflict checks | `addin/vba/TkCellSnapshot.cls` |
| One-level undo record, Ctrl+Z (`Application.OnUndo`) and Undo Toolkit Action | `addin/vba/TkUndo.bas` |
| Audit Formulas | existing Phase B (`prototypes/integration/PhaseBBridge.bas`, `PhaseBResultsHook.cls`) |
| Build (separate hidden Excel → `.xlam`, then ribbon XML injected and verified) | `addin/build_addin.py` |
| Scripted live check (disposable Excel process and workbooks) | `addin/live/m2_live_check.py` |

Implementation choices to confirm live:
- **Text results** are written with the cell's number format briefly set to Text (`@`) and then
  restored, so `"0012"`, `"1/2"`, `"12%"` and `"=abc"` stay text without an apostrophe prefix (closest
  to Excel's own paste values). Empty text is written as `'` so it stays text rather than a blank cell.
  If a case fails verification, the command rolls back rather than storing a different value.
- Numbers, booleans and error values in all-formula areas are written in one bulk call; areas with
  text results or mixed constants are written cell by cell (constants are never written).
- Spill detection uses `Range.HasSpill` and `Range.SpillParent`; hidden-cell detection uses
  `SpecialCells(xlCellTypeVisible)`. Both are exercised by the live check.

## Offline evidence

`tests/addin/test_addin_offline.py` — 26 tests: ribbon callbacks resolve to Public VBA procedures with
the right signatures; ribbon buttons, registry and dispatch agree; built-in icons only; every cross-module
call resolves; the Ctrl+Z handler exists with no arguments; sources are ASCII, CRLF, `Option Explicit`,
correctly named, within VBA line/continuation limits; no clipboard, Shell, process launch, file output or
registry use in the add-in modules; ribbon injection into a real Office package adds exactly one
relationship, keeps every other part byte-identical, is idempotent, and refuses malformed XML.
Module-level declarations precede the first procedure in every module (a placement VBA rejects at compile time). Full suite: **119 passed** (93 existing + 26).

Not provable offline: that the VBA compiles and behaves correctly in Excel.

## Live tests (pending approval)

Scope: disposable Excel processes and workbooks; the development `.xlam` opened manually or by the
script; no installer, Add-ins dialog, registry, Trusted Location or trust change; PERSONAL.XLSB not
modified; stop at once if endpoint protection intervenes.

1. **Build:** `<scratch venv>/python addin/build_addin.py` (separate hidden Excel; prints `package OK`).
2. **Scripted check:** `<scratch venv>/python addin/live/m2_live_check.py --perf` — types and exact
   restore, three areas, overlapping, no formulas, cell limit, filtered (refused; visible-only allowed),
   hidden column, merged, spill anchor / range / spilled cells, legacy array, data table, protected
   (locked refused, unlocked allowed), forced failures on both write paths (rolled back), a forced failure after verified writes but before undo is
   stored (rolled back, never reported as refused), forced restore failure (incomplete, record unusable), conflict (refused, then allowed after revert), record replaced,
   sheet renamed / Save As / workbook closed (refused), Excel state restored, timings at
   1 / 1,000 / 10,000 / 100,000 cells with undo timed separately, PERSONAL.XLSB unchanged.
3. **Manual checklist** (in your Excel, File → Open `addin/build/ExcelToolkit.xlam`, disposable new workbook):

| # | Step | Expected | Result |
|---|---|---|---|
| 1 | Look at the ribbon | "Excel Toolkit" tab; three buttons with icons and tooltips; Undo Toolkit Action greyed out || Pass — tab, three buttons with icons and tooltips; Undo Toolkit Action greyed (screenshot) |
| 2 | Enter `=1+1`, `="0012"`, `=TODAY()` in A1:A3; select; click **Convert to Values** | Values remain, formula bar shows no formulas; Undo Toolkit Action enabled, tooltip names the range || Pass |
| 3 | Press **Ctrl+Z** | Excel offers "Undo Convert to Values"; the three formulas return; Undo Toolkit Action greyed || **Pass — real Ctrl+Z restored the formulas** |
| 4 | Convert again; type something in C1; press Ctrl+Z | Ctrl+Z undoes the C1 typing only (Excel's own undo) || Pass |
| 5 | Click **Undo Toolkit Action** | The A1:A3 formulas return || Pass |
| 6 | Convert again; overwrite A2; click Undo Toolkit Action | Refused, names A2, nothing changed || Pass |
| 7 | Filter a list so rows are hidden; select a range across hidden rows; Convert | Refused with the Alt+; hint. Then Alt+; and Convert: only visible cells converted || **Pass — refused with the Alt+; hint; visible-only selection converted only the visible rows** |
| 8 | Click **Audit Formulas** with no worker running | Clear message that the worker is not running; nothing launched || Pass — "The audit worker is not running … This add-in will not start Python itself." (worker deliberately not started; screenshot) |
| 9 | Press Ctrl+Shift+V on a selection | Still runs the existing PERSONAL.XLSB macro || Pass |
| 10 | Close Excel choosing **Don't Save** for every workbook | PERSONAL.XLSB unchanged (hash); no security events || Pass |

## Results — automated live run (2026-10-11, owner's managed laptop)

Separate hidden Excel processes for the build and the checks (DispatchEx: no PERSONAL.XLSB, no
add-ins); nothing installed. PERSONAL.XLSB SHA-256 unchanged; 0 Defender detections; no SentinelOne
events in the run window; no Excel process left running.

**VBA compilation:** the first build did not compile — *"Ambiguous name detected: Tk_Version"*
(VBA is case-insensitive: `Const TK_VERSION` vs `Function Tk_Version`). Every macro then reports
"Cannot run the macro". Found with VBE Debug > Compile in a disposable instance; constant renamed.
Offline tests now check module-level and public names case-insensitively. Project compiles cleanly.

**Ribbon packaging:** `customUI14.xml` injected and verified in the built `.xlam` (`package OK`).
Ribbon display and clicks are not exercised by the script (manual checklist).

**Defects found by the live checks and fixed:**
1. Error results (`#N/A`, `#DIV/0!`) raised *Type mismatch*: VBA's `And` does not short-circuit, so
   `VarType(x) = vbString And Len(x) > 0` evaluated `Len` on an error. The command rolled back, and
   the restore reported "incomplete" although the formulas were back — over-reporting, never a false
   success. Fixed with a nested test; offline test bans the pattern.
2. Text starting with an apostrophe (`'quoted`) lost it (Excel swallows one leading apostrophe even
   in a Text cell). Verification caught it and the command **rolled back with nothing changed**.
   Fixed by escaping the apostrophe.
3. Harness only: data-table fixture needed `pythoncom.Empty` for the omitted argument; the rename
   check reused a cell that was no longer a formula.

**Final automated result: 40/40** (`addin/live/results/m2_live_20261011_232739.json`):

| Area | Result |
|---|---|
| Data types | 15 formula results (number, `"0012"`, `"1/2"`, date, TRUE, `#N/A`, `#DIV/0!`, empty text, `"=abc"`, padded text, `1/3` with format, `"TRUE"`, `'quoted`, 300-char text, `"12%"`) converted with identical type, value and number format; constants untouched |
| Restore (programmatic, same procedure as Ctrl+Z) | Identical `Formula2` and recalculated values; second undo has nothing to do |
| Selections | Three areas converted, nothing else changed; overlapping refused; no formulas → nothing; over-limit refused, not truncated |
| Hidden / filtered | Selection including filtered-out rows refused; **visible-cells-only selection converted**, hidden cells untouched; hidden column refused |
| Unsupported | Merged, spill anchor, spill range, spilled cells only, legacy array, data table — all refused, nothing changed |
| Protection | Locked formula refused before writing; unlocked formula converted and undone |
| Failures | Forced failure on per-cell and bulk paths rolled back, sheet identical, Excel state restored, no record left; failure after verified writes but before undo is stored → rolled back (never "refused"); forced restore failure → "incomplete", record unusable |
| Conflicts and lifecycle | Undo refused after an edit (names the cell, writes nothing), allowed once reverted; new command replaces the record; record cleared after undo |
| Identity | Undo refused after sheet rename, Save As under another name, and workbook close |

**Performance** (`m2_live_20261011_232617.json`; end-to-end from call to undo registered; restore timed separately; ms):

| Cells | Kind | Convert total | of which capture / apply / verify | Restore |
|---|---|---|---|---|
| 1 | number | 14 | 0.4 / 4.7 / 8.4 | 12 |
| 1 | text | 15 | 0.3 / 5.6 / 8.5 | 11 |
| 1,000 | number | 47 | 21 / 18 / 8 | 15 |
| 1,000 | text | 181 | 25 / 141 / 14 | 157 |
| 10,000 | number | **306** | 170 / 113 / 22 | 63 |
| 10,000 | text | **1,596** | 204 / 1,308 / 83 | 1,342 |
| 100,000 | number | 3,338 | 2,073 / 1,098 / 163 | 482 |
| 100,000 | text | 15,574 | 2,366 / 12,509 / 696 | 12,070 |

Numbers meet the 10,000-cell target (< 1 s for convert and restore). **Text misses it** (1.6 s convert,
1.3 s restore at 10,000), because each text cell is written individually with a temporary Text format.
At 100,000 cells capture dominates for numbers (per-cell array/spill checks). Not optimised in M2.

**Not established by the automated run:** real Ctrl+Z (Excel runs `OnUndo` handlers only for a
real keypress) and the ribbon — the manual checklist above.

## Results — manual checklist (2026-10-11, owner, own Excel)

The development `.xlam` was opened with File → Open in the owner's Excel, against a new blank
workbook. All 10 steps passed (results in the table above). In particular: **a real Ctrl+Z after
Convert to Values restored the original formulas**, the ribbon's Undo Toolkit Action enabled and
disabled correctly, the filtered-selection refusal and visible-only conversion behaved as specified,
Audit Formulas reported the stopped worker clearly and launched nothing, and Ctrl+Shift+V still ran
the PERSONAL.XLSB macro.

Afterwards: PERSONAL.XLSB SHA-256 unchanged; no Excel process left; 0 Defender detections;
SentinelOne logged one routine burst of 14 event-143 entries, none naming Excel, Python or the add-in;
no Trusted Documents record for the add-in; the only add-in `OPEN` key is still Arixcel's (no install).

## Open issues

- **Text conversion speed** misses the 10,000-cell target (1.6 s convert, 1.3 s restore; 15.6 s /
  12.1 s at 100,000) because text cells are written one at a time. Numbers meet it (306 ms / 63 ms).
  Planned: one focused optimisation pass on bulk text writes and the per-cell safety checks, without
  weakening exact type preservation or undo. The provisional 100,000-cell limit stays meanwhile.
  Whether this blocks M2 acceptance is a review decision.
