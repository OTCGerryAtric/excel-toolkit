# M2 validation — ribbon, Convert to Values, one-level undo

Status: **implemented and offline-tested; live tests NOT yet run** (awaiting explicit approval).
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
| 1 | Look at the ribbon | "Excel Toolkit" tab; three buttons with icons and tooltips; Undo Toolkit Action greyed out | |
| 2 | Enter `=1+1`, `="0012"`, `=TODAY()` in A1:A3; select; click **Convert to Values** | Values remain, formula bar shows no formulas; Undo Toolkit Action enabled, tooltip names the range | |
| 3 | Press **Ctrl+Z** | Excel offers "Undo Convert to Values"; the three formulas return; Undo Toolkit Action greyed | |
| 4 | Convert again; type something in C1; press Ctrl+Z | Ctrl+Z undoes the C1 typing only (Excel's own undo) | |
| 5 | Click **Undo Toolkit Action** | The A1:A3 formulas return | |
| 6 | Convert again; overwrite A2; click Undo Toolkit Action | Refused, names A2, nothing changed | |
| 7 | Filter a list so rows are hidden; select a range across hidden rows; Convert | Refused with the Alt+; hint. Then Alt+; and Convert: only visible cells converted | |
| 8 | Click **Audit Formulas** with no worker running | Clear message that the worker is not running; nothing launched | |
| 9 | Press Ctrl+Shift+V on a selection | Still runs the existing PERSONAL.XLSB macro | |
| 10 | Close Excel choosing **Don't Save** for every workbook | PERSONAL.XLSB unchanged (hash); no security events | |

## Results

_Not yet run._
