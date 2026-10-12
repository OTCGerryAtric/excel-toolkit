# M3c validation — Clear Sheet References and Find Constants

Status: **offline implementation only (2026-10-12).** Offline tests pass; nothing has been built or run in
Excel. Live tests need separate scoped approval. M3d not started.
Plan: `docs/product-roadmap.md` (M3, "Clear Sheet References (M3c)"). Source: `addin/`.

## What M3c adds

| Part | File |
|---|---|
| Clear Sheet References (formula-aware rewrite) and Find Constants | `addin/vba/TkFormulas.bas` |
| Undo record for rewritten formulas (Formula2 before and after, cell by cell) | `addin/vba/TkFormulaSnapshot.cls` |
| Ribbon: new **Formulas** group with both buttons (keytips C and H) | `addin/customUI/customUI14.xml`, `TkMain.bas` |
| Rewrite cases shared by the offline tests and the live check | `addin/live/m3c_cases.py` |
| Scripted live check (prepared, **not run**) | `addin/live/m3c_live_check.py` |

No keyboard shortcuts yet: PERSONAL.XLSB keeps Ctrl+Shift+F.

## Clear Sheet References

**The legacy macro** runs `Cells.Replace` across the whole sheet, replacing `Name!` (or `'Name'!` only when
the name contains a space) anywhere in any cell, ignoring case. It therefore also changes text cells,
string literals (`"Sheet1!A1"`), similarly named sheets (`MySheet1!` on `Sheet1`), external references
(`[Book.xlsx]Sheet1!`) and 3D ranges, and it misses quoted names without spaces (`'P&L'!`): roadmap
known weakness 2.

**The toolkit** looks only at **formula cells** (`SpecialCells(xlCellTypeFormulas)`; text is never read or
written) and rewrites each formula by reading it piece by piece:
- **String literals** (`"..."`, with `""` escapes) are copied unchanged, so `INDIRECT("'P&L'!A1")` keeps its
  text.
- **Quoted sheet names** (`'...'!`, with `''` escapes) are removed only when the name, unescaped, is this
  sheet (ignoring case). Quoted external names (`'[Book.xlsx]P&L'!`, `'C:\…\[Book.xlsx]P&L'!`) and quoted 3D
  ranges (`'My Sheet:Other'!`) never match.
- **Unquoted names** followed by `!` are removed only when the name is this sheet and it does not follow
  `]` (an external workbook), `#` (an error value such as `#REF!`) or a `:` that ends a 3D range
  (`Sheet1:Sheet3!`). A `:` after a cell, column or row reference, or after `)`, is an ordinary range, so
  `Sheet1!A1:Sheet1!B2` becomes `A1:B2`.
- Similar names (`P&L 2`, `MySheet1`, `Sheet11`) never match, because whole names are compared.

Removing a sheet's own name from its own formulas does not change what they refer to. A sheet-scoped name
(`Sheet1!MyName` on Sheet1) resolves to the same name without the prefix.

**Scope:** the selection, which can be mixed (formulas, constants, text, blanks) or multi-area. With a
**single cell selected, the whole sheet**, as Excel's Find & Replace and as the legacy macro (which always
worked on the whole sheet).

**Safeguards:**
- A protected sheet is refused.
- A legacy array formula (Ctrl+Shift+Enter) that would change is refused rather than rewritten cell by
  cell.
- More than 100,000 changed formulas is refused.
- After writing, every changed cell must still hold a formula that no longer names this sheet. If not,
  everything is rolled back.

**Undo:**
- Every changed formula's exact `Formula2` before the command is restored and verified.
- Conflict check: each changed cell must still hold the formula Excel stored after the command.
- Identity check (workbook, sheet name) as in M2.

## Find Constants

Same as the legacy macro: `SpecialCells(xlCellTypeConstants, xlNumbers).Select` on the selection. A single
selected cell means the whole sheet (Excel's own rule). It only changes the selection: no undo record, the
existing record is left alone, and nothing is written. When there are no typed-in numbers it says so,
where the legacy macro did nothing.

## Decisions for review

1. **Scope:** selection; a single cell = the whole sheet.
2. **Refusals:** protected sheets, legacy array formulas that would change, over 100,000 changed formulas.
3. **Text inside formulas is never rewritten**, including `INDIRECT("Sheet!A1")`. Such references keep
   their sheet name, which is harmless.
4. **No shortcuts in M3c** (as M3b).

## Offline evidence

`tests/addin/test_addin_offline.py`: 138 tests in this file; full offline suite 231 passed.
- **The rewrite rules.** The 41 cases in `m3c_cases.py` pass through a Python model of
  `Tk_RemoveSheetPrefix` that follows the same steps, and the model is idempotent on every result. The
  cases cover quoted and unquoted names, `''` escapes, non-ASCII names, case differences, ranges with a
  prefixed second half, other sheets, similar names, external references (quoted and not), 3D ranges,
  string literals with `""` escapes, `#REF!`, function names and sheet-scoped names. The VBA function is
  checked statically for the same rules, and live against the same cases.
- **The legacy text replace** is shown to damage MySheet1, external, string-literal, 3D, `#REF!` and
  text-cell cases that the toolkit keeps.
- **Formula cells only:** `SpecialCells(xlCellTypeFormulas)`, `HasFormula`, array formulas refused, and no
  `Range.Replace` anywhere in the add-in.
- **Find Constants** uses the legacy call (`xlNumbers` = 1) and never writes or records undo.
- The snapshot class provides the undo-record interface, plus all the earlier structural checks.

Deliberately breaking a rule (the `#` guard, case-insensitive matching, numbers-only constants) makes the
tests fail. Not provable offline: compilation, how Excel stores the rewritten formulas, timings.

## Live tests (pending approval)

1. **Build** in a separate hidden Excel, then a **VBA Compile** check in a disposable instance.
2. **Scripted** (`addin/live/m3c_live_check.py --perf`, disposable Excel):
   - all 41 rewrite cases through the VBA function itself;
   - on a sheet `P&L` with sheets `P&L 2`, `Other` and `O'Brien`: a selection mixing own-sheet,
     other-sheet, similar-name, string-literal, INDIRECT-text and spill formulas with text and number
     cells, giving the expected formulas, untouched text and unchanged values;
   - exact undo; single cell = whole sheet; mixed selection with nothing to do; multi-area;
   - the same result as the legacy macro on plain references, while the legacy macro also damages a string
     literal and a text cell;
   - an array formula refused; conflict, identity and protection refusals; forced write and restore
     failures;
   - Find Constants selecting the same cells as the legacy macro (block, single cell, two areas), reporting
     "nothing" when appropriate, and leaving the undo record alone;
   - timings at 10,000 formulas.
3. **Manual** (owner): ribbon, real Ctrl+Z after Clear Sheet References, Find Constants on a real model;
   PERSONAL.XLSB hash and security checks as before.
