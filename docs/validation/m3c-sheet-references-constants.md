# M3c validation — Clear Sheet References and Find Constants

Status: **scripted live check passed, 28/28 (2026-10-12).** Offline implementation reviewed (`2ea8167`);
the design and decisions were approved, with one amendment (ranges Excel could read two ways), made below.
Manual testing awaits approval. Not merged. M3d not started.
Plan: `docs/product-roadmap.md` (M3, "Clear Sheet References (M3c)"). Source: `addin/`.

## What M3c adds

| Part | File |
|---|---|
| Clear Sheet References (formula-aware rewrite) and Find Constants | `addin/vba/TkFormulas.bas` |
| Undo record for rewritten formulas (Formula2 before and after, cell by cell) | `addin/vba/TkFormulaSnapshot.cls` |
| Ribbon: new **Formulas** group with both buttons (keytips C and H) | `addin/customUI/customUI14.xml`, `TkMain.bas` |
| Rewrite cases shared by the offline tests and the live check | `addin/live/m3c_cases.py` |
| Python model of the rewrite, used by the offline tests and the live check | `addin/live/m3c_rewrite_model.py` |
| Scripted live check | `addin/live/m3c_live_check.py` |
| Probe of how Excel itself writes and reads such references (no add-in) | `addin/live/m3c_syntax_probe.py` |

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
  `]` (an external workbook) or `#` (an error value such as `#REF!`).
- **After a `:`** (quoted or unquoted) the prefix is removed only when the first half of the range is `)`
  (`INDEX(...):Sheet1!B9`) or a real cell reference (`A1`, `$A$1`, within `XFD1048576`) with no prefix or
  this sheet's prefix, so `Sheet1!A1:Sheet1!B2` becomes `A1:B2`. Anything else before the `:` makes it a 3D
  range, which Excel itself does (see below): `A:Sheet1!B2`, `1:Sheet1!B2` and `Data:Sheet1!B2` keep their
  names. A first half on another sheet keeps the prefix too: `Sheet2!A1:Sheet1!B2` is `#VALUE!`, but
  `Sheet2!A1:B2` would mean cells on Sheet2.
- **Formulas Excel could read two ways are left entirely unchanged and reported.** If a formula that would
  change holds a `:` followed by a sheet prefix whose first half is not `)`, a cell or itself prefixed
  (`AB:Sheet1!B2`, `Sheet1:Data!B2`), or the rewrite would create one (`Sheet1!AB:Other!B2`), the whole
  formula is kept as it is. The command then ends with a **warning** naming how many were kept and the
  first cell ("Check those by hand"); if nothing else changed, it says **nothing** was changed and why.
- Similar names (`P&L 2`, `MySheet1`, `Sheet11`) never match, because whole names are compared.

### How Excel reads these references (probe, 2026-10-12)

`m3c_syntax_probe.py` in a disposable Excel, sheets named `Sheet1`, `A`, `AB`, `XFD`, `XFE`, `A1`, `XFD1`,
`XFE1`, `1`, `R`, `C`, `RC`, `R1C1`, `Data`, `Other`:
- Excel **quotes** sheet names that are cells or R1C1 references (`'A1'!`, `'XFD1'!`, `'1'!`, `'R'!`,
  `'C'!`, `'RC'!`, `'R1C1'!`) but **not** names that look like columns (`A!`, `AB!`, `XFD!`) or cells beyond
  the sheet (`XFE1!`).
- `=SUM(A:Sheet1!B2)` is read as the **3D range** from sheet A to Sheet1 and stored as `=SUM(Sheet1:A!B2)`;
  the same for `AB:`, `XFE:`, `XFE1:`, `Data:`, and `1:` (stored `'Sheet1:1'!B2`). Only a real cell before
  the `:` (`A1:Sheet1!B2`, `$A$1:`, `A$1:`, `XFD1:`, `B1048576:`) gives an ordinary range.
- `=SUM(AB:'Sheet1'!B2)` is **not** 3D (it is `#NAME?`) but is stored as `=SUM(AB:Sheet1!B2)`, the same
  text as the 3D range would have before Excel reorders it, and that text re-entered is read as 3D. This is
  why such formulas are never rewritten: writing any change back means Excel parses the text again.
- `Sheet1!A:Sheet1!C` and `Sheet1!1:Sheet1!3` are not formulas Excel accepts as ranges (the first parses
  as something else, the second is refused), so the earlier offline cases for them were removed.
- `Sheet2!A1:Sheet1!B2` and `INDEX(Sheet2!A:A,1):B2` are `#VALUE!`; `Sheet2!A1:B2` is a range on Sheet2.
- Excel stores an unprefixed rewritten range in order (`C5:B2` is stored `B2:C5`); the conflict check
  uses the formula as Excel stored it.

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

## Decisions (approved at review, 2026-10-12)

1. **Scope:** selection; a single cell = the whole sheet.
2. **Refusals:** protected sheets, legacy array formulas that would change, over 100,000 changed formulas.
3. **Text inside formulas is never rewritten**, including `INDIRECT("Sheet!A1")`. Such references keep
   their sheet name, which is harmless.
4. **Find Constants** selects typed-in numbers only.
5. **No shortcuts in M3c** (as M3b).
6. **Review amendment:** a formula Excel could read two ways is left unchanged and reported, never
   rewritten.

## Offline evidence

`tests/addin/test_addin_offline.py`: 161 tests in this file; full offline suite 254 passed.
- **The rewrite rules.** The 57 cases and 6 two-way cases in `m3c_cases.py` pass through
  `m3c_rewrite_model.py`, a Python model of `Tk_RewriteSheetPrefix` that follows the same steps, and the
  model is idempotent on every result. The cases cover quoted and unquoted names, `''` escapes, non-ASCII
  names, case differences, ranges with a prefixed second half, first halves on other sheets, sheets named
  like columns (`A`, `AB`), 3D ranges including those, cells beyond the sheet (`XFE1`, `A0`), other sheets,
  similar names, external references (quoted and not), string literals with `""` escapes, `#REF!`,
  function names and sheet-scoped names; cell-reference bounds have their own test. The VBA is checked
  statically for the same rules, and live against the same cases.
- **The legacy text replace** is shown to damage MySheet1, external, string-literal, 3D, `#REF!` and
  text-cell cases that the toolkit keeps.
- **Formula cells only:** `SpecialCells(xlCellTypeFormulas)`, `HasFormula`, array formulas refused, and no
  `Range.Replace` anywhere in the add-in.
- **Find Constants** uses the legacy call (`xlNumbers` = 1) and never writes or records undo.
- The snapshot class provides the undo-record interface, plus all the earlier structural checks.

Deliberately breaking a rule (the `#` guard, case-insensitive matching, numbers-only constants, and for
the amendment: columns or rows counted as the first half of a range, the other-sheet first-half rule, the
two-way check) makes the tests fail. Not provable offline: compilation, how Excel stores the rewritten formulas, timings.

## Live results (2026-10-12)

Scoped approval: build, compile and automated live tests only; scratch environment; nothing installed;
no Office trust or registry changes; PERSONAL.XLSB unchanged; the audit worker not started.

- **Build** in a separate hidden Excel: package OK. **VBA Compile** in a disposable instance: compiled
  cleanly.
- **Probe** (`m3c_syntax_probe.py`, no add-in): findings above;
  `addin/live/results/m3c_probe_20261012_151133.json`.
- **Scripted** (`m3c_live_check.py --perf`): **28/28 passed**;
  `addin/live/results/m3c_live_20261012_151952.json`. Two earlier attempts stopped on harness bugs, not
  add-in faults, after 14/14 and 15/15 checks passed: the legacy sheet was selected while another workbook
  was active, then renamed back to the wrong name, and pywin32's `Address` needed `GetAddress`.
  - VBA `Tk_RemoveSheetPrefix` gives the expected result for all 57 cases and leaves all 6 two-way cases
    unchanged.
  - **Formula meaning:** a workbook with sheets `Sheet1`, `A`, `AB`, `Data`, `Sheet2`, `Other` (distinct
    numbers on each). On `Sheet1`, 19 formulas: 14 rewritten, 3D ranges and the other-sheet first half
    kept, 2 two-way formulas kept and reported (`warning: … 2 formula(s) … the first is in E27`). On `A`,
    5 formulas: 3 rewritten, 1 3D range kept, 1 two-way formula kept and reported. **Every calculated
    result was identical before and after** (including `#VALUE!` and `#NAME?` results that a wrong rewrite
    would have turned into numbers). Undo restored every formula and result exactly. Selecting only the
    two-way formulas gives "nothing", with the reason.
  - Selection fixture (`P&L`): own-sheet prefixes removed; `P&L 2`, `Other`, string literals, INDIRECT
    text and text cells untouched; values unchanged; exact undo; single cell = whole sheet; mixed selection
    with nothing to do; multi-area.
  - Legacy comparison: same result as the legacy macro on plain references; the legacy macro also changed
    a string literal and a text cell, the toolkit did not.
  - Refusals and failures: CSE array formula refused; conflict; identity (sheet renamed); protected sheet;
    forced write failure rolled back; forced restore failure reported as incomplete.
  - Find Constants: same cells as the legacy macro (block, single cell = whole sheet, two areas);
    "nothing" when there are none; the undo record is left alone.
  - **Timings** (10,000 formulas): rewrite 2.6 s (capture 0.6 s, write and verify 2.0 s), undo 2.0 s;
    10,000 formulas with nothing to rewrite 19 ms.
- **Security:** no Defender detections; no SentinelOne Operational or Firewall events during the runs; no
  Excel process left behind; PERSONAL.XLSB SHA-256 `A4F5900D…87431BB81EB7D`, unchanged.

## Live test plan (as approved)

1. **Build** in a separate hidden Excel, then a **VBA Compile** check in a disposable instance.
2. **Scripted** (`addin/live/m3c_live_check.py --perf`, disposable Excel):
   - all rewrite cases through the VBA function itself;
   - calculated results before and after on sheets named like columns, with 3D ranges and two-way
     formulas;
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
