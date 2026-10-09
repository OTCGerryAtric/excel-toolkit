# Phase 0 — E5 results: dependency tracing feasibility

Date: 2026-10-10. Machine: owner's PC (Excel M365 x64 16.0.20326.20142, Python 3.12.10, pywin32 312, openpyxl 3.1.5).

## Question

How far can we rely on Excel's own precedent tracing, how far can a standalone Python parser go, and can each method tell us when it is incomplete?

## Method (`prototypes/e5/`)

- **Fixture** (`make_fixtures.py`): `E5_Deps.xlsx` with sheets Inputs, Calc, Data (table `Sales`), My Sheet; defined name `MyDefinedName` → `Inputs!$B$1`; an external workbook `E5_External.xlsx`. 14 scenario formulas in `Calc!B1:B14`.
- **Ground truth:** what each formula really depends on, including run-time targets (e.g. the cell INDIRECT actually points at).
- **Methods compared** (`run_e5.py`), in a new Excel process on a disposable copy:
  1. **COM `DirectPrecedents`**: Excel's built-in property.
  2. **COM trace arrows**: `ShowPrecedents` + `NavigateArrow` over every arrow and link (the programmatic version of Trace Precedents, including the dashed arrow to other sheets and workbooks), then `ClearArrows`.
  3. **Python static parser** (`deps_parser.py`): openpyxl's formula tokenizer (MIT) plus workbook metadata (defined names, tables, sheet order) read once via COM. No calculation. Every result carries a status: resolved, partial (flagged) or indeterminate (flagged).
- **Verdicts:** compared cell by cell with the ground truth. "Resolved" means exactly right. A method that misses dependencies **without saying so** is marked **WRONG, silently** — the dangerous case for an audit tool.

## Results

| # | Construct | Formula | COM DirectPrecedents | COM trace arrows | Python parser |
|---|---|---|---|---|---|
| S1 | Same-sheet cells | `=F1+G1` | Resolved | Resolved | Resolved |
| S2 | Cross-sheet | `=Inputs!A1*2` | **Silent miss** (same sheet only) | Resolved | Resolved |
| S3 | Range | `=SUM(F2:F101)` | Resolved | Resolved | Resolved |
| S4 | Absolute + relative | `=$F$1+G2` | Resolved | Resolved | Resolved, with anchoring: F1 absolute, G2 relative (matches Excel's R1C1 `=R1C6+R[-2]C[5]`) |
| S5 | Table structured reference | `=SUM(Sales[Revenue])` | **Silent miss** (table on another sheet) | Resolved (`Data!B2:B6`) | Resolved (via table metadata) |
| S6 | Defined name | `=MyDefinedName*2` | **Silent miss** | Resolved (`Inputs!B1`) | Resolved (via names) |
| S7 | INDIRECT | `=INDIRECT(H1)` | **Silent miss** of `Inputs!A2` | **Silent miss** of `Inputs!A2` (only `H1`) | **Indeterminate, flagged** (`H1` + "target is text built at run time") |
| S8 | OFFSET | `=OFFSET(F1,0,1)` | **Silent miss** of `G1` | **Silent miss** of `G1` (only the anchor `F1`) | **Partial, flagged** (`F1` + "actual target depends on arguments") |
| S9 | Spill reference | `=SUM(D1#)` | Resolved (`D1:D5`) | Resolved | **Partial, flagged** (anchor `D1`; extent only known at run time) |
| S10 | External workbook | `='…\[E5_External.xlsx]Sheet1'!$A$1` | **Silent miss** | **Silent miss** (no arrow link, no error; external file not opened) | Resolved (`[E5_External.xlsx]Sheet1!A1`) |
| S11 | LET local name | `=LET(x,F1,x*G1)` | Resolved | Resolved | Resolved (`x` recognised as local, not a name) |
| S12 | 3D reference | `=SUM(Inputs:Data!A1)` | **Silent miss** | Resolved (all 3 sheets) | Resolved (expanded by sheet order) |
| S13 | Whole column | `=SUM(F:F)` | Resolved (as used cells `F1:F101`) | Resolved | Resolved (`F:F`) |
| S14 | Quoted sheet name | `='My Sheet'!A1+1` | **Silent miss** | Resolved | Resolved |

| Method | Exactly right | Flagged as incomplete | Silently wrong |
|---|---|---|---|
| COM DirectPrecedents | 6 / 14 | — | **8 / 14** |
| COM trace arrows | 11 / 14 | — | **3 / 14** (INDIRECT, OFFSET, external) |
| Python static parser | 11 / 14 | 3 / 14 (INDIRECT, OFFSET, spill) | **0 / 14** |

Timing per formula (median): DirectPrecedents 1.6 ms, trace arrows 24 ms (max 54 ms), Python parser 0.1 ms.

Safety: workbook contents unchanged (full snapshot); trace arrows change the active sheet and selection, which were restored; no external workbook was opened.

## Findings

1. **Excel's COM tracing never says when it is incomplete.** For INDIRECT, OFFSET and external links it returns a plausible but short list with no warning. An audit built only on COM would say "these are all the dependencies" when they are not. COM results alone must never be reported as complete for formulas containing INDIRECT, OFFSET or external references.
2. **`DirectPrecedents` only sees the formula's own sheet.** It missed 8 of 14 cases. It is useful only as a fast same-sheet check.
3. **Trace arrows (`NavigateArrow`) cover other sheets, tables, names and 3D references**, but are slow (~24 ms per formula, ~4 minutes for 10,000 formulas), change the selection and active sheet (must be restored), and missed the external reference entirely.
4. **The static parser was right or honestly flagged in every case**, and is ~200 times faster than trace arrows. Its gaps are exactly the run-time constructs: INDIRECT, OFFSET (correctly flagged) and spill extent (flagged; Excel's `SpillingToRange` fills that gap at run time).
5. **The two complement each other.** The parser gives structure, coverage status, anchoring and external links; COM gives run-time facts (spill extents, current values, and possibly evaluating INDIRECT/OFFSET targets). This points to a design where **the parser is primary** and COM fills specific run-time gaps.

## Coverage statuses for the audit engine (added to ADR-001's list)

| Construct | Status the engine should report |
|---|---|
| A1 cells/ranges, cross-sheet, quoted sheets, whole rows/columns, 3D, LET locals | Resolved |
| Defined names pointing at ranges | Resolved; partial if the name holds a formula or constant |
| Table column references (`Table[Column]`) | Resolved; other specifiers (`[#This Row]`, `[#Totals]`, column ranges) not yet handled → partial |
| Spill references (`A1#`) | Partial statically; resolved when Excel's `SpillingToRange` is read |
| External workbook references | Resolved as "external: [Book]Sheet!Cell"; the external file's contents are out of scope |
| OFFSET | Partial (anchor known) |
| INDIRECT | Indeterminate |
| Formula the parser can't read | Indeterminate, never "no dependencies" |

## Limitations

- One fixture, 14 formulas, one machine. No nested combinations (e.g. INDIRECT inside a name, names with relative references, LAMBDA, dynamic-array functions returning references such as `CHOOSECOLS`/`TAKE`).
- The tokenizer cannot read the spill operator `#`; the prototype strips it before tokenizing. Production needs a parser that understands it natively.
- Defined names with sheet scope, relative references in names, and table specifiers other than a single column are not handled.
- Ground truth was written by hand for these formulas.
- External references were tested with the external file closed only.

## Outcome

Feasible. A Python static parser, flagged honestly, gives good coverage and is fast; Excel COM must not be trusted on its own. Recommended direction for the production dependency engine: **parser primary, COM for run-time gaps, explicit coverage status on every result.** Stopping before E6 as agreed.
