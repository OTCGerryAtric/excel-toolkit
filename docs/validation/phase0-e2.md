# Phase 0 — E2 results: Python reading the live workbook

Date: 2026-10-09. Machine: owner's PC (Excel M365 x64 16.0.20326.20142, Python 3.12.10, pywin32 312, xlwings 0.37.5 open-source core).

## Questions

1. Can Python reliably attach to the **right** Excel instance and workbook, given an explicit handoff from VBA?
2. Does it read the live state correctly (unsaved edits, multi-area selections, other sheets, errors, dynamic arrays, hidden/protected sheets)?
3. Is reading strictly read-only?
4. Does it fail cleanly when the target is gone?
5. Direct pywin32 or the free xlwings core?

## Method

- **Fixtures** (`prototypes/e2/make_fixtures.py`, built in a hidden Excel): `E2_FixtureA.xlsx` with sheets Inputs, Calc, Errors, Dyn, Hidden (hidden), VeryHidden (very hidden), Protected (hidden formula, protected), Bulk (10,000 formulas); a **decoy** with the same name and layout but different values; and `E2_FixtureB.xlsx`.
- **Two new Excel processes** (`prototypes/e2/run_e2.py`): X opens FixtureB, then FixtureA (target), then the handoff add-in. Y opens the decoy **under the same file name** `E2_FixtureA.xlsx`. The owner's own Excel was also running (a third process).
- **Setup writes only:** in X, unsaved edits (`Calc!B2 = 999`, `Calc!B3 = =B2*3`) and a multi-area selection `Calc!B2:B5,D7:E9`.
- **Handoff** (`prototypes/e2/E2Handoff.bas`): VBA `Shell`s the reader with `Application.Hwnd`, `GetCurrentProcessId()` and `ActiveWorkbook.Name`. Five handoffs per backend.
- **Reader** (`prototypes/e2/e2_reader.py`):
  - **pywin32:** checks the window belongs to the given process, finds the workbook window (`XLDESK` → `EXCEL7`), gets Excel's object model from it via `WM_GETOBJECT` with a 5 s timeout, then checks the attached Application's process id.
  - **xlwings:** checks the window/process, `xw.apps[pid]`, checks process id, `app.books[name]`, reads with `formula2` and `options(err_to_str=True)`.
  - **naive:** `GetActiveObject("Excel.Application")`, for comparison only.
- **Read-only check:** full before/after snapshots of the target and decoy (every formula and value, number format, fill, bold, formula-hidden flag, sheet visibility and protection, sheet order, names, saved flag, active sheet, selection, calculation mode).
- **Failure cases:** workbook not open; window/process mismatch; target workbook closed; Excel process closed.

## Results

### Correctness

| | pywin32 | xlwings |
|---|---|---|
| Checks passed | **27 / 27** | 25 / 27 |
| All 5 runs identical | Yes | Yes |
| Correct instance and file, with a same-named decoy in another process | Yes | Yes |
| Correct workbook with several open in the instance | Yes | Yes |
| Unsaved edits (value, formula, recalculated result, "unsaved" flag) | Yes | Yes |
| Multi-area selection (`$B$2:$B$5`, `$D$7:$E$9`) | Yes | Yes (via `.api`) |
| Cross-sheet formulas, reference into hidden sheet | Yes | Yes |
| Error types (`#DIV/0!` `#N/A` `#REF!` `#VALUE!` `#NAME?` `#NUM!` `#SPILL!`) | All 7 | 6 of 7 — `#SPILL!` came back as a raw number |
| Text `'#N/A` kept distinct from the error `#N/A` | Yes | **No** — `err_to_str` turns errors into strings, so the two can't be told apart |
| Dynamic arrays (`Formula2`, spilled cells, `A1#`, spill range) | Yes | Yes (spill info via `.api`) |
| Hidden vs very hidden sheets | Yes | Via `.api` only (xlwings `.visible` is true/false) |
| Protected sheet: detected, values readable | Yes | Yes |
| Bulk 1k / 10k formulas: count, first and last cell | Yes | Yes |
| Explicit handoff to the decoy instance reads the decoy (Inputs!A1 = 77) | Yes | Yes |

xlwings needed raw COM (`.api`) for 5 things: workbook saved flag, multi-area areas, hidden vs very hidden, sheet protection, spill info.

### Wrong-instance risk confirmed

The naive `GetActiveObject("Excel.Application")` attached to **the owner's own Excel** (pid 11752), not the target (pid 30676). It would have silently audited the wrong workbook set. The explicit handoff never did.

### Timings (ms, median of warm runs 2–5)

| | Handoff to result (wall) | Python process | Attach (incl. COM imports) | Read 7 small sheets | Read 1k formulas + values | Read 10k formulas + values |
|---|---|---|---|---|---|---|
| pywin32 | 893 | 783 | 629 | 89 | 4.1 | 19.6 |
| xlwings | 1,134 | 1,034 | 824 | 115 | 8.0 | 27.2 |

First (cold) run: pywin32 854 ms, xlwings 1,177 ms.

### Failure handling

All 8 cases (4 per backend) failed cleanly with a clear message and exit code 2:
workbook not open, window/process mismatch, target workbook closed, Excel process closed. pywin32 took 160–720 ms, xlwings about 1,250–1,300 ms.

### Read-only

Target and decoy snapshots identical before and after all reads.

## Findings

1. **The identity is the process id, not the window handle.** Since Excel 2013 each workbook has its own top-level window and `Application.Hwnd` follows the active one. The first full run caught this: xlwings found a different window of the same process. The handoff now passes window + process id + workbook name, and the reader checks the window belongs to that process and that the attached Application is in that process. Workbook names are unique within one Excel process, so (process id, workbook name) identifies the workbook.
2. **Never use `GetActiveObject` / `Dispatch("Excel.Application")` to find the user's workbook.** Confirmed above.
3. **Hidden formulas on protected sheets can't be read.** `Formula2` on a range containing one raises an error rather than returning anything. The audit engine must treat such ranges as **indeterminate** (and could fall back cell by cell), never as clean.
4. **xlwings retries forever by default when Excel is busy.** During E2, a coding slip of mine (a lower-case COM attribute name) sent xlwings' retry wrapper into an endless loop, because it mistakes "attribute not found under that spelling" for "Excel busy" and `N_COM_ATTEMPTS` defaults to unlimited. With pywin32 directly, a busy Excel fails immediately ("Call was rejected by callee", seen earlier in E1), so we choose the retry and timeout policy ourselves.
5. **Each handoff costs about 0.9 s**, mostly starting Python and importing the COM libraries. Fine for audits; too slow for anything that should feel instant. Hotkey commands stay in VBA (E1). E4 should test a long-running Python process to remove the start-up cost.

## Limitations

- One machine, one session; 5 handoffs per backend.
- `Shell` is fire-and-forget: the reader writes JSON to a file and VBA doesn't yet get a result back. Returning results to Excel is for E4.
- Not tested: Excel busy (cell edit mode or a dialog open) during a read; Excel closing in the middle of a read; an unsaved new workbook (`Book1`) as the target; workbooks open in two windows.
- Fixture bug found and fixed during the run: writing the text `#N/A` without a leading apostrophe makes Excel store a real error.
- The `pywin32` attach code calls `WM_GETOBJECT` itself; xlwings uses the equivalent `AccessibleObjectFromWindow`. Both depend on a workbook window existing (an Excel instance with no workbooks open can't be attached to — reported cleanly).

## Recommendation

**Use direct pywin32.**

- It passed every correctness check; xlwings failed two (error fidelity) that matter for an audit tool.
- It is faster (about 240 ms less per handoff, and around twice as fast on bulk reads).
- xlwings needed raw COM for five of the things we read, so it would not save us from writing COM code.
- Its unlimited retry default is a hang risk; with pywin32 we set our own timeouts.
- One fewer dependency to bundle. (pywin32 is needed by xlwings anyway.)
- xlwings' main convenience, `RunPython`, isn't needed: the VBA `Shell` handoff with explicit identity works.

Licences: pywin32 — PSF licence; xlwings core — BSD 3-clause (the `xlwings.pro` subpackage is paid and was not used). If the recommendation is accepted, xlwings will be removed from the project's dependencies.

Stopping before E3 as agreed.
