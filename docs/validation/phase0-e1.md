# Phase 0 — E1 results: number-format cycle in an .xlam

Date: 2026-10-09. Machine: owner's PC (Excel M365 x64 16.0.20326.20142).

## Question

Can the existing `Number_Format_Toggle` (PERSONAL.XLSB, Ctrl+Shift+I) move into a distributable `.xlam` on its own hotkey with identical behaviour and no loss of speed?

## Method

- **Code:** `prototypes/e1/E1Test.bas`. The cycle is held as data (an array of three formats) instead of nested `If`s, with the same rule: if the whole selection has a format in the cycle, move to the next; otherwise (including mixed formats) go to the first.
- **Hotkey:** Ctrl+Shift+J via `Application.OnKey` (in `Auto_Open` / `E1_Init`). No existing mapping touched.
- **Build:** `prototypes/e1/build_xlam.py` — imports the `.bas` into a new workbook in a **separate hidden Excel instance** (no XLSTART or add-ins loaded) and saves as `.xlam`. Uses `pywin32` (dev dependency).
- **Run:** `prototypes/e1/run_e1.py` — opens the `.xlam` in the running Excel and calls the harness, which works on a **new scratch workbook closed without saving**.
- **Equivalence:** 3 selections (single cell `B2`, block `B2:D10`, multi-area `B2:B5,D7:E9`) × 6 start states (each of the 3 cycle formats, `General`, `0.00`, mixed) = 18 single-press checks, plus a 5-press cycle from `General` per selection (3 checks). Each compares the legacy macro's result with the new one.
- **Timing:** for each selection and each macro, 10 warm-up runs then 200 timed runs via `Application.Run`, timed in-process with `QueryPerformanceCounter`. Both macros are called the same way, so the comparison is like-for-like.

## Results

**Equivalence: 21 checks, 0 mismatches.** Cycle order confirmed for all three selection types:
`#,##0` → `#,##0.0` → `#,##0.00` → back to `#,##0` (all with `_-;[Red](…)_-;-_-`). Mixed selections reset to the first format in both versions.

**Timing (ms):**

| Selection | Legacy median | New median | Legacy p95 | New p95 |
|---|---|---|---|---|
| `B2` | 0.600 | 0.558 | 0.903 | 0.755 |
| `B2:D10` | 0.883 | 0.895 | 1.309 | 1.165 |
| `B2:B5,D7:E9` | 0.990 | 0.982 | 1.252 | 1.254 |

No meaningful difference; both are around 1 ms, far below the 200 ms target.

**Safety:**
- PERSONAL.XLSB unchanged on disk (SHA-256 before = after).
- Excel add-in states unchanged; Arixcel Explorer still installed.
- No Excel settings changed. "Trust access to the VBA project object model" (`AccessVBOM`) was **already 1** before E1 and is still 1.

Raw data: `prototypes/e1/results/e1_equivalence.csv`, `e1_timings.csv`, `e1_summary.md`.

## Findings

1. Moving the VBA formatting toggles into an `.xlam` costs nothing in speed and keeps behaviour identical.
2. Holding the cycle as data works, so formats can later come from a shared settings file.
3. The scripted build needs `AccessVBOM = 1` **on the developer's machine only**. Colleagues load a finished `.xlam`; that setting is not needed to install or use it. (To be confirmed on a colleague PC in E6.)
4. `Auto_Open` does not fire when an `.xlam` is opened from code, so the runner calls `E1_Init`. When installed through the Add-ins dialog, `Auto_Open` should fire — to verify in E6.

## Limitations

- Timings measure the macro via `Application.Run`, not a physical keypress. Key dispatch is the same Excel mechanism for both, so it doesn't affect the comparison, but the hotkey itself was not exercised by the script. **Manual check:** press Ctrl+Shift+J on a selection while `E1Test.xlam` is open.
- One deliberate difference: the new version does nothing if the selection isn't a range (e.g. a chart). The original would error. Not covered by the tests.
- Undo was not part of E1 (it is E3). Both versions clear Excel's undo history.
- Single machine, single run.

## Outcome

E1 passed. Stopping before E2 as agreed.
