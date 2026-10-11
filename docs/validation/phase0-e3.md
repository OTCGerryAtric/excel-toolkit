# Phase 0 — E3 results: reversibility and resilience

Date: 2026-10-09. Machine: owner's PC (Excel M365 x64 16.0.20326.20142, Python 3.12.10, pywin32 312).

**Status: complete.** Automated checks plus two manual checks by the owner (`prototypes/e3/manual_checks.py`, results in `prototypes/e3/results/e3_manual.json`).

## Scope

Reversibility by command category (agreed policy):

| Category | Requirement | Mechanism tested |
|---|---|---|
| Formatting toggle | Native Ctrl+Z desirable, not a gate | Excel Ctrl+Z via `Application.OnUndo` (VBA) |
| Bulk transformation | Tested reversal or equivalent safeguard | Change journal + **custom Undo** with conflict detection; **automatic rollback** on failure |
| Destructive | Confirmation, preview, recovery | Preview, explicit confirmation, journal, rollback, custom Undo |
| Read-only audit | Never modifies | (E2) |

The three mechanisms are kept separate: **Excel Ctrl+Z** (one step, registered by VBA), **custom Undo** (explicit command, refuses if cells changed since), **rollback** (automatic, after a failed operation).

Plus connection resilience: unsaved workbook, rename, close/reopen, stale process, Excel busy, Excel closing mid-operation.

## Method

- Code: `prototypes/e3/` — `excel_link.py` (explicit attach from E2 + bounded retries), `journal.py` (capture → apply → verify → undo/rollback), `E3Commands.bas` (toggle with `OnUndo`, identity handoff, busy helpers), `run_e3.py` (orchestrator), `test_call_retry.py` (retry policy against simulated busy Excel), `make_fixtures.py`.
- All changes in **disposable copies in new Excel processes**. The owner's Excel was never contacted. The code under test always attached through an explicit identity (window + process id + process start time + workbook name + full name); the orchestrator's own reference played the user (edits, renames, closing).
- Paste-as-values was the bulk example; delete-all-names the destructive example.

## Results — 27 automated checks passed, plus 5/5 retry-policy checks

### Bulk transformation (paste as values)

| Check | Result |
|---|---|
| Values, text, errors, booleans and number formats correct after the operation (formats are untouched by the operation, not journalled) | Pass |
| Custom Undo restores formulas and constants exactly (verified cell by cell) | Pass |
| Undo **refused** when a user edited an affected cell afterwards; conflicting cell reported (`Data!$A$2`) | Pass |
| Nothing restored after a refused undo (no partial overwrite) | Pass |
| Operation over an unlocked block then a locked block on a protected sheet: second block refused by Excel, **first block rolled back**, sheet identical to before | Pass |

**Finding — naive paste-as-values corrupts text.** `Range.Value2 = Range.Value2` turned the text `'0012` into the number 12 and `'1/2` into a date (46024). The journal writes constants in a form Excel stores verbatim (text with a leading apostrophe, error constants as `#N/A`), and the test confirms text, errors and booleans survive.

### Destructive (delete names)

| Check | Result |
|---|---|
| Preview lists every name: workbook, hidden, sheet-scoped, print area, broken (`#REF!`) | Pass |
| Refuses to run without explicit confirmation | Pass |
| Failure after 2 deletions rolls back to the identical name set | Pass |
| Delete removes all; Undo restores every name with scope, definition and visibility | Pass |

### Identity

| Check | Result |
|---|---|
| Unsaved new workbook (`Book2`, empty `Path`): identity by name; operation and undo work | Pass |
| Undo after Save As under a new name: refused ("no longer open (closed or renamed)") | Pass |
| Undo against the same file closed and reopened (same name and path, unsaved changes lost): **refused by the content check** — all 5 cells differ from the post-operation state | Pass |
| Stale identity after close: refused | Pass |
| Process id with a different start time (simulated pid reuse): refused | Pass |

### Excel busy

| Check | Result |
|---|---|
| Bounded retry, permanently busy (simulated `RPC_E_CALL_REJECTED`): gives up after the limit with a clear message | Pass (2.01 s for a 2 s limit) |
| Transient busy: succeeds once Excel is free | Pass |
| Busy surfacing as `AttributeError` (pywin32 quirk): recognised by probing, bounded | Pass |
| Genuine missing attribute: not retried | Pass |
| Excel gone: fails at once, no retries | Pass |
| VBA `MsgBox` open in the target Excel: read **and write** succeeded immediately | Informational |
| Built-in dialog (Format Number) open: read **and write** succeeded immediately | Informational |
| Cell edit mode via posted keystrokes | Not achieved — keystrokes posted to the window did not enter edit mode. **Manual check.** |

**Finding — dialogs do not block automation.** Neither a VBA message box nor a built-in dialog made Excel reject calls; Python could still write cells. The only real rejection seen so far was in the owner's Excel before E1 (most likely cell edit mode). Two consequences: (1) the retry policy is proven against simulated rejections, and against real edit mode only once the manual check is done; (2) the add-in can change cells while the user has a dialog open — production commands should check for that (e.g. `Application.Interactive` / `Ready`) before writing.

### Excel closing mid-operation

| Check | Result |
|---|---|
| Excel process killed 1 s into a 3,000-cell cell-by-cell operation: clear `ExcelGone` error after 1.0 s, no hang, journal marked **incomplete** (about 400 of 3,000 cells written) | Pass |
| The dead instance's identity is refused afterwards | Pass |

Nothing can be restored in that case: the unsaved workbook state is lost with Excel. The journal records what happened.

### Excel Ctrl+Z (formatting toggle) — not demonstrated

| Check | Result |
|---|---|
| Toggle on a mixed, multi-area selection goes to the first format | Pass |
| `Application.Undo` from automation runs the `OnUndo` restore | **No** (no error, nothing restored) |
| `Application.Undo` called inside VBA runs the `OnUndo` restore | **No** |
| `OnUndo` hook registered for a Python operation fires on `Application.Undo` | **No** |

`OnUndo` is documented for user-initiated undo; when the macro itself is run via `Application.Run` from automation, the registration may not take effect. This does **not** prove a real Ctrl+Z fails. **Manual check:** press Ctrl+Shift+J then Ctrl+Z in a test Excel.

## Manual checks (owner, 2026-10-09)

In a separate test Excel with a scratch workbook and the E3 add-in loaded:

| Check | Result |
|---|---|
| Ctrl+Shift+J (hotkey via `Application.OnKey`) changed the format | Pass — also closes the outstanding E1 hotkey check |
| **Real Ctrl+Z** after the toggle: Excel offered "Undo number format" and restored `0.00` | **Pass** |
| Cell edit mode (typing in F2, not committed): read attempt | **ExcelBusy after exactly 5.0 s** with the clear message |
| After Escape: read attempt | Succeeded at once |

So: Excel Ctrl+Z integration via `OnUndo` **works for a real keypress**; it only fails when Undo is triggered programmatically. And cell edit mode is a real busy state that the bounded retry handles as designed.

## Coverage statuses carried forward (for the audit engine)

| Situation | Status to report |
|---|---|
| Hidden formula on a protected sheet (E2) | Indeterminate for those cells |
| Excel busy beyond the retry window | Not checked (with reason) |
| Workbook closed, renamed or replaced since the handoff | Refused — re-run from Excel |
| Excel closed during an operation | Incomplete — journal shows how far it got |

## Limitations

- Single machine; disposable fixtures; one run of each scenario (plus earlier runs while fixing the harness).
- Busy behaviour was only exercised with simulated rejections; real edit mode pending the manual check.
- Excel's own undo history is cleared by the journal's writes (any COM write does this). Custom Undo is our command, not Ctrl+Z.
- The journal stores cell contents in a local JSON file. Fine for a prototype; production needs a retention/clean-up policy, and care because it holds workbook data.
- The prototype journal captures and restores **formulas and constants only** (`Formula2` + `Value2`, plus defined names for delete-names). Number formats, fills, borders, spill ranges, merged cells, data validation and conditional formats are not journalled. (Corrected 2026-10-11: this line previously said number formats were covered; `journal.py` never captured them.)
- Two test Excel instances left dialogs open and did not quit; the harness now terminates its own test instance (by recorded pid) if it hasn't exited 15 s after `Quit`.

## Outcome

E3 passed. Bulk and destructive safeguards (journal, conflict-checked undo, rollback, confirmation) meet the agreed policy. Formatting toggles can offer a real one-step Ctrl+Z via `OnUndo`. Real busy state (cell edit mode) fails cleanly after the 5 s bounded wait. Stopping before E4.
