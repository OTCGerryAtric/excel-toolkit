# ADR-001: Excel integration architecture

- **Status:** Accepted for Phase 0 — provisional until gate G0 (E3–E6 still to run).
- **Date:** 2026-10-09
- **Evidence:** `docs/validation/phase0-e1.md` to `phase0-e4.md`

## Context

Internal Excel add-in for the owner and 6–20 colleagues, all on Windows desktop Excel, installing add-ins themselves. Equal weight on fast productivity commands (the existing PERSONAL.XLSB macros) and formula auditing. **Permanent constraint: zero licence cost** for development, installation and use.

## Decision

1. **VBA `.xlam` front end** for the ribbon, keyboard shortcuts and instant commands (formatting toggles etc.).
   - E1: a ported format cycle matched the original in 21/21 checks at ~1 ms, the same as the original.
2. **Direct pywin32** as the Python-to-Excel layer.
   - E2: 27/27 correctness checks (xlwings core 25/27); ~240 ms faster per handoff; full error fidelity including `#SPILL!`; we control retry and timeout behaviour.
3. **Explicit handoff from VBA to Python:** Excel main window handle + process id + workbook name, all passed by the caller. Python checks the window belongs to that process, attaches through a workbook window, and checks the attached Application is in that process.
   - **Never** locate the user's Excel with `GetActiveObject` / `Dispatch("Excel.Application")`. E2 showed it attaches to whichever instance registered first — the wrong one in the test.
4. **Independent Python engine** (`src/excel_toolkit`), with no COM imports outside the adapter layer.
5. **Hotkey commands stay in VBA.** A Python handoff costs ~0.9 s (start-up and imports), acceptable for audits, not for instant commands.
6. **Proposed (E4, awaiting review): hybrid execution.** A read-only persistent worker, started on demand and exiting when idle, serves interactive audit requests (~50–200 ms warm); the same engine runs one-shot as fallback and for batch reports. Requests carry an id and an explicit workbook (never `ActiveWorkbook`); results return to VBA by callback and are rejected on any identity mismatch. Short-lived COM connections per request.

## Coverage rules carried forward (from E2)

The audit engine must report each check as complete, partial or indeterminate. Known indeterminate cases so far:

- Hidden formulas on protected sheets: `Formula2` raises an error, so those cells are **indeterminate**, never clean.
- Excel busy (dialog open, cell being edited): reads fail; the result is "not checked", not "no issues".

## Rejected

| Option | Reason |
|---|---|
| PyXLL | Paid per-user licence |
| xlwings PRO | Paid for internal business use |
| xlwings open-source core | Lower error fidelity (2 failed checks), needed raw COM for 5 of the things we read, unlimited retry default can hang, extra dependency. Removed from the project 2026-10-09. |
| Office.js | Mostly TypeScript; Python would need a server or browser runtime; restricted shortcut keys; no central deployment need |
| `GetActiveObject` / `Dispatch` to find Excel | Attaches to an arbitrary instance |

## Licences (so far)

| Component | Licence | Use |
|---|---|---|
| Python 3.12 | PSF | Runtime |
| pywin32 | PSF | Excel COM access |

Full licence register, including the bundling and installer tools, to be completed in E6.

## Consequences

- Two languages: VBA for instant commands and UI, Python for analysis and safeguarded bulk operations.
- We own COM resilience: bounded retries, timeouts, clear failure messages (E3).
- Excel's native undo is cleared by any macro or COM change; reversibility needs our own mechanisms (E3).
- Colleagues need a bundled Python runtime; install route to be proven in E6.

## Open until G0

E5 (Excel's own precedent tracing), E6 (clean install on a colleague's PC, licence register).
