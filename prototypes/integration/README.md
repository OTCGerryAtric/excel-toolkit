# Excel integration — Excel-triggered formula audit

**Current status:** the Phase B flow is **live-validated** (2026-10-11, commit `58be1ca`) on the owner's
managed laptop, against the disposable `demo_workbook.xlsx`. Nothing is installed: the VBA is imported
into an unsaved development workbook and the worker is started by hand. Distribution remains blocked (G0).

## 1. The working integration (Phase B)

```
Excel (VBA, PhaseBBridge)                    Python (started manually)
  PhaseB_AuditActiveSheet ──JSON request──▶  audit_worker.py  (E4 named pipe, one per user)
     id + token + identity + sheet              │
                                                ▼
                                             audit_engine.py
                                               attach by identity (excel_link) → confirm sheet
                                               → snapshot_from_com (2 bulk reads) → scan_worksheet
                                               → display_json
  PhaseB_OnResult ◀──COM callback (Application.Run '<macro book>'!PhaseBBridge.PhaseB_OnResult)
     correlation check → results in a SEPARATE workbook → double-click → cell in the audited workbook
```

| File | Role |
|---|---|
| `audit_worker.py` | Persistent worker. Reuses E4's `pipe_name` / `create_pipe` / `accept_loop` verbatim (one worker per user, local inbound-only pipe, remote clients rejected). Only the `audit` op is allowed. **Started by hand; never by Excel.** |
| `audit_engine.py` | One read-only request; never raises for an Excel-side problem — failures return `status="failed"` with the error, never an empty success. The callback procedure is fixed in Python; the request supplies only the macro workbook's name. |
| `phaseb_contract.py` | Correlation rules: reject unknown, already answered, stale (superseded), identity-mismatch and token-mismatch callbacks. Mirrored in VBA. |
| `PhaseBBridge.bas` | VBA front end: pipe write only (no Shell, one-shot or auto-start); says clearly if no worker is listening; per-request id + token; flat JSON reader; renders the results workbook. |
| `PhaseBResultsHook.cls` | Application-level `SheetBeforeDoubleClick` sink for navigation; no VBA-project object-model access, so no trust change. |

Shared engine pieces: `src/excel_toolkit/formula_audit/` (`snapshot_from_com`, `scan_worksheet`,
`display_json`) and `prototypes/e6/engine/excel_link.py` (identity, attach, bounded retries).
Offline tests: `tests/phaseb/` (fake COM) and `tests/formula_audit/` — 93 passing at `58be1ca`.

## 2. How the controlled demonstration was run

1. Start the worker from the scratch venv: `python audit_worker.py <state_dir> --idle-seconds N`.
2. In Excel, create an unsaved workbook (e.g. `Book2`) and import `PhaseBBridge.bas` and
   `PhaseBResultsHook.cls` (both must have CRLF line endings — `.gitattributes` enforces this).
3. Immediate window: `Application.Run "'Book2'!PhaseB_Configure", "<state_dir>"` (the Immediate window
   runs in the selected project, so qualify the workbook).
4. Open `demo_workbook.xlsx`, change `Demo!D3` from `=C3*(1+$A$10)` to `=C3*(1+$A$11)`, **do not save**.
5. With `Demo` active: Alt+F8 → `PhaseB_AuditActiveSheet`.
6. Double-click a results row to navigate. Stop the worker with a `shutdown` request or let it idle out.

## 3. Evidence and limitations

**Live result (2026-10-11):** the worker attached to the right Excel (by window, process and workbook)
and read the unsaved D3 edit; status ok, **exactly 2 findings (D3, D4)** in ~138 ms; callback accepted;
results in a separate workbook with hidden navigation columns; double-click navigated to
`demo_workbook.xlsx / Demo / $D$3`; audited workbook unchanged in memory (formulas and values) and on
disk; no Defender or SentinelOne detection; Excel started no Python.

**Defects found and fixed during the run:**
- The callback was unqualified (`PhaseBBridge.PhaseB_OnResult`), so Excel looked only in the active
  (audited) workbook — "Cannot run the macro". Now workbook-qualified.
- `PhaseBResultsHook.cls` had LF line endings, so the VBA editor imported it as a standard module.
  Fixed; `.bas`/`.cls` pinned to CRLF.

**Limitations:** worker started manually; trigger is Alt+F8 (no ribbon yet); the worker-failure
message is tested offline only; active sheet only; single machine, one disposable workbook. ADR-001
production requirements still open: busy-safe delivery, a user-only pipe security descriptor and
message-size bound, and discarding requests when their workbook closes.

## 4. Historical scaffolding (superseded — kept for reference)

Earlier staging, before live runs were authorised. Do not extend these; Phase B replaces them.

- **Stage 2 (`stage2_live_audit.py`, `make_demo_workbook.py`):** read-only live check, later run under
  explicit approval (`92a0ac6`): saved {D4} vs live {D3, D4}, saved/live coverage parity, results
  workbook and navigation, workbook unchanged. Its docstring still carries the pre-run "DO NOT RUN"
  warning; it is retained as the record of that test.
- **`ExcelToolkitAudit.bas` + `audit_live.py`:** first one-command sketch with a handoff of
  `hwnd | workbook | full path | sheet` and the live attach deliberately left `NotImplementedError`.
  **Never completed** — superseded by Phase B's request, which also carries pid and a correlation token.
- **Assumptions that were open at that stage, now settled by the live runs:** a bulk read is used, but
  as two calls (`UsedRange.Formula` for formula text, `UsedRange.Value` for typed constants — a single
  `.Formula` read misclassified numeric constants as text); attaching via the E2 identity works; JSON
  returns via the COM callback and is parsed by a small flat-JSON reader in VBA. Still open: `UsedRange`
  over-reporting, and merged cells and dynamic-array/spill formulas are not special-cased.

### One detector, two sources

The detector depends only on `WorksheetSource` (`src/excel_toolkit/formula_audit/worksheet_source.py`):
saved `.xlsx` via `OpenpyxlWorksheet`, live Excel via `snapshot_from_com` → `InMemoryWorksheet`, with
identical findings proven offline in `tests/formula_audit/test_snapshot.py` and live in Stage 2.
