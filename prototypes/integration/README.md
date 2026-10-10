# Excel integration scaffolding (offline prep — not executed/installed)

Prepared for the minimal end-to-end Excel demonstration. **None of this is run, built or
installed on the current managed laptop** (Excel COM + external Python execution remain
blocked post-E6). It is staged so the authorised test environment can wire it quickly.

## The one detector, two sources

The detector depends only on a small read-only interface
(`src/excel_toolkit/formula_audit/worksheet_source.py` → `WorksheetSource`):

```
saved .xlsx ──OpenpyxlWorksheet──┐
                                 ├─▶ scan_worksheet  (identical findings/coverage/order)
live Excel  ──snapshot_from_com──┘   InMemoryWorksheet
```

Parity between the two sources is proven offline in
`tests/formula_audit/test_snapshot.py` (identical reports across the whole corpus, a
fake-COM conversion, and an unsaved-edit scenario).

## Files

- `ExcelToolkitAudit.bas` — minimal VBA: `Toolkit_AuditFormulas` (one command), a results
  list in a *separate* workbook (audited book untouched), and `Toolkit_GotoFinding`
  (jump to a cell). Source only; not imported/compiled/installed.
- `audit_live.py` — Python entry point sketch: attach (reuse E2 handoff) → `snapshot_from_com`
  → `scan_worksheet` → emit findings. Only the snapshot+scan steps run offline (tested);
  the live COM attach is left `NotImplementedError` on purpose.

## Unresolved live-Excel assumptions (confirm on the authorised machine)

- `Range.Formula` (US A1 syntax, `,` separators — what the tokenizer expects), not
  `FormulaLocal`; `Cell.HasFormula` to split formulas from constants.
- `UsedRange` for bounds (can over-report); merged cells, array/dynamic-array (`Formula2`,
  spill) formulas not special-cased yet.
- Attaching to the correct instance via the E2 handoff; reading unsaved edits live.
- Delivering findings back to VBA (stdout lines vs a worker channel) and not modifying the
  audited workbook.
