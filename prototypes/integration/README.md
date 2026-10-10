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
  list in a *separate* workbook (audited book untouched), and `Toolkit_GotoFinding`, which
  navigates to the cell **in the original audited workbook** (resolved via the object model,
  so apostrophes/spaces in names are safe), never the results workbook. Source only.
- `audit_live.py` — Python entry point sketch: parse handoff → attach (reuse E2) →
  `snapshot_from_com` → `scan_worksheet` → `findings_to_json`. Only the snapshot/scan/JSON
  steps run offline (tested); the live COM attach is left `NotImplementedError` on purpose.

## Identity contract (handoff)

`hwnd | workbookName | workbookFullPath | worksheetName`. The **PID is not supplied** — Python
derives it from the window handle, as E2's tested attachment does. The eventual
implementation **reuses E2's attachment logic**, not a new instance-finding method.

## Findings transport

Structured **JSON** (array of objects: `workbook, workbook_path, sheet, cell, kind,
confidence, conflict, expected_by_axis, message`) — robust to tabs, newlines and unusual
names, unlike delimited text. Each record keeps the audited workbook identity for navigation.

## Unresolved live-Excel assumptions (confirm on the authorised machine)

- `UsedRange.Formula` read in **one bulk COM call** (US A1 syntax, `,` separators — what the
  tokenizer expects), not `FormulaLocal`; a string starting `=` is a formula, other values
  are constants; Excel errors / odd COM types coerce to strings (treated as boundaries).
  Handles the scalar return for a single-cell range.
- `UsedRange` bounds can over-report; merged cells and array/dynamic-array (`Formula2`,
  spill) formulas are not special-cased yet.
- Attaching to the correct instance/worksheet via the E2 handoff; reading unsaved edits live.
- Returning the JSON to VBA (stdout vs a worker channel vs a temp file) and parsing it there
  with a bundled JSON parser; never modifying the audited workbook.
