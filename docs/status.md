# Project status

Last updated: 2026-10-09 — Phase 0 approved, setup in progress. Update at every approval gate.

## Objective

Internal Excel add-in for the owner and 6–20 colleagues. All users run Windows desktop Excel and install add-ins themselves (no central IT deployment). Productivity and formula auditing are equal priorities. The existing VBA macros are the productivity baseline; the longer-term value is formula intelligence and workbook auditing.

## Architecture decision (provisional, pending G0)

- **PyXLL + Python** as the Excel integration layer.
- Independent, framework-free Python core and analysis engine (no PyXLL/COM imports outside `adapters/`).
- Fallback if PyXLL fails G0: Office.js. Rejected as first choice because there is no central deployment need, shortcut keys are restricted, most code would be TypeScript, and Python analysis would need a server or Wasm.
- ADR-001 to be written at the end of Phase 0.

## Environment (owner's machine, checked 2026-10-09)

| Item | Value |
|---|---|
| Excel | Microsoft 365 Apps, x64, build 16.0.20326.20142 |
| Python | 3.12.10, 64-bit (matches Excel bitness) |
| Tooling | uv 0.11.8, git 2.49 |
| PyXLL | Not installed |
| Loaded add-ins | Arixcel Explorer (.xlam) |

Colleague machines: not yet checked.

## Existing VBA (PERSONAL.XLSB)

Source extracted to `baseline/vba/PERSONAL_XLSB_olevba.txt`. Shortcuts are set via Macro Options (Ctrl+Shift+letter). Full map in `docs/hotkeys.md`.

| Key | Macro | Kind |
|---|---|---|
| I / O / P | Number / currency (£) / percentage format cycles | Formatting toggle |
| M | Date format toggle (dd-mmm-yy ↔ mmm-yy) | Formatting toggle |
| R / C | Row height toggle / column width cycle | Formatting toggle |
| Q / E / W / Z | Yellow / green / light blue / grey fill toggles | Formatting toggle |
| K | Best-fit column width (rounds to multiple of 5) | Formatting |
| G | Grey outline toggle | Formatting toggle |
| V | Paste as values | Bulk transformation |
| F | Clear the sheet's own name from formulas | Bulk transformation |
| N | Delete all names | Destructive |
| X | ReportExternalLinks (Module2) | Audit — but currently **writes** an "External Links Report" sheet into the workbook |
| — | Find_Constants (was C, clashing) | Audit (selection) |

### Known weaknesses — to become regression tests when ported

1. **Delete Names** deletes every name (including hidden names, Print_Area / Print_Titles) with no confirmation.
2. **Clear Sheet References** only quotes sheet names containing spaces (misses e.g. `P&L`, `2024-25`) and replaces matching text in non-formula cells.
3. **Paste as Values** uses the clipboard (overwrites it) and fails on multi-area selections.
4. **Colour toggles** reset to white fill instead of No Fill (hides gridlines).
5. **ReportExternalLinks** looks like an audit but modifies the workbook: it creates or clears an "External Links Report" sheet. The replacement must be read-only.

## Agreed decisions

- 2026-10-09: The 16 macros (plus ReportExternalLinks) are intended functionality, with their current keys as configurable defaults. Porting is **outside Phase 0**; before porting, produce a behavioural inventory (exact formats, cycle order, selection handling, bugs vs intent).
- 2026-10-09: Find_Constants moves to the add-in on **Ctrl+Shift+H**; preserve its semantics (`SpecialCells(xlCellTypeConstants, xlNumbers)` on the selection) unless a change is approved.
- 2026-10-09: ReportExternalLinks is ported as a **read-only** audit capability (output to task pane / report, not a new sheet). Later: extend to names, charts and other structures, flagging uninspected areas.
- 2026-10-09: PERSONAL.XLSB and all existing shortcuts are preserved during Phase 0. PyXLL experiments use non-conflicting test shortcuts. VBA stays as rollback until replacements pass testing.
- 2026-10-09: Undo requirements by category:
  - Simple formatting toggles — native Ctrl+Z desirable, not a gate.
  - Bulk transformations — tested reversal or equivalent safeguard required.
  - Destructive commands — confirmation, impact preview and recovery strategy required.
  - Read-only auditing — must never modify the workbook.

## Phase 0 state

| Step | Status |
|---|---|
| Phase 0 plan | **Approved** 2026-10-09 |
| Baseline captured (PERSONAL.XLSB, Excel add-in registry keys, VBA source) | Done — see `docs/rollback.md` |
| Rollback instructions | Done — `docs/rollback.md` |
| git + uv project (`excel-toolkit`, Python 3.12) | Done |
| PyXLL trial installed and registered | Not started |
| E1 Hotkey format cycle, timed against VBA (non-conflicting test key) | Not started |
| E2 Live workbook reads (unsaved edits, multi-area, other sheets) | Not started |
| E3 Reversibility by category (toggle / bulk / destructive) incl. `Application.OnUndo` and intervening-edit detection | Not started |
| E4 Read/write timing: 1, 1k, 10k, 100k cells | Not started |
| E5 Excel's own precedent tracing via COM | Not started |
| E6 Clean install on a colleague's PC | Not started — needs a colleague machine |
| Phase 0 results document + ADR-001 | Not started |

## Open items / dependencies

- PyXLL licence cost (quoted ~£259 per user per year; **not verified**) — needs approval before G0.
- Colleague machine for E6.
- Colleague Excel builds/bitness — unverified.

## Next gate

**G0** — PyXLL integration proven, licence acceptable, maintainable install demonstrated, bitness/Python compatibility confirmed, safety strategy agreed per command category, no undocumented critical blockers. Stop for review at G0; do not start production commands.

## Related documents

- `docs/hotkeys.md` — hotkey register
- `docs/rollback.md` — baseline and rollback
- `reports/Excel add in feature ideas.md` — competitor feature research
- `baseline/vba/PERSONAL_XLSB_olevba.txt` — existing VBA source
