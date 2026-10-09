# Project status

Last updated: 2026-10-09 — PyXLL ruled out; revised Phase 0 approved, not started. Update at every approval gate.

## Objective

Internal Excel add-in for the owner and 6–20 colleagues. All users run Windows desktop Excel and install add-ins themselves (no central IT deployment). Productivity and formula auditing are equal priorities. The existing VBA macros are the productivity baseline; the longer-term value is formula intelligence and workbook auditing.

## Permanent constraint: zero licence cost

**The add-in must be entirely free to develop, install and use** — no per-user licences, subscriptions or paid tiers, for the owner or colleagues. Every bundled component must be free for internal redistribution, with its licence documented. Non-negotiable. (Set 2026-10-09.)

## Architecture decision (provisional, pending G0)

- **VBA `.xlam` front end:** ribbon, keyboard shortcuts, formatting toggles and simple commands. Runs in-process, so hotkeys are as fast as the current PERSONAL.XLSB macros.
- **Python integration layer:** local process talking to the running Excel via COM. Bridge to be chosen in Phase 0 between:
  - direct VBA → Python with `pywin32`, or
  - the free open-source **xlwings** core (`RunPython`), no PRO features.
- **Independent Python analysis engine** (`src/excel_toolkit`): formula classification, inconsistencies, dependency analysis, audit rules. No COM/xlwings imports outside the adapter layer.
- **Results:** navigable results in Excel (sheet-free where possible, e.g. a form or task pane) and HTML reports.
- Macro security: the `.xlam` must install **without weakening macro security globally**. Candidate approaches (to verify in E6): per-user Trusted Location for the add-in folder, and/or a code-signing certificate.
- ADR-001 to be written at the end of Phase 0.

### Ruled out

| Option | Reason |
|---|---|
| PyXLL | Paid per-user licence (~£259/user/year). Ruled out 2026-10-09. Trial install was cancelled before download; no Excel registration was made (verified). |
| xlwings PRO features | Paid for commercial/internal business use. Free core only. |
| Office.js | Mostly TypeScript; Python analysis would need a server or browser runtime; restricted shortcut keys; no central-deployment need. |

## Environment (owner's machine, checked 2026-10-09)

| Item | Value |
|---|---|
| Excel | Microsoft 365 Apps, x64, build 16.0.20326.20142 |
| Python | 3.12.10, 64-bit |
| Tooling | uv 0.11.8, git 2.49 |
| Loaded add-ins | Arixcel Explorer (.xlam) only |

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
| X | ReportExternalLinks (Module2) | Audit — but currently **writes** a report sheet into the workbook |
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
- 2026-10-09: ReportExternalLinks is ported as a **read-only** audit capability. Later: extend to names, charts and other structures, flagging uninspected areas.
- 2026-10-09: PERSONAL.XLSB and all existing shortcuts are preserved until replacements are validated. Experiments use non-conflicting test shortcuts.
- 2026-10-09: Undo requirements by category:
  - Simple formatting toggles — native Ctrl+Z desirable, not a gate.
  - Bulk transformations — tested reversal or equivalent safeguard required.
  - Destructive commands — confirmation, impact preview and recovery strategy required.
  - Read-only auditing — must never modify the workbook.
- 2026-10-09: **PyXLL ruled out; zero-licence-cost constraint adopted; architecture changed to VBA `.xlam` + Python.** Roadmap, six initial feature groups and approval gates unchanged.

## Phase 0 state (revised for VBA + Python)

| Step | Status |
|---|---|
| Revised Phase 0 plan | **Approved** 2026-10-09 |
| Baseline captured, rollback documented | Done — `docs/rollback.md` |
| git + uv project | Done (`740fb1e`) |
| PyXLL removed from project; no Excel registration | Done (pyproject clean, registry verified). `.venv` cleanup pending — `pyxll.exe` locked by the cancelled installer. |
| E1 `.xlam` hotkey format cycle on a test key, timed against the VBA original | Not started |
| E2 Python reads the live workbook (correct Excel instance, unsaved edits, multi-area, other sheets) — via pywin32 **and** via free xlwings | Not started |
| E3 Reversibility by command category (toggle / bulk / destructive), incl. `Application.OnUndo` from VBA and intervening-edit detection | Not started |
| E4 Bridge latency (VBA → Python → result) and read/write timing for 1, 1k, 10k, 100k cells | Not started |
| E5 Excel's own precedent tracing via COM from Python | Not started |
| E6 Install on another Windows PC with **no manual Python/dependency install**, no global macro-security changes | Not started — needs a colleague machine |
| Bridge comparison: pywin32 vs xlwings core (simplicity, reliability, speed, install, licences, maintainability) | Not started |
| Licence register for all bundled components | Not started |
| Phase 0 results document + ADR-001 | Not started |

## Open items / dependencies

- Colleague machine for E6.
- Colleague Excel builds/bitness — unverified.
- Python runtime bundling approach (e.g. embeddable Python or uv-managed standalone Python) — choose and confirm licences in Phase 0.

## Next gate

**G0** — bridge chosen with evidence, hotkey latency matches VBA, live-workbook access proven, safety strategy agreed per command category, install on a clean PC without manual Python setup or global macro-security changes, all components free with licences documented, no undocumented critical blockers. Stop for review at G0; do not start production commands.

## Related documents

- `docs/hotkeys.md` — hotkey register
- `docs/rollback.md` — baseline and rollback
- `reports/Excel add in feature ideas.md` — competitor feature research
- `baseline/vba/PERSONAL_XLSB_olevba.txt` — existing VBA source
