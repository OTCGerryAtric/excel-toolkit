# Project status

Last updated: 2026-10-10 — E6 functional work passed but its **deployment method was rejected** after the EDR
removed project files during the install test (see `docs/incidents/2026-10-10-edr-file-deletion.md`).
Deployment redesign approved as **design/investigation only**; no install/security changes pending review. Update at every approval gate.

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
| PyXLL removed from project; no Excel registration | Done (pyproject clean, `.venv` clean, registry verified, no installer process) |
| E1 `.xlam` hotkey format cycle on a test key, timed against the VBA original | **Passed** 2026-10-09 — 21/21 equivalence checks; ~1 ms, same as VBA. See `docs/validation/phase0-e1.md`. Ctrl+Shift+J keypress confirmed manually in E3. |
| E2 Python reads the live workbook (correct Excel instance, unsaved edits, multi-area, other sheets) — via pywin32 **and** via free xlwings | **Done** 2026-10-09 — pywin32 27/27, xlwings 25/27; explicit handoff (window + process id + workbook name) never attached to the wrong instance; naive `GetActiveObject` did. Reads verified read-only. Recommendation: direct pywin32. See `docs/validation/phase0-e2.md`. Awaiting review. |
| E3 Reversibility by command category (toggle / bulk / destructive), incl. `Application.OnUndo` and intervening-edit detection; resilience (Book1, rename, reopen, stale identity, busy, Excel closing mid-operation) | **Passed** 2026-10-09 — journal, conflict-checked custom Undo, rollback and confirmed destructive commands work; real Ctrl+Z via `OnUndo` works (manual keypress); cell edit mode fails cleanly after the 5 s bounded wait. See `docs/validation/phase0-e3.md`. Awaiting review. |
| E4 One-shot vs persistent worker; results returned to VBA; size scaling 1–100k cells; isolation and failure handling | **Done** 2026-10-09 — 28/28 checks. Worker warm ~50 ms (fresh) to ~200 ms vs one-shot ~0.8 s; 69/69 interleaved requests correct; crash, restart and Excel-closure handled; 0 CPU idle. **Accepted** (hybrid, worker first, one-shot fallback) with 5 production requirements recorded in ADR-001. See `docs/validation/phase0-e4.md`. |
| E5 Dependency tracing: Excel COM vs a static Python parser, 14 constructs | **Done** 2026-10-10 — COM `DirectPrecedents` silently wrong in 8/14, trace arrows silently wrong in 3/14 (INDIRECT, OFFSET, external); Python parser right in 11/14 and honestly flagged the other 3 (0 silent). Recommendation: parser primary, COM for run-time gaps. See `docs/validation/phase0-e5.md`. Awaiting review. |
| E6 Install on another Windows PC with **no manual Python/dependency install**, no global macro-security changes | **Not passed** 2026-10-10 — packaged PyInstaller engine + PowerShell installer worked functionally (runtime 11/11, install lifecycle 19/19) but **the EDR classified the install as an intrusion** (`windows.penetration`), killed it and quarantined the exe/script files (project, `.venv`, `claude.exe`). E4 slowdown explained (power-throttling, fixed). Deployment method rejected; redesign in `docs/architecture/deployment-decision.md`. See `docs/validation/phase0-e6.md` and the incident record. |
| ADR-001 drafted (docs/architecture/ADR-001-excel-integration.md); xlwings removed from dependencies | Done 2026-10-09 |
| Bridge comparison: pywin32 vs xlwings core (simplicity, reliability, speed, install, licences, maintainability) | Done in E2 — recommends direct pywin32 (install/packaging comparison still to come in E6) |
| Licence register for all bundled components | Not started |
| Phase 0 results document + ADR-001 | Not started |

## Security incident (2026-10-10)

the EDR (the laptop's endpoint protection; Defender is passive) removed project files during the E6 install
test. Confirmed it was the EDR, by one detection, deleting by file type; git intact, all recoverable; our
scripts/Claude did not delete anything. Blast radius is confined to the project directory plus `claude.exe`
(shared toolchains — uv cache, base Python, PowerShell — all intact). **Recovery and deployment are separate
workstreams.** Recovery (preserve evidence → git restore from `cf67e37` → re-apply 4 saved fixes → `uv sync` from
lockfile → verify + safe non-deployment tests → leave the E6 installer unexecuted) is **planned, not done**, and is
**gated on confirming the the EDR detection `<redacted-threat-id>` is no longer actively responding**. (`claude.exe`
was deleted but self-restored via the 16:48 auto-update to 2.1.296 — no reinstall needed.) Evidence preserved outside the repo at
`C:\Users\GerryPidgeon\IncidentEvidence\2026-10-10-ExcelAddIn\`. Do **not** re-run the flagged prototype. Full
record and the four confirmed/unverified distinctions: `docs/incidents/2026-10-10-edr-file-deletion.md`.

## Architecture & plan (APPROVED 2026-10-10; `docs/architecture/deployment-decision.md`)

Architecture and development direction **approved**; no further architecture documents needed — next milestone is
working formula-analysis functionality with automated tests. Deployment stays blocked (G0).

**VBA `.xlam` + Python is the preferred working baseline** (decision 2026-10-10). Excel-DNA (`.xll`) + Python is a
credible **alternative to investigate**, not an equal commitment; a full C# rewrite is a contingency. The **Python
analysis engine (E2/E4/E5) is preserved** regardless of front end. **E6 blocks deployment, not all development.**

Work is split into **three workstreams**: (1) **Python analysis engine** — development priority, built/tested offline
without Excel or any add-in; (2) **Excel front end** — keep VBA baseline, evaluate Excel-DNA only where it offers a
measurable advantage; (3) **Packaging & deployment** — G0 blocked, normal install strategy, no bypass/EDR
workaround/unapproved trust change.

**Next deliverable:** a standalone **Python formula-inconsistency detector** (offline workbook fixtures, automated
regression tests) — front-end-independent. **Pattern analysis** (does a formula follow its neighbours' relative-
reference pattern) is distinct from **dependency analysis** (E5 resolve statuses), each with its own
coverage/confidence; the **full E5 dependency parser is NOT a prerequisite**. The front-end choice is not a
prerequisite either. **Precision is a first-class acceptance criterion:** fixtures must include legitimate
variations that must not be flagged, tests must measure false-positive rate (not just coverage), and findings are
ranked by confidence (high-confidence vs weaker "for review"), never asserted as definite errors. Authorising this =
ordinary **offline** Python dev/tests only; Excel automation/COM, packaged-exe execution and deployment stay
separately gated.

Verified facts: a signature does **not** exempt a MOTW `.xlam` (MS16-088) and Excel **blocks untrusted `.xll` from
the internet** (2023, unblock the specific file first) — changing front end does not remove the Office-trust problem;
**Trusted Locations** bypass Office checks and may be policy-locked (don't adopt by default). **VSTO needs paid VS** →
excluded on cost; free .NET route is **Excel-DNA (Zlib) + .NET SDK**. **Office add-in trust** and **the EDR
acceptance** are separate, neither assumed solved by signing. Corrected claims: embeddable Python is **PSF/python.org**
(not Microsoft) and **`.py` files can't be Authenticode-signed** (integrity via packaging/hash/manifest); **Arixcel's
.NET component loads via the VSTO `Excel\Addins` registration, not `OPEN`** (only its companion `.xlam` uses `OPEN`).
Minimal Excel-DNA `.xll` test = **feasibility only**, gated on security clearance.

## Open items / dependencies

- **IT/security question** (precise form in the decision doc): what triggered detection `<redacted-threat-id>`; is it
  still active; the approved way to distribute an internal Excel add-in with a local Python component (MOTW/Trusted
  Location, Python-launch); company-CA code-signing availability; whether Trusted Locations are usable on managed devices.
- Recover the deleted files (post security-clearance; `claude.exe` already self-restored).
- Colleague machine for the redesigned deployment.
- "Trust access to the VBA project object model" was already on (`AccessVBOM = 1`) before the project started. The scripted `.xlam` build relies on it (dev machine only). Owner to decide whether to keep it on.
- Colleague Excel builds/bitness — unverified.
- Python runtime bundling approach (e.g. embeddable Python or uv-managed standalone Python) — choose and confirm licences in Phase 0.

## Next gate

**G0** — bridge chosen with evidence, hotkey latency matches VBA, live-workbook access proven, safety strategy agreed per command category, install on a clean PC without manual Python setup or global macro-security changes **and without tripping endpoint protection** (the E6 blocker), all components free with licences documented, no undocumented critical blockers. Stop for review at G0; do not start production commands.

## Related documents

- `docs/hotkeys.md` — hotkey register
- `docs/rollback.md` — baseline and rollback
- `reports/Excel add in feature ideas.md` — competitor feature research
- `baseline/vba/PERSONAL_XLSB_olevba.txt` — existing VBA source
