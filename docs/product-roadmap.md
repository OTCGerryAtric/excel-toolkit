# Product roadmap — capability inventory, gaps and next milestone

Date: 2026-10-11. Basis: repository at `58be1ca`, offline suite re-run (93 passed).
Consolidates three inputs: the **Master Product & Development Brief v2.0** (11 Oct 2026), the
reviewer's response to the first assessment (five agreed corrections), and a second, independent
Claude assessment (reconciled in [Reconciliation](#reconciliation-of-the-two-assessments)).
**Nothing in this document has been implemented; M2 starts only after the decisions in §E.**

## Direction (from the brief)

Excel Toolkit becomes a **professional Excel power-user add-in**: everyday productivity commands,
formula intelligence and auditing in one coherent product, for 6–20 colleagues first, possibly wider
later. Progress is measured by **working functionality in Excel**, not test counts or documents.
Non-negotiables: no paid licences; respect Office and endpoint security; **every modifying command
must be safely reversible** (Ctrl+Z where feasible); read-only commands never modify the workbook.

Architecture is unchanged: **VBA `.xlam`** (ribbon, shortcuts, instant and modifying commands) +
**Python persistent worker** (analysis) over pywin32, with the analysis engine independent of the
front end. Deployment (G0) stays blocked.

### Agreed corrections (reviewer, 2026-10-11)

1. Phase B's controlled live demonstration at `58be1ca` counts as **Milestone 1 achieved**.
2. Routine modifying commands run **in-process in VBA** (E3 design: capture, conflict check,
   rollback, `Application.OnUndo`). Single-operation custom Undo is distinct from preserving Excel's
   native Undo history, which is **not** promised. Python commands stay read-only; any future Python
   modification needs its own tested recovery.
3. **No production ribbon XML or UserForms exist.**
4. External Links **Quick Scan inspects the live workbook**; saved-file structural inspection is an
   additional capability; coverage is always reported.
5. Architecture and deployment restrictions stand; the manually started worker has run successfully
   on the managed laptop (not an installation method).

## A. Capability inventory

Status: **Live** = ran in real Excel with recorded evidence; **Offline** = automated tests only;
**Absent**; **Blocked**.

| Capability | Status | Source | Evidence | Remaining |
|---|---|---|---|---|
| 17 legacy macros (16 hotkeys) incl. ReportExternalLinks | Live as legacy; **none ported** | `baseline/vba/PERSONAL_XLSB_olevba.txt`, `docs/hotkeys.md` | Daily use; 5 known defects (`docs/status.md`) | Behavioural inventory; port onto the command framework; defects become regression tests |
| Format cycle in an `.xlam` (E1) | Live (prototype, 1 command) | `prototypes/e1/E1Test.bas`, `build_xlam.py` | 21/21 equivalence, ~1 ms; hotkey confirmed manually | Scripted build needs AccessVBOM=1 (dev only) |
| Attach to the right Excel; read unsaved workbook (E2) | Live | `prototypes/e6/engine/excel_link.py` (= E3 copy) | 27/27; identity = window + pid (+ start time) + workbook name + full path | — |
| Ctrl+Z via `Application.OnUndo` | Live, **number format toggle only** | `prototypes/e3/E3Commands.bas` | Real keypress restored the format; programmatic `Application.Undo` does **not** fire it | Untried for value/formula changes |
| Journal, conflict-checked Undo, rollback | Live, **Python-side** | `prototypes/e3/journal.py` | 27 checks + 5 retry checks; text kept (`'0012`, `'1/2`); refuses after intervening edit; rollback on protected sheet | Captures **formulas and values only** — not number formats, fills or borders (`phase0-e3.md` overstates this). Port the design to VBA |
| Persistent worker + named pipe (E4) | Live, started manually | `prototypes/e4/e4_worker.py`, `prototypes/integration/audit_worker.py` | 28/28; warm ~47 ms | ADR-001 requirements open: busy-safe delivery, user-only pipe security descriptor (currently default + remote rejection, no message-size bound), stale-request cancel. One-shot fallback deliberately removed (it means Excel launching Python) |
| Dependency parser (E5) | Live comparison, prototype | `prototypes/e5/deps_parser.py` | 11/14 right, 3 flagged, 0 silently wrong | Not in `src/`; basis for Formula Explorer |
| Formula-inconsistency detector | Offline + live via Phase B | `src/excel_toolkit/formula_audit/` | Gates 1–3; saved/live parity; 93 tests | Active sheet only; Gate 4 deferred |
| Excel-triggered audit (Phase B) | **Live — Milestone 1** | `prototypes/integration/` | `58be1ca`: unsaved edit read, exactly 2 findings, navigation, workbook unchanged, no EDR intervention | Trigger is Alt+F8; worker manual; failure MsgBox untested live; code comments still say "not executed" |
| External-link report | Live as legacy, **modifies the workbook** | PERSONAL.XLSB Module2 | — | Read-only rewrite; wider coverage |
| Ribbon, UserForms, command search | **Absent** | — | — | All new |
| Deployment (E6) | **Blocked (G0)** | `prototypes/e6/` | EDR quarantined the installer | IT/security questions unanswered; never re-run E6 |

Do not rebuild: instance attachment, live read, worker/pipe transport, callback correlation, detector,
JSON findings transport, navigation to the original workbook.

## B. Gaps

- **Ribbon and GUI:** no `customUI14.xml`, no build step to add it to the `.xlam`, no tooltips,
  icons or shortcut labels, no reusable results view. Use built-in Office icons (`imageMso`); avoid
  ActiveX controls (TreeView availability on 64-bit Office and policy are uncertain).
- **Reliable Undo:** a VBA command framework — declare impact → capture before-state (formulas,
  values with E3's text-preservation rules, number formats, fills, borders per area) → apply →
  verify → register `OnUndo` → keep the record for an explicit **Undo Toolkit Action**, which refuses
  and names the conflicting cell if anything changed since. One level; earlier native history lost.
- **Safe productivity commands:** none ported. Defects to fix: clipboard overwrite and multi-area
  failure (Ctrl+Shift+V), white fill instead of No Fill, unquoted sheet names and non-formula
  replacement (Ctrl+Shift+F), Delete Names without preview/confirmation, link report that writes into
  the workbook. Add a busy/edit-mode check before writing.
- **Terminology:** today's Ctrl+Shift+V does `Selection.Copy` + `PasteSpecial` onto itself — it
  **converts the selection to values** in place. Under the brief's definitions that is **Convert to
  Values**; **Paste Values** (clipboard → values) does not exist yet.
- **External Links Inspector:** legacy scan covers registered links, formulas (cell-by-cell loop),
  names and chart series. Missing: conditional formatting, data validation, pivot caches,
  connections, Power Query, OLE objects; bulk reads; coverage reporting; read-only output.
- **Formula Explorer:** parsers exist; tree model, sub-expression values (flagged where unreliable)
  and the window do not.
- **Bracket Tracker:** Excel runs no VBA or ribbon callbacks while a cell is being edited, so a live
  count while typing in the formula bar needs a keyboard hook (excluded). Use the brief's fallbacks:
  live count inside our own Explorer/editor box; count for the selected formula.
- **Command search:** no registry or Find a Tool. The registry (name, description, keywords,
  shortcut, category, undo class) should drive ribbon labels and tooltips so they cannot drift.
- **Distribution:** G0 blocked on IT/security answers, not code. Only in-process VBA commands can
  reach colleagues before the worker-install question is solved; even a VBA-only `.xlam` still needs
  the downloaded-file trust question answered.

## C. Roadmap

Each milestone is usable in Excel on its own. Deployment (G0) runs in a separate lane.

| # | Milestone | Reuse | New | Main risk | Acceptance |
|---|---|---|---|---|---|
| M1 | Excel-triggered audit | Phase B | — | — | **Done** (`58be1ca`); docs reconciled |
| **M2** | **First ribbon + command framework + reversible Convert to Values** | E1 build, E3 `OnUndo` and text rules, Phase B | Ribbon tab, command registry, VBA Undo framework | `OnUndo` unproven for value changes | See §D |
| M3 | Port the productivity commands; add clipboard Paste Values | M2 framework, legacy macros | Defect fixes; Delete Names with preview + confirmation + custom Undo | Volume; shortcut migration | Each legacy behaviour matches (or is a documented fix); each modifying command reverses live; shortcuts move one by one |
| M4 | External Links Quick Scan (read-only) | Module2 logic, E5 external-ref detection, bulk reads, Phase B results pattern | Live scan of formulas, names, links, charts; coverage summary; Deep Scan later | Speed on large books | ~2 s or less on ordinary workbooks; workbook unchanged; coverage statement correct on a fixture |
| M5 | Formula Explorer v1 + bracket count | E5 parser, normaliser | Tree, values, navigate and back | Tree control; VBA vs Python parser | 10 representative formulas incl. INDIRECT/OFFSET shown as "not evaluated" |
| M6 | Find a Tool + cleanup pack | M2 registry | Search, favourites, trim / text→number / case / fill blanks | Scope creep | "values" finds both commands, clearly distinguished; colleague reverses three commands unaided |

Later: Deep Scan, Audit hardening (ADR requirements), detector extensions, G0 with IT.

## D. Next milestone — M2

**Delivered in Excel**

1. An **Excel Toolkit** ribbon tab, one group (Quick Tools): **Convert to Values**, **Undo Toolkit
   Action**, and **Audit Formulas** (existing Phase B; says clearly when the worker isn't running).
   Icons, tooltips and shortcut labels come from the command registry.
2. **Convert to Values** in VBA on the shared framework: clipboard-free, multi-area, text and errors
   preserved (E3 rules), busy/edit-mode check, refuses protected cells with a clear message and rolls
   back any area already written.
3. **Ctrl+Z** restores formulas, values and number formats via `OnUndo`. **Undo Toolkit Action**
   does the same on demand and refuses, naming the cell, if anything changed since. (After a user
   edit, Ctrl+Z undoes that edit instead — so the conflict check is tested on the explicit command.)
4. Tooltip and docs state the limitation: Excel's earlier Undo history is cleared, as with any macro.

**Not in M2:** other macros, clipboard Paste Values, Find a Tool, any installer, registry or trust change.

**How it is shown to work** (results in `docs/validation/m2-convert-to-values.md`)

- Before → after → restore identical for: leading zeros (`'0012`), date-like text (`'1/2`), a
  formula returning number-like text, `#N/A`, TRUE, dates, blanks, number formats; single range,
  three-area selection; filtered rows, a dynamic-array spill, merged cells (handled or refused
  clearly); protected sheet (refused or rolled back).
- **Real Ctrl+Z keypresses** (programmatic Undo doesn't fire `OnUndo`); Undo Toolkit Action refused
  after an intervening edit.
- Ribbon buttons clicked by hand; the built `.xlam` checked for `customUI14.xml` and modules.
- Timing at 1 / 1,000 / 100,000 cells (target: 10,000 cells under 1 s); PERSONAL.XLSB SHA-256 unchanged.

## E. Decisions requiring approval

| # | Decision | Recommendation |
|---|---|---|
| 1 | Ribbon build: inject `customUI14.xml` into the `.xlam` zip after the scripted build (dev machine only; AccessVBOM already on) | Approve |
| 2 | Loading the dev `.xlam` | **File → Open each session**; not the Add-ins dialog (writes the registry `OPEN` key = installing) |
| 3 | First command | **Convert to Values** (today's Ctrl+Shift+V behaviour, clipboard-free) in M2; clipboard Paste Values in M3 |
| 4 | Ctrl+Shift+V in M2 | **Ribbon only**; take over the key once the command passes |
| 5 | Test method | Manual checklist **plus** a scripted before/after comparison in a separate, disposable Excel instance (Excel automation — needs explicit approval) |
| 6 | Undo depth | Single level; revisit only on user demand |
| 7 | Results view for M4 | Separate results workbook (proven in Phase B); decide on a window at M5 |
| 8 | Delete Names (M3) | Release only as preview + confirm + custom Undo, never one-click |
| 9 | Where features run while G0 is blocked | Quick Scan in VBA (colleague-usable sooner); decide Explorer's parser (VBA copy vs Python) at M5 |

Not decisions now: Excel-DNA, installer method, worker auto-start — unchanged until IT answers the
questions in `docs/architecture/deployment-decision.md`.

## Reconciliation of the two assessments

Both reached the same conclusions (M1 done; ribbon and Undo framework are the gap; M2 first). Points
settled against the code:

- **Taken from the second assessment:** the E3 journal does not capture number formats; the macro
  count is 17 on 16 hotkeys; results workbook for M4; Delete Names as a confirmed custom-Undo command.
- **Corrected in the second assessment:** it described clipboard paste as Ctrl+Shift+V's legacy
  meaning (the macro actually converts in place); it expected Ctrl+Z to refuse after an intervening
  edit (only the explicit command can); it took over Ctrl+Shift+V within M2 (deferred until the
  command passes); its test list lacked filtered rows, spills, merged cells and real keypresses; it
  described identity as pid + workbook name and listed the one-shot fallback as open.
