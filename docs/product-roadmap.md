# Product roadmap — capability inventory, gaps and next milestone

Date: 2026-10-11. Basis: repository at `58be1ca`, offline suite re-run (93 passed).
Consolidates three inputs: the **Master Product & Development Brief v2.0** (11 Oct 2026), the
reviewer's response to the first assessment (five agreed corrections), and a second, independent
Claude assessment (reconciled in [Reconciliation](#reconciliation-of-the-two-assessments)).
Revised after the reviewer's full review of `0d5d543` (M2 Undo contract, selection behaviour,
performance definition, decisions 8 and 9).
**Nothing in this document has been implemented; M2 starts only after this revision is reviewed and
scoped live testing is authorised.**

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
- **Reliable Undo:** a VBA modifying-command framework with one contract for every command:
  validate → capture → apply → verify (roll back on failure) → register `OnUndo` → keep one record for
  an explicit **Undo Toolkit Action** with workbook-identity and conflict checks. **M2 implements only
  the state M2 needs** (formulas, values and their types); other property types (number formats,
  fills, borders) are added when a command that changes them is built, behind the same contract.
  One level only; Excel's earlier native Undo history is lost and is not promised. Full M2 contract: §D.
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
| M3 | Port the productivity commands; add clipboard Paste Values | M2 framework, legacy macros | Defect fixes; the framework gains number formats, fills and borders as commands need them. **Delete Names stays disabled** (E8) | Volume; shortcut migration | Each legacy behaviour matches (or is a documented fix); each modifying command reverses live; shortcuts move one by one |
| M4 | External Links Quick Scan (read-only) | Module2 logic, E5 external-ref detection, bulk reads, Phase B results pattern | Live scan (unsaved changes included) of formulas, names, links, charts; coverage summary; Deep Scan later. **Implementation (VBA, Python worker or hybrid) chosen by a small benchmark at the start of M4** — coverage, time and COM-call count | Speed and COM overhead on large books | ~2 s or less on ordinary workbooks; workbook unchanged; coverage statement correct on a fixture |
| M5 | Formula Explorer v1 + bracket count | E5 parser, normaliser | Tree, values, navigate and back | Tree control; VBA vs Python parser | 10 representative formulas incl. INDIRECT/OFFSET shown as "not evaluated" |
| M6 | Find a Tool + cleanup pack | M2 registry | Search, favourites, trim / text→number / case / fill blanks | Scope creep | "values" finds both commands, clearly distinguished; colleague reverses three commands unaided |

Later: Deep Scan, Audit hardening (ADR requirements), detector extensions, G0 with IT.

## D. Next milestone — M2

**Delivered in Excel**

1. An **Excel Toolkit** ribbon tab, one group (Quick Tools): **Convert to Values**, **Undo Toolkit
   Action**, and **Audit Formulas** (existing Phase B; says clearly when the worker isn't running).
   Labels, tooltips, icons (`imageMso`) and shortcut labels come from a small command registry.
2. **Convert to Values** in VBA, clipboard-free, on the shared command contract below.
3. **Ctrl+Z** immediately after the command restores the formulas via `OnUndo`; **Undo Toolkit
   Action** does the same on demand, under the checks below.
4. Tooltip and docs state the limitation: Excel's earlier Undo history is cleared, as with any macro.

**Not in M2:** other macros, clipboard Paste Values, Find a Tool, generic format/fill/border capture,
any installer, registry or trust change. Ctrl+Shift+V is not taken over.

### Selection behaviour (approved 2026-10-11)

| Selection | Behaviour |
|---|---|
| Single range, or several **non-overlapping** areas | Converted; each area handled separately |
| Overlapping areas | Refused |
| Any hidden or filtered-out row/column **included in the selection** | **Refused**, with the hint to select visible cells only. A filtered sheet is not refused as such: an explicit visible-cells-only selection (Alt+;) is allowed. Hidden cells are never converted silently |
| Merged cells, dynamic-array spill ranges (anchor or spilled cells), legacy array or data-table formulas | Refused (reliable conversion unproven) |
| Locked formula cells on a protected sheet | Refused before anything is written |
| No formulas in the selection | Message; nothing changed; no Undo record |
| Above the size limit (provisional 100,000 cells; a configurable safeguard, revisited after M2 timings) | Refused, never truncated |

### Undo contract (M2)

- **What changes, what is recorded.** Only cells that held a formula are written. Constants in the
  selection are never written, so they cannot be retyped and need no restore — this is how an
  original constant is distinguished from a converted formula. The record holds, per converted cell,
  the address, the original `Formula2`, and the value and type written.
- **Exactness.** A converted cell must hold exactly the formula's prior result: same `VarType` and
  value of `Value2` (text stays text, e.g. `"0012"` and `"1/2"`; dates stay serial numbers; booleans
  and error values unchanged). A restored cell must hold the identical `Formula2` text and be a
  formula again. Number formats must be unchanged throughout (verified, not captured). Matching
  displayed values is **not** sufficient.
- **Identity.** The record holds the `Workbook` and `Worksheet` objects plus the workbook `FullName`
  and sheet name at capture. Restore is refused if the workbook was closed, renamed or saved under
  another name, or the sheet was deleted or renamed.
- **Apply, verify, roll back.** After writing, every converted cell is re-read and checked. On any
  write or verification failure, cells already written get their formulas back and are re-verified.
  If that rollback verifies, the user is told nothing changed; if not, the user gets an
  **incomplete** message listing the affected cells, no Undo is registered, and success is never
  claimed.
- **Registration and invalidation.** On success the record replaces any previous one and
  `Application.OnUndo "Undo Convert to Values"` is registered. The record is cleared after a
  successful restore, replaced by the next toolkit modifying command, and lost when the add-in
  unloads or Excel closes (memory only; no workbook data written to disk).
- **Ctrl+Z versus Undo Toolkit Action.** Both run the same restore procedure with the same checks.
  Excel withdraws our Ctrl+Z entry as soon as the user does anything undoable, and the add-in cannot
  observe that; from then on the explicit button is the only path. The button is enabled only while a
  record exists, and its tooltip names the command, sheet, range and time.
- **Conflict check.** Before restoring, every converted cell must still hold exactly what was
  written (no formula, same type and value). Otherwise restore is refused, nothing is written, and the
  first conflicting cell is named. Cells outside the converted set are neither checked nor touched.
- **Partial restore.** A restore that fails part-way re-verifies and reports exactly which cells were
  and were not restored; the record is then marked unusable.

### How it is shown to work (results in `docs/validation/m2-convert-to-values.md`)

- **Type fixtures:** formulas returning numbers, `"0012"`, `"1/2"`, a date, TRUE, `#N/A`, `#DIV/0!`
  and empty text; constants `'0012`, `'1/2`, numbers and text that must stay untouched; custom number
  formats. Each checked by `VarType` and value after conversion, and by `Formula2` after restore.
- **Selections:** single range, three areas, overlapping (refused), filtered and hidden rows
  (refused), merged, spill and array cells (refused), protected sheet (refused), no formulas.
- **Failure:** a forced mid-operation failure rolls back and verifies; the incomplete path is
  reported, never shown as success.
- **Undo:** **real Ctrl+Z keypresses** (programmatic Undo doesn't fire `OnUndo`); Undo Toolkit Action
  after other work; refused after an edit to a converted cell; refused after the workbook is closed or
  renamed or the sheet renamed; button disabled once the record is used or replaced.
- **Ribbon:** every button clicked by hand in the packaged `.xlam`; the build verified to contain
  `customUI14.xml`, valid against the schema, with callbacks resolving.
- **Performance:** end-to-end time from click to Undo registered (validate, capture, apply, verify,
  register) at 1 / 1,000 / 10,000 / 100,000 cells, with restore timed separately; target under 1 s
  for each at 10,000 cells, using bulk reads and writes wherever they cannot change a type.
- PERSONAL.XLSB SHA-256 unchanged; only disposable workbooks used.

## E. Decisions (reviewer verdicts on `0d5d543`, 2026-10-11)

| # | Decision | Verdict |
|---|---|---|
| 1 | Inject `customUI14.xml` into the `.xlam` after the scripted build (dev machine only; AccessVBOM already on) | **Approved** — reproducible build; packaged XML and callbacks validated |
| 2 | Load the dev `.xlam` with File → Open each session, not the Add-ins dialog | **Approved** for the scoped test; not a distribution solution |
| 3 | Convert to Values first (today's Ctrl+Shift+V behaviour, clipboard-free); clipboard Paste Values in M3 | **Approved** |
| 4 | Ribbon only in M2; Ctrl+Shift+V stays with PERSONAL.XLSB until the replacement passes | **Approved** |
| 5 | Manual checklist plus scripted before/after comparison in a disposable Excel instance | **Approved conditionally** — disposable workbooks and processes; live testing separately authorised |
| 6 | One-level Undo | **Approved** — lifecycle, invalidation and conflicts as defined in §D |
| 7 | Separate results workbook for M4 | **Approved provisionally** — a richer window is not ruled out |
| 8 | Delete Names | **Deferred** — stays disabled until broader workbook safety and recovery are established (restoring name definitions alone does not prove recovery) |
| 9 | External Links Quick Scan technology | **Amended** — VBA, Python or hybrid chosen from an M4 benchmark of coverage and speed |

Selection behaviour approved with the visible-cells-only clarification (2026-10-11). Still open: explicit
authorisation of the scoped M2 live tests (not deployment approval).

Not decisions now: Excel-DNA, installer method, worker auto-start — unchanged until IT answers the
questions in `docs/architecture/deployment-decision.md`.

## Reconciliation of the two assessments

Both reached the same conclusions (M1 done; ribbon and Undo framework are the gap; M2 first). Points
settled against the code:

- **Taken from the second assessment:** the E3 journal does not capture number formats; the macro
  count is 17 on 16 hotkeys; results workbook for M4. (Its proposal to release Delete Names as a confirmed custom-Undo
  command was later deferred by the reviewer — E8.)
- **Corrected in the second assessment:** it described clipboard paste as Ctrl+Shift+V's legacy
  meaning (the macro actually converts in place); it expected Ctrl+Z to refuse after an intervening
  edit (only the explicit command can); it took over Ctrl+Shift+V within M2 (deferred until the
  command passes); its test list lacked filtered rows, spills, merged cells and real keypresses; it
  described identity as pid + workbook name and listed the one-shot fallback as open.
