# M3d — native Paste Values experiment

Status: **approved in principle (review of `7bd3025`), amended, not yet run (2026-10-12).** The amendment
adds a no-paste mechanism control and pass / fail / inconclusive verdicts (below). The run waits for the
machine owner's explicit clipboard consent. No custom Paste Values fallback will be written until the experiment has
been reviewed. The add-in is unchanged: the probe is a standalone module that is never built into it.
Plan: `docs/product-roadmap.md` (M3, "Clipboard Paste Values (M3d)").

## What M3d is for, and what the legacy macro does

The legacy **Ctrl+Shift+V** (`Paste_As_Values` in PERSONAL.XLSB) does **not** paste the clipboard:

```vba
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, ...
Application.CutCopyMode = False
```

It copies the **selection itself** and pastes its values back on top. That converts the selection to values
in place, the job M2's Convert to Values already does with exact undo, and it **empties the clipboard**.
So clipboard Paste Values (paste what the user copied, as values, into the selection) is a **new**
capability in M3d, not a port of that macro.

**Reviewer decision (product design):**
- **Convert to Values** is M2's existing in-place conversion command.
- **Paste Values** is M3d's proposed clipboard-to-destination command.
- The preferred eventual **Ctrl+Shift+V** is clipboard Paste Values. The shortcut is **not** migrated now.

## The question

Can the toolkit use **Excel's own Paste Values**, run from VBA with `CommandBars.ExecuteMso "PasteValues"`,
instead of writing cells itself? It is adopted only if it is shown live to:
1. **keep the clipboard**: the copy marquee stays and a second paste works;
2. give a **native Undo** for the paste (Ctrl+Z puts the destination back);
3. **keep the earlier Undo history**: the user's actions before the paste can still be undone;
4. handle **single and multi-area destinations** with sensible behaviour.

Any object-model write from VBA clears Excel's Undo history (E1/E3). The hope is that a built-in command
run through `ExecuteMso` behaves like the user clicking it, and does not.

## Method (`addin/live/m3d_native_paste_experiment.py`, not run)

Boundaries are as for M2–M3c:
- a **separate disposable Excel** (`DispatchEx`: no XLSTART, no PERSONAL.XLSB, no add-ins);
- workbooks created and closed without saving;
- the probe `addin/live/m3d_paste_probe.bas` imported into a disposable workbook, as the legacy macros were
  for M3a–M3c;
- the add-in is not opened, and nothing is installed;
- PERSONAL.XLSB is hashed before and after.

The probe's functions:

| Function | What it does |
|---|---|
| `Probe_NativePasteValues` | `GetEnabledMso("PasteValues")`, then `ExecuteMso "PasteValues"`; reports `CutCopyMode` before and after |
| `Probe_NativeBold` | A user-interface command (`ExecuteMso "Bold"`) that creates the **earlier** Undo entry |
| `Probe_NativeUndo` | Excel's own Undo command, one step (`ExecuteMso "Undo"`) |
| `Probe_UndoState` | Whether Undo is available, and the Undo list if Excel still exposes it (classic control 128) |
| **Controls:** `Probe_ControlObjectModelWrite`, `Probe_ControlPasteSpecial` | An ordinary VBA write and the legacy `Range.PasteSpecial`, both expected to **clear** the history. They show what a cleared history looks like, so a "kept" result can be trusted |

**Mechanism control (amendment, runs first, no paste).** Bold J1 with `Probe_NativeBold`, read the Undo
state, then `Probe_NativeUndo`. It passes only if the Bold applied, Undo was available, and Undo removed the
Bold. If it fails, the probe cannot build or reverse a native Undo entry, and **nothing is concluded** about
whether Paste Values keeps the earlier history.

**Verdicts.** Each criterion is reported as:
- **pass**: measured, and the property holds in every core destination (single cell, same size, twice the
  size);
- **fail**: measured with trustworthy controls, and the property does not hold. This is a genuine failure;
- **inconclusive**: the measurement cannot be trusted or did not happen. Causes: the mechanism control
  failed; a negative control did **not** show a cleared history; the earlier Bold did not apply; native
  Paste Values did not run (disabled or an error, for example in a hidden Excel); or the scenario did not
  run. The two Undo criteria depend on the controls; the values and clipboard criteria do not.

An inconclusive mechanism control or a paste that did not run sets "visible retry suggested".

In each case the source A1:B3 holds a formula, number formats, text, a date and TRUE; the destination area
starts as "old". The steps are:
1. Set up the sheet (object-model writes, done first because they clear the history).
2. Bold J1 with the user-interface command (the earlier history).
3. Copy the source.
4. Select the destination and run the native Paste Values.
5. Record the values, formulas, formats, copy mode and Undo state.
6. Undo once: the paste should go and J1 stay bold.
7. Undo again: J1's bold should go, which shows the earlier history survived.

| Scenario | What is recorded |
|---|---|
| Mechanism control: native Bold, Undo state, native Undo (no paste) | Bold applied; Undo available; Undo removed it |
| Negative controls: object-model write; `Range.PasteSpecial` | History cleared? (expected yes) |
| Destination: single cell; same size; twice the size (Excel repeats the copy) | Values correct, no formulas or formats pasted; clipboard kept; Undo reverses the paste; earlier history kept |
| Destination: two areas (D1,G1); two blocks (D1:E3,G1:H3) | What Excel does (it normally refuses with a message) and any message text |
| Source: two areas (A1,A3) | What Excel does |
| After **Cut** instead of Copy | Whether Paste Values is offered |
| Two pastes from one copy | Clipboard kept between pastes |
| Source in another workbook | Values and Undo |
| Clipboard holding text from outside Excel; empty clipboard | Whether Paste Values is offered, and what it pastes |
| Protected sheet; filtered destination (hidden rows) | Refusal or message; which rows receive values |

Any Excel message box is recorded and dismissed with OK. A watchdog ends the disposable Excel if it stalls.
The run writes a JSON results file with a summary of the four criteria.

## What automation cannot show (a later manual stage)

- **Invocation from the ribbon or a keyboard shortcut.** Automation can only call `Application.Run`. Excel
  may treat a macro started from the ribbon or a key differently for Undo.
- **A real Ctrl+Z key press,** rather than Excel's Undo command run from VBA.
- Whether this Excel build has its **own** Ctrl+Shift+V (paste values).

If the automated results are promising, a short owner checklist would cover these. It would use a
disposable workbook and the probe imported into it, without the add-in, and would need its own approval.

## Risks and precautions

- **The clipboard is shared** with the owner's Windows session. The run overwrites it, so the owner should
  not copy anything during the run. The clipboard is emptied at the end.
- **A hidden Excel** may not run ribbon commands (`ExecuteMso`). The run records this. One rerun with
  `--visible` would then be needed: a disposable Excel window appears and closes when the run ends.
- **Message boxes** are dismissed automatically and recorded. A dialog that OK cannot close ends the run
  through the watchdog.
- Same security checks as before: Defender detections, SentinelOne Operational and Firewall, no Excel left
  running, PERSONAL.XLSB SHA-256.

## Offline evidence

`tests/addin/test_addin_offline.py`: 262 passed. The amendment adds six checks on the verdict logic
(synthetic results): all pass; a failed mechanism control makes the Undo criteria inconclusive, not failed;
a negative control that keeps the history makes them inconclusive; a genuine failure reads fail; a paste
that did not run reads inconclusive and suggests the visible retry; an empty run is inconclusive throughout.
The harness check also confirms that the mechanism control runs first. The earlier two checks:
- The probe is ASCII with CRLF. Its native function uses `ExecuteMso "PasteValues"` and never
  `PasteSpecial`, which appears only in the labelled control. The probe is not in the add-in build, and the
  add-in has no Paste Values code yet.
- The harness uses `DispatchEx` only and does not open the add-in. It only reads PERSONAL.XLSB's bytes to
  hash them, never saves, and is marked "DO NOT RUN until authorised".

## Approval and run conditions

Approved in principle. Before running, the machine owner must confirm explicitly that the Windows clipboard
may be overwritten and emptied, and that they will not copy anything during the run. Then run
`m3d_native_paste_experiment.py` in a disposable Excel, hidden first, with one visible retry only if needed.
The usual limits still apply:
- the scratch Python environment;
- no installation;
- no trust or registry changes;
- no PERSONAL.XLSB changes;
- no audit worker;
- stop on any unexpected security event.

The complete results, including inconclusive measurements and clipboard behaviour, are reported before any
implementation is proposed. Not authorised: merging M3d, a custom fallback, shortcut migration, further
milestones.
