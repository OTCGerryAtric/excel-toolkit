# M3d — native Paste Values experiment

Status: **prepared, not run (2026-10-12).** Nothing has been run in Excel for M3d. Running the experiment
needs separate scoped approval. No custom Paste Values fallback will be written until the experiment has
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

**For review:** what Ctrl+Shift+V should eventually do when shortcuts migrate. It could convert in place
(the legacy behaviour, now Convert to Values) or paste the clipboard as values (M3d). The roadmap already
asks for three Ctrl+Shift+V behaviours to be kept distinct: Excel's own (if this build has one), the
PERSONAL.XLSB assignment, and the toolkit's.

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
| Controls: object-model write; `Range.PasteSpecial` | History cleared? (expected yes) |
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

`tests/addin/test_addin_offline.py`: 256 passed. Two new checks:
- The probe is ASCII with CRLF. Its native function uses `ExecuteMso "PasteValues"` and never
  `PasteSpecial`, which appears only in the labelled control. The probe is not in the add-in build, and the
  add-in has no Paste Values code yet.
- The harness uses `DispatchEx` only and does not open the add-in. It only reads PERSONAL.XLSB's bytes to
  hash them, never saves, and is marked "DO NOT RUN until authorised".

## Approval requested

Run `m3d_native_paste_experiment.py` in a disposable Excel, hidden first and visible once only if Excel
needs a window. Report the results before any further M3d work.
