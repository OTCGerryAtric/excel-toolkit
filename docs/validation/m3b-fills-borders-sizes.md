# M3b validation — fills, Grey Outline, row height, column width, best fit

Status: **offline implementation only (2026-10-12).** Offline tests pass; nothing has been built or run in
Excel. Live tests need separate scoped approval. M3c not started.
Plan: `docs/product-roadmap.md` (M3, Gate 2). Source: `addin/`.

## What M3b adds

| Part | File |
|---|---|
| Fill toggles (Yellow, Green, Light blue, Grey), Grey Outline, Row height (15 ↔ 4.5), Column width (2 → 8.43 → 12), Best fit (multiple of 5 from the active cell); legacy colours, borders, sizes and decisions | `addin/vba/TkStyle.bas` |
| Undo record for fills, borders, row heights and column widths: captures only the changed properties, split into uniform blocks as in M3a | `addin/vba/TkStyleSnapshot.cls` |
| Ribbon: Format ▸ Fill menu (Q E W Z keytips), Outline button (G), Size menu (R C K) | `addin/customUI/customUI14.xml`, `TkMain.bas` |
| Scripted live check (prepared, **not run**) | `addin/live/m3b_live_check.py` |

**No keyboard shortcuts for the M3b commands yet.** Keys move one command at a time after the command
passes live, so the toggle still binds only Ctrl+Shift+I/O/P/M and PERSONAL.XLSB keeps Q E W Z G R C K.

## Legacy behaviour kept, and the one fix

| Command | Legacy rule (kept exactly) | Change |
|---|---|---|
| Fills | If the selection's `Interior.Color` is the colour, turn it off; otherwise apply the colour (a mixed selection gets the colour). RGB(255,255,210) / (145,245,190) / (221,235,247) / (242,242,242). Only `Interior.Color` is set, so an existing pattern stays | **Off = No Fill**, not white (roadmap known weakness 4) |
| Grey Outline | If the top edge has no line: left, top, bottom, right, inside vertical, inside horizontal set to continuous, theme colour 1, tint −0.249946592608417, thin (same order). Otherwise, including a mixed top edge, remove all six. Applied to the whole selection at once, as the macro does | — |
| Row height | 15 → 4.5, anything else (including mixed) → 15 | — |
| Column width | 2 → 8.43 → 12 → 2; anything else → 2 | — |
| Best fit | `CColWidth As Integer = Round(ActiveCell.ColumnWidth + 5) / 5`; width = CColWidth × 5 for the selection (both roundings are banker's, as in the macro: 8.43 → 15, 7.5 → 10, 12.5 → 20) | Over 255 (an active column wider than 252.5) is refused instead of a run-time error |

## Undo design

The same contract as M2/M3a: validate → capture → apply → verify → roll back on failure → one-level undo,
identity and conflict checks, and stage-based reporting.

- **Fill:** pattern, colour (RGB, or theme colour + tint) and, for patterned fills, the pattern colour
  (automatic, RGB or theme + tint). Gradient fills are refused.
- **Borders:** every edge of every selected cell (style, weight, colour: automatic, RGB, or theme + tint).
  Also captured is the facing edge of the cells **around** the selection: the row above, the row below
  and the columns either side. They share the selection's outside edges and are restored last.
- **Row heights / column widths:** the **whole affected rows/columns** (column A of the rows, or row 1
  of the columns), not just the selected cells. **Hidden rows/columns are refused.** Setting a height
  unhides a row, and the height it had before it was hidden can't be read, so an exact undo would be
  impossible. A visible-cells-only selection around hidden rows is allowed.
- **Conflict check:** every captured block (including the neighbouring edges) must still hold what the
  command left there.
- **Protection:** fills and borders need *Format cells*, heights *Format rows*, widths *Format columns*.
- **Limit:** more than 50,000 uniform blocks is refused, as in M3a.

## Decisions for review

1. **Hidden rows/columns refused** for the size commands. The legacy macro unhides them; exact undo isn't
   possible.
2. **Restored row heights become fixed heights.** Excel exposes no "custom height" flag, so a row that was
   auto-sized to its wrapped text comes back at the same height but no longer auto-grows. Standard-height
   rows report `UseStandardHeight` from the value, so they are unaffected.
3. **M3a's `TkFormatSnapshot` is left untouched** (accepted and live-tested). The new class repeats its
   block-splitting; merging the two is deferred to a later tidy-up.
4. **No shortcuts in M3b** (see above).

## Questions only the live run can answer

1. **ThemeColor on RGB colours.** If Excel raises an error when reading `ThemeColor` on several RGB cells,
   the capture can't rule out a hidden mix of RGB and theme colours of the same shade. It therefore splits
   to single cells: exact, but slow for large RGB areas. The common "fill off" case (10,000 toolkit-yellow
   cells) would take 10,000 blocks, and more than 50,000 RGB cells would be refused. The diagnostics record
   what Excel actually reports; if it is safe, this collapses to one read per block.
2. **Shared edges.** Does Excel report the border between two cells from both sides? The design assumes
   yes, and the neighbouring-edge capture makes undo exact either way, if the read-back agrees. The
   equivalence fixtures include neighbouring borders, and the diagnostics record the answer.
3. **Multi-area selections.** The toolkit applies fills and sizes area by area; the legacy macro sets the
   whole selection. The equivalence cases compare them on a two-area selection.
4. **Read-back of 4.5 / 8.43** at the owner's display scaling (the legacy rules compare exact values).
5. **Merged cells** inside a mixed selection, and the three ribbon icons (`CellFillColorPicker`,
   `BordersAll`, `ColumnWidth`).

## Offline evidence

`tests/addin/test_addin_offline.py` (81 tests in this file, 12 of them new; full offline suite 174 passed):
- fill colours and decision parsed from each legacy macro; "off" is No Fill and white is never used;
- the outline's six edges, their order, theme colour, tint, weight and decision match the macro;
- row and column values and best fit's expression (including the Integer rounding) match;
- size undo captures whole rows/columns and refuses hidden ones;
- border undo captures the neighbouring edges after the selection;
- gradients are refused, and theme markers are never compared with `=` (that is a Type mismatch on
  numbers);
- the new snapshot class provides the undo-record interface;
- plus all the M2/M3a structural checks over the new modules.

Deliberately changing a legacy value in `TkStyle.bas` (a colour, the 12 width, best fit's +5, the outline
tint) makes the tests fail.

Not provable offline: compilation and Excel behaviour.

## Live tests (pending approval)

1. **Build** in a separate hidden Excel, then a **VBA Compile** check in a disposable instance.
2. **Scripted** (`addin/live/m3b_live_check.py --perf`, disposable Excel, legacy macros imported from
   `baseline/vba`):
   - diagnostics;
   - fills (4 colours × 3 selections × 6 start states: none, same colour, other RGB, theme + tint,
     patterned, mixed);
   - outline (3 selections × 5 start states, including neighbouring borders, × 1–2 presses);
   - sizes (including whole rows/columns and multi-area; best fit from 7 active widths), each compared
     property by property with the legacy macro and undone exactly;
   - refusals (hidden rows/columns, gradient), merged cells, conflicts (including an edited neighbouring
     edge), identity, protection (each permission);
   - forced apply/restore failures; too many blocks;
   - timings (uniform, RGB fill off, theme, alternating, whole column with exceptions, outline 100 × 100,
     whole-column row height, whole-row column width).
3. **Manual** (owner): ribbon, and **real Ctrl+Z once per new undo kind** (fill, outline, row height,
   column width). PERSONAL.XLSB hash and security checks as before.
