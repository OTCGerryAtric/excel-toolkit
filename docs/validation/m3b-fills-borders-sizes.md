# M3b validation — fills, Grey Outline, row height, column width, best fit

Status: **build, VBA compilation and automated live checks passed: 39/39 on the third run (2026-10-12),
after two add-in defects found by the first runs were fixed.** Automated results reviewed and accepted;
**manual checklist passed 13/13 (2026-10-12)**; PERSONAL.XLSB unchanged, no security events. Not approved for merge. M3c not started.
Plan: `docs/product-roadmap.md` (M3, Gate 2). Source: `addin/`.

## What M3b adds

| Part | File |
|---|---|
| Fill toggles (Yellow, Green, Light blue, Grey), Grey Outline, Row height (15 ↔ 4.5), Column width (2 → 8.43 → 12), Best fit (multiple of 5 from the active cell); legacy colours, borders, sizes and decisions | `addin/vba/TkStyle.bas` |
| Undo record for fills, borders, row heights and column widths: captures only the changed properties | `addin/vba/TkStyleSnapshot.cls` |
| Ribbon: Format ▸ Fill menu (Q E W Z keytips), Outline button (G), Size menu (R C K) | `addin/customUI/customUI14.xml`, `TkMain.bas` |
| Scripted live check | `addin/live/m3b_live_check.py` |

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

## Undo design (as revised by the live runs)

The same contract as M2/M3a: validate → capture → apply → verify → roll back on failure → one-level undo,
identity and conflict checks, and stage-based reporting.

- **Fill:** pattern, colour (RGB, or theme colour + tint) and, for patterned fills, the pattern colour
  (automatic, RGB or theme + tint). A uniform area is one read; a mixed one is split into uniform blocks.
  Gradient fills are refused.
- **Borders:** every edge of every selected cell (style, weight, colour: automatic, RGB, or theme +
  tint), **read cell by cell**. Adjacent cells share an edge, so each edge is held once: every cell's top
  and left, then the last row's bottoms and the last column's rights. Those last two are also the facing
  edges of the cells around the selection. Equal neighbours are joined into runs and rectangles, so
  uniform areas restore in a few writes. **Over 50,000 cells is refused** (for example, a whole column).
- **Row heights / column widths:** the **whole affected rows/columns** (column A of the rows, or row 1
  of the columns), not just the selected cells. **Hidden rows/columns are refused.** Setting a size would
  unhide them, and their size before hiding can't be read. A visible-cells-only selection around hidden
  rows is allowed.
- **Conflict check:** every restored property must still hold what the command left there; borders are
  checked cell by cell.
- **Protection:** fills and borders need *Format cells*, heights *Format rows*, widths *Format columns*.
- **Limit:** more than 50,000 blocks is refused, as in M3a.

## Decisions (reviewer, 2026-10-12)

1. **Hidden rows/columns refused** for the size commands — *approved*.
2. **Automatically sized rows** — *conditionally approved; see the separate limitation below*.
3. **M3a's `TkFormatSnapshot` left untouched** — *approved*. (It splits with `Rows(n)`/`Columns(n)`,
   which the M3b live run showed count rows/columns rather than cells. M3a never relies on that count,
   and its live run passed with mixed and whole-column selections, so it is not affected.)
4. **No shortcuts in M3b** — *approved*.

### Limitation: automatically sized row heights (recorded separately, not claimed as exact)

Undo puts back each row's **numeric height** exactly. It does **not** restore *automatic* sizing. A row
that Excel had auto-sized to fit wrapped text comes back at the same height, but as a fixed height, so it
no longer grows or shrinks when its text changes. The legacy macro has the same effect when it sets a
height. Excel's object model has no property that reports whether a row is auto-sized; `UseStandardHeight`
only says whether the height equals the sheet's default. So this live run could not establish whether the
state can be detected and refused. **Row-height undo is therefore exact for heights, not for auto-fit
behaviour.** This needs a product decision before release: accept and document it, or investigate
detection further (for example, by comparing a row's height with its auto-fitted height on a copy).

## Live results

Separate hidden Excel processes for build, compile check and scripted checks; nothing installed.

| Item | Result |
|---|---|
| PERSONAL.XLSB | SHA-256 unchanged before, during (read-only, while the owner's own Excel was open) and after; last written 15/06/2026 |
| Security | 0 Defender detections; SentinelOne logged its regular 20-minute bursts of 14 anti-tamper entries (70 across the session), none naming Excel, Python or the add-in; no firewall events; no Excel left running |
| Environment | Scratch venv only; the project `.venv` was not touched again after the earlier `uv run` incident, which is preserved as recorded |

**VBA compilation:** clean on every build (VBE Debug > Compile; no dialogs).

**Diagnostics** (the open questions):

| Question | Answer |
|---|---|
| ThemeColor on an RGB fill | **0** for one cell and for several (no error); **Null** for a mix of a theme colour and an RGB colour of the same shade; −4142 for no fill. So a uniform RGB area is **one block**: removing yellow from 10,000 cells is about 13 ms, and the 50,000 limit does not affect ordinary fills |
| ThemeColor on an automatic border colour | Error (handled: automatic colour is recognised first) |
| Shared edges | **Yes**, for reading and writing: setting one cell's bottom sets the next cell's top, and clearing either side clears both |
| Multi-cell border reads | Excel reports the **first cell's** edge, never "mixed", even "no line" when later cells have one. This was the cause of defect 1 below |
| Read-back of 4.5 and 8.43 | Exact |

**Run 1: 34/37.** Investigated before re-running:
1. *Add-in defect — borders.* The capture relied on multi-cell border reads reporting "mixed", which
   Excel never does (above). Mixed borders were taken as uniform, and undo reported **"incomplete"**
   (never a false success). **Fix:** borders are read cell by cell, as described above.
2. *Add-in defect — merged cells.* A selection containing a merged cell was **refused** (safely). The
   second half of a split, taken with `Columns(n)`, was a range whose count is its number of columns (1),
   which the capture read as a single unreadable cell. **Fix:** blocks are split by corner cells
   (`Range(Cells(r1, c1), Cells(r2, c2))`).
3. *Harness defect.* The "neighbouring edge" conflict test edited A1, which is diagonal to the selection
   and not an edge undo restores, so undo correctly proceeded. Now edits C1's bottom (shared with C2's top).

**Run 2: 37/39.** The outline on a **two-area** selection was **rolled back** (safely). The legacy macro
applies the edges to the whole selection at once, which does not leave each original block of the second
area uniform, but the post-change record assumed it would. **Fix:** the post-change border state is
recorded cell by cell. The conflict check also now runs with screen updating off: at 40,000 cells it took
about 9 s before.

**Run 3: 39/39** (`addin/live/results/m3b_live_20261012_114622.json`):

| Area | Result |
|---|---|
| Fills vs legacy | 72 cases (4 colours × 3 selections × 6 start states: none, same colour, other RGB, theme + tint, patterned, mixed) match **property by property**; the 12 switch-offs give No Fill where the macro gives white |
| Fill undo | Every case undone exactly (RGB, theme + tint, patterns, mixed) |
| Outline vs legacy | 30 cases (3 selections incl. two areas × 5 start states incl. neighbouring borders, thick coloured box and mixed edges × 1–2 presses) match **edge by edge** across the fixture, neighbours included |
| Outline undo | Every case undone exactly, including the neighbours' facing edges |
| Sizes vs legacy | 50 cases (row height and column width incl. whole rows/columns and two areas; best fit from 7 active widths) match |
| Size undo | Every case undone exactly (whole affected rows and columns) |
| Refusals | Hidden row / hidden column in the selection refused, nothing changed; visible rows around a hidden one applied and undone (row stays hidden); gradient fill refused; whole-column outline (over 50,000 cells) refused, nothing changed; 60,000 alternating fills refused |
| Merged cells | Mixed selection with a merged cell: fill and outline applied and undone exactly |
| Conflicts | Undo refused after a fill edit, an edit to the shared edge above the selection, and a width edit; refused after a sheet rename |
| Protection | Each of Format cells / rows / columns: refused when not allowed, applied and undone when allowed |
| Failures | Forced failure after two outline edges, and after the first area of a row-height change: rolled back exactly. Forced restore failure: reported as incomplete, record unusable |

**Performance** (third run; ms; apply = click to undo registered; undo wall includes the conflict check):

| Case | Blocks | Apply | Undo |
|---|---|---|---|
| Fill on, 10,000 uniform (no fill) | 1 | 13 | 23 |
| **Fill off, 10,000 uniform RGB yellow** | 1 | **13** | **25** |
| Fill on, 10,000 uniform theme colour | 1 | 13 | 26 |
| Fill on, whole column uniform | 1 | 13 | 23 |
| Fill on, whole column, 10 exceptions | 175 | 30 | 117 |
| Fill on, 10,000 alternating rows (worst case) | 10,000 | 1,394 | 3,813 |
| Outline on, 10,000 cells (100 × 100) | 3 | 821 | 917 |
| Outline off, 10,000 outlined cells | 3 | 937 | 920 |
| Outline on, 40,000 cells (100 × 400) | 3 | 3,360 | 3,608 |
| Column width, whole row (every column) | 1 | 113 | 183 |
| Row height, whole column (every row) | 1 | 7,101 | 14,719 |

Ordinary fills on uniform selections, including RGB fill removal and whole columns, take about 13 ms to
apply and 25 ms to undo. **Grey Outline is the slow command:** reading every cell costs about 0.8–0.9 s
per 10,000 cells, so it misses the 50 ms aspiration well short of the 50,000-cell limit. Worst-case
alternating fills take 1.4 s to apply and 3.8 s to undo. **Row height on a whole column is slow** (7 s /
15 s): reading and setting the height of all 1,048,576 rows costs Excel seconds, and the run-to-run
variance was large (run 1: 4.7 s / 2.7 s). These are recorded as optimisation candidates. They are
correct, but slow for very large selections.

## Offline evidence

`tests/addin/test_addin_offline.py` (82 tests in this file; full offline suite 175 passed):
- fill colours and decision parsed from each legacy macro; "off" is No Fill and white is never used;
- the outline's six edges, their order, theme colour, tint, weight and decision match the macro;
- row and column values and best fit's expression (including the Integer rounding) match;
- size undo captures whole rows/columns and refuses hidden ones;
- border undo reads every cell and each shared edge once, with the cell limit;
- the snapshot splits ranges by corner cells only (no `Offset`, `Resize`, `Rows(n)`, `Columns(n)`);
- gradients are refused, and theme markers are never compared with `=`;
- the new snapshot class provides the undo-record interface;
- plus all the M2/M3a structural checks.

Deliberately changing a legacy value in `TkStyle.bas` makes the tests fail.

## Review of the automated results (2026-10-12, at `33aef76`)

39/39 accepted; the border, merged-selection and multi-area fixes are accepted. Uniform-fill speed is
satisfactory. Grey Outline speed is acceptable for M3b, with large-selection optimisation tracked.
Whole-column row height needs a future safeguard or optimisation. Numeric row-height restoration is
accepted for M3b development, with automatic sizing explicitly **not** restored. The border-cell and
block safeguards stay. No change to M3a, and no M3b shortcuts.

## Manual checklist (approved 2026-10-12)

Owner's normal Excel (PERSONAL.XLSB loads as usual). File → Open the development
`addin/build/ExcelToolkit.xlam` (the build from run 3), new blank workbook. Nothing installed.
PERSONAL.XLSB SHA-256 before: `A4F5900D…BB81EB7D`.

| # | Step | Expected | Result |
|---|---|---|---|
| 1 | Excel Toolkit tab | Format group: Number formats, **Fill** menu (Yellow, Green, Light blue, Grey), **Grey Outline**, **Size** menu (Row height, Column width, Best fit), each with a tooltip; note any blank icon | **Pass** (2026-10-12) |
| 2 | Select B2:D6; Fill ▸ Yellow | Yellow; Undo Toolkit Action enabled | **Pass** (2026-10-12) |
| 3 | **Ctrl+Z** | No fill again (gridlines visible); button greyed | **Pass** (2026-10-12) |
| 4 | Fill ▸ Yellow twice on B2:D6, then **Ctrl+Z** | Second click gives **No Fill** (gridlines visible, not white); Ctrl+Z brings the yellow back | **Pass** (2026-10-12) |
| 5 | Select F4; Home ▸ Borders ▸ **Thick Box Border**. Select B2:E6; Grey Outline | Thin grey grid on B2:E6; F4's left side becomes thin grey | **Pass** (2026-10-12). The owner first expected an outline only; the inside lines are the legacy macro's behaviour (inside vertical and horizontal), confirmed identical in the automated run |
| 6 | **Ctrl+Z** | Grid gone; **F4's thick box complete again**, left side included | **Pass** (2026-10-12) |
| 7 | Select B10; Size ▸ Row height; then **Ctrl+Z** | Row 10 changes (15 → 4.5, or → 15); Ctrl+Z restores it | **Pass** (2026-10-12) |
| 8 | Select C1; Size ▸ Column width; then **Ctrl+Z** | Column C changes (8.43 → 12); Ctrl+Z restores it | **Pass** (2026-10-12) |
| 9 | Click B1; Size ▸ Best fit; then Undo Toolkit Action **button** | Column B widens to 15; the button restores it and greys out | **Pass** (2026-10-12) |
| 10 | Fill ▸ Green on B2:D6; change C3's fill by hand; Undo Toolkit Action | Refused with a message; C3 keeps its hand-set fill | **Pass** (2026-10-12) |
| 11 | Toolkit shortcuts **off**: on a spare cell, Ctrl+Shift+Q, G, R, K | PERSONAL.XLSB macros (Undo Toolkit Action does not light up) | **Pass** (2026-10-12) |
| 12 | Toolkit shortcuts **on**: the same keys; then Ctrl+Shift+I; then switch off | Q, G, R, K still PERSONAL.XLSB; I runs the toolkit Number format | **Pass** (2026-10-12) |
| 13 | Close Excel, **Don't Save** everywhere | PERSONAL.XLSB unchanged (hash); no security events | **Pass** (2026-10-12) |

After the checklist: PERSONAL.XLSB SHA-256 unchanged (`A4F5900D…BB81EB7D`, last written 15/06/2026); 0 Defender
detections; SentinelOne one routine burst of 14 anti-tamper entries (12:40), none naming Excel, Python or the
add-in; no Excel left running.

## Owner request: use the workbook's default sizes (proposed change, not implemented)

The legacy macros use fixed numbers for Excel's usual defaults: row height **15** (row toggle 15 ↔ 4.5)
and column width **8.43** (cycle 2 → 8.43 → 12). Some workbooks have different defaults (another font or
size, or a changed standard width), and the owner would like these values to be **the workbook's own
default** instead: the sheet's `StandardHeight` and `StandardWidth`. This changes legacy behaviour, so it
is listed here for a decision, not made in M3b:
- **Row height:** standard height ↔ 4.5 (instead of 15 ↔ 4.5).
- **Column width:** 2 → standard width → 12 → 2 (instead of 2 → 8.43 → 12).
- **Best fit** is unchanged (it is based on the active cell's width).
- The standard width is often only *around* 8.43 (for example 8.38 or 9.14, depending on the workbook's
  font and the screen's pixel rounding). An exact comparison with 8.43, as in the legacy macro, then never
  matches, so the cycle jumps straight back to 2. The change would compare with the sheet's own standard
  width, allowing for pixel rounding, rather than with a fixed number.
- On a workbook whose defaults are 15 and 8.43, nothing changes. The legacy-equivalence checks would then
  compare against the documented change on other workbooks.
