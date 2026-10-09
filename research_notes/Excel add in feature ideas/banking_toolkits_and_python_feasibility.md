# Banking/Consulting Excel Toolkits (Formatting & Shortcuts) and Python Add-in Feasibility

Research date: 2026-10-09. Note: Macabacus help-centre articles (help.macabacus.com) now redirect to a login wall, so most detail on Macabacus comes from search-result extracts of its docs and blog. The official Macabacus PDF cheatsheet URL returned 404 when fetched.

## Q1. Which formatting/shortcut features do banking/consulting toolkits offer, by category?

### Takeaway
Macabacus is the reference implementation for banker-style Excel productivity. It offers keyboard-driven format cycles, AutoColor, flip sign, anchor cycling, IFERROR wrap, Fast Fill, Paste Exact and a multi-slot paintbrush, and all of its shortcuts can be remapped. UpSlide and Power-user are stronger on Excel-to-PowerPoint linking and corporate branding than on cell-level modelling shortcuts. FactSet bundles hotkeys and auto colour-coding from its DealMaven lineage. I found no evidence that Capital IQ's plug-in includes formatting tools; it is a data or formula plug-in. BIWS and WSP sell training and cheat sheets, not formatting add-ins.

### Cited Findings

**Format cycles (number / colour / alignment / font)**
- Macabacus format cycles step through preset formats with repeated presses of one shortcut. For example, pressing the Fill Color Cycle shortcut three times applies a different fill colour on each press, and after the last format the cycle wraps back to the first — [Macabacus Format Cycles (search extract)](https://www.macabacus.com/docs/excel/cycles)
- Macabacus alignment cycles: Horizontal Cycle, Vertical Cycle (top/centre/bottom), and Center Cycle, which includes Center Across Selection as an alternative to merging. Left and Right Indent Cycles add indent on each press up to a maximum set in Settings, then reset to zero. All are also in the Format dropdown on the ribbon — [Macabacus formatting page](https://macabacus.com/excel/formatting)
- Shortcuts are remappable in Macabacus > Settings > Keyboard > Shortcut Manager — [Macabacus formatting page](https://macabacus.com/excel/formatting); [Macabacus keyboard docs](https://www.macabacus.com/docs/excel/keyboard)
- Cycle colours must exist in the user's palette. You build a palette by formatting cells and clicking Import Colors. The docs recommend including modelling colours such as blue for inputs and green for worksheet links — [Macabacus Color Palettes](https://help.macabacus.com/article/766-color-palettes); [Macabacus Colors](https://help.macabacus.com/article/318-colors)
- Font cycle and font tips are covered in the Macabacus blog — [Macabacus blog: Excel font tips and shortcuts](https://macabacus.com/blog/excel-font-tips-and-shortcuts)
- Excel has no built-in keyboard shortcuts for font or fill colours. Excel Campus sells a "Formatting Shortcuts" add-in (with a free version) that binds keys to font colour, fill colour or cell styles. A cell style can include number format, borders, alignment and protection, so changing the style updates every input cell — [Excel Campus](https://www.excelcampus.com/vba/font-fill-color-style-excel-add-in/); [Excel Campus tool page](https://www.excelcampus.com/tools/excel-formatting-shortcuts-add-in-2/)

**Font colour coding / AutoColor (inputs vs formulas vs links)**
- Macabacus AutoColor assigns font colours by cell type: numeric inputs, partial inputs (a formula with a hardcoded number, e.g. =A1+12.34), same-sheet formulas, cross-sheet formulas, cross-workbook formulas, hyperlinks, and data pulls (FactSet or CapIQ formulas). The default scheme follows finance-industry convention and can be customised under Excel > Format > Colors — [Macabacus Colors](https://help.macabacus.com/article/318-colors); [Macabacus Colors (alt)](https://macabacus.com/docs/excel/colors)
- Macabacus Blue-Black Toggle: blue for inputs, black for formulas that reference only the same sheet. Its blue cannot be customised — [Macabacus Colors](https://help.macabacus.com/article/318-colors)
- AutoColor shortcuts reported by a third party: Ctrl+Alt+A for the selection and Ctrl+Alt+S for the whole sheet. These are not confirmed by an official source — [BetterSolutions third-party add-in shortcuts](https://bettersolutions.com/excel/ribbon/shortcut-keys-thirdpartyaddins.htm)
- Ampler offers "Autocolor sheet or selection according to company standards" — [Ampler guide](https://ampler.io/guides-and-tutorials/autocolor-sheet-or-selection-according-to-company-standards/)
- A FactSet brochure markets "keystroke-driven spreadsheet tools and custom hotkeys such as automated color-coding". It also offers hotkeys that run a specified action when cells change, alignment and spacing tools for shapes and logos, and source-linked imports from HTML or PDF — [FactSet Office Integration brochure](https://go.factset.com/hubfs/Website/Resources%20Section/Brochures/office-integration-brochure.pdf)
- FactSet's banker tooling descends from DealMaven (Excel FastTrack, PresLink), acquired in 2007/08. Sources disagree on the year — [BetterSolutions FactSet](https://bettersolutions.com/excel/add-ins/factset.htm)

**Formula manipulation (sign flip, IFERROR/ROUND wrap, anchoring)**
- Macabacus Flip Sign, reported default Ctrl+Alt+' : negates numeric inputs and multiplies formulas by -1 across the selection, e.g. to change the expense sign convention — [Macabacus Formulas & Data (search extract)](https://www.macabacus.com/docs/excel/formulas-data)
- Macabacus Anchor, reported default Ctrl+Shift+L: each press cycles the references in the selected formulas through =A1, =$A$1, =A$1, =$A1. This works on a whole selection, unlike F4, which only works while editing a cell — [Macabacus Formulas & Data](https://www.macabacus.com/docs/excel/formulas-data)
- Macabacus Error Wrap, reported default Ctrl+Shift+E: wraps formulas in IFERROR, and pressing again removes the wrapper. It can prompt for the error value or use a default — [Macabacus Formulas & Data](https://www.macabacus.com/docs/excel/formulas-data)
- Macabacus Wrap Parentheses, Ctrl+Shift+N: wraps the formula in parentheses and opens the cell for editing — [Macabacus Formulas & Data](https://www.macabacus.com/docs/excel/formulas-data)
- Shortcut defaults are configurable and may differ by version, so the Shortcut Manager is the authoritative list — [Macabacus keyboard docs](https://www.macabacus.com/docs/excel/keyboard)

**Fill / paste tricks**
- Macabacus Fast Fill Right (Ctrl+Shift+R) and Fast Fill Down (Ctrl+Shift+D) need no pre-selection. They work out how far to fill from the surrounding data, skip empty spacer columns or rows, and extend date series to labels such as "4Q13" and "2014E" — [Macabacus Formulas & Data](https://www.macabacus.com/docs/excel/formulas-data)
- Macabacus Paste Exact (Ctrl+Alt+E) pastes copied cells and keeps the original formula references. The cheatsheet also lists Paste Duplicate — [Macabacus Formulas & Data](https://www.macabacus.com/docs/excel/formulas-data)
- Macabacus Paintbrush stores up to 8 captured formats, replaced oldest-first — [Macabacus blog: formatting paintbrush](https://macabacus.com/blog/excel-formatting-paintbrush)
- Macabacus has its own undo system, because running code that changes a sheet clears Excel's native undo and redo stacks — [Macabacus Undo/Redo](https://www.macabacus.com/docs/excel/undo-redo)

**Chart formatting / slide export links**
- UpSlide links text, tables and charts from Excel to PowerPoint or Word, exported as text links, dynamic images or editable tables. "Resize and Export" fits the item to a selected PowerPoint placeholder. The Excel Link Manager updates many links at once, groups them by slide or source, and flags links with update issues — [UpSlide Excel-PowerPoint link](https://upslide.net/features/link-excel-powerpoint); [UpSlide support: linking](https://support.upslide.net/hc/en-us/articles/360015592579-Linking-tables-charts-and-text-from-Excel-to-PowerPoint-Word); [Microsoft Marketplace listing](https://marketplace.microsoft.com/en-us/product/web-apps/upslide.upslide)
- Power-user: custom keyboard shortcuts, company-level table-style presets for financial tables in Excel, a "cleaning" menu that finds formatting inconsistencies, Excel-to-PowerPoint links for charts, tables and values, and templates that follow the brand colour scheme. Reviews say it can be buggy or slow. One site lists about $240 per user per year (unverified) — [Capterra](https://www.capterra.com/p/206542/Power-user/); [MakeTechEasier review](https://maketecheasier.com/power-user-addin-powerpoint-excel); [AppSource](https://appsource.microsoft.com/product/office/WA200006804)
- FactSet: charts and visuals with custom colours, fonts and logos for corporate branding, plus 100+ report templates — [FactSet brochure](https://go.factset.com/hubfs/Website/Resources%20Section/Brochures/office-integration-brochure.pdf)
- Macabacus positions itself against UpSlide for Office-wide finance workflows (vendor comparison) — [Macabacus vs UpSlide](https://macabacus.com/lp/macabacus-vs-upslide)

**Capital IQ / BIWS / WSP**
- The S&P Capital IQ plug-in adds an "S&P Capital IQ" ribbon tab and a Formula Builder wizard. The official quick-start pages mention no formatting or colour-cycle tools — [S&P CIQ Screening & Excel Plug-in](https://pages.marketintelligence.spglobal.com/SP-Capital-IQ-Jump-Start-Screener-and-Excel-Plug-in.html)
- Wall Street Prep offers an Excel shortcuts cheat sheet (100+ shortcuts, Windows and Mac). It is a PDF, not an add-in — [WSP Excel shortcuts](https://www.wallstreetprep.com/excelshortcuts/)
- BIWS teaches native Excel shortcuts and claims they save 30-40% of Excel time — [BIWS IB Excel shortcuts](https://breakingintowallstreet.com/kb/excel/investment-banking-excel-shortcuts)
- 2026 competitor "Macabacus alternative" roundups say analysts want replacements for Format Painter, Pro Audit and Quick Save shortcuts. These are vendor-authored and biased. Macabacus pricing is reported inconsistently, from $268 to $441 per user per year — [CT Acquisitions roundup](https://ctacquisitions.com/?p=43413); [Deckary review](https://deckary.com/blog/macabacus-review)

### Inferences
- **Feature matrix, from the findings above** (✓ = documented, ? = unconfirmed, – = not found):
  | Feature | Macabacus | FactSet | CapIQ | UpSlide | Power-user | Excel Campus |
  |---|---|---|---|---|---|---|
  | Number/fill/font format cycles | ✓ | ? (hotkeys) | – | – | – | ✓ (static, not cycles) |
  | Alignment/indent cycles | ✓ | – | – | – | – | via styles |
  | AutoColor inputs/formulas/links | ✓ | ✓ | – | – | – | – |
  | Flip sign | ✓ | ? | – | – | – | – |
  | IFERROR wrap | ✓ | ? | – | – | – | – |
  | ROUND wrap | ? (not found) | ? | – | – | – | – |
  | Anchor ($) cycle on selection | ✓ | ? | – | – | – | – |
  | Fast fill right/down | ✓ | ? | – | – | – | – |
  | Paste exact / multi-format paintbrush | ✓ | ? | – | – | – | – |
  | Corporate table/chart styling | ✓ (per its marketing) | ✓ | – | ✓ | ✓ | – |
  | Excel→PPT live links | ✓ (per its marketing) | ✓ (PresLink lineage) | – | ✓ | ✓ | – |
- The banker features break into four groups: (1) cell formatting (cycles, AutoColor), (2) formula text rewriting (anchor, wraps, flip sign), (3) navigation and fill (Fast Fill, paste), and (4) Office linking (charts and slides). Groups 1-3 are pure Excel object-model operations on the selection. Group 4 needs PowerPoint automation as well.
- Every vendor builds its own undo, or would have to. Missing undo is a real UX cost for any add-in that changes cells.

### Gaps
- No full official Macabacus default shortcut list could be retrieved: the help centre is behind a login and the PDF returned 404. Shortcuts above are from search extracts and should be checked.
- I found no "ROUND wrap" feature documented for any tool, though a generic "wrap" is plausible.
- I found no product docs for "Analyst Pro" or a specific "Excel ribbon for bankers" product. TTS Turbo Macros came up in a listing (colour schemes, precedent tracer) but I did not get a primary source.
- I could not get FactSet hotkey specifics, such as a named list of format tools.
- Chart formatting tools (e.g. a Macabacus chart-format cycle) were not confirmed in the sources retrieved.

## Q2. What can Python-based Excel add-in frameworks easily do or not do for these features?

### Takeaway
PyXLL (Windows, in-process COM) is the best fit for Macabacus-style keyboard tools. It offers native shortcut binding on macros, ribbon XML, context menus, Application events, Qt task panes and full object-model access. The trade-off is that COM edits clear Excel's undo stack. Office.js add-ins (including xlwings Server and xlwings Wasm/Lite) can bind shortcuts, read formulas, precedents and dependents, and since September 2025 support undo. However, each call is an async round trip, shortcuts are restricted, and the Python layer adds latency. Python in Excel (=PY) cannot do any of this: it is a cloud calculation cell, not an automation layer.

### Cited Findings

**PyXLL (in-process, COM-based)**
- Keyboard shortcuts: `@xl_macro(shortcut="Ctrl+Shift+R")` or a `[SHORTCUTS]` config section, which can override shortcuts set in code. Format is modifier(s) plus key. A combination Excel already uses may not be assignable. `@xl_menu` takes the same argument — [PyXLL shortcuts](https://www.pyxll.com/docs/userguide/macros/shortcuts.html); [PyXLL macros API](https://www.pyxll.com/docs/api/macros.html); [PyXLL menus API](https://www.pyxll.com/docs/api/menus.html)
- Ribbon: CustomUI XML referenced in pyxll.cfg, with callbacks pointing to Python functions (`module.func`). Right-click context menus also use CustomUI — [PyXLL llms.txt index](https://www.pyxll.com/llms.txt)
- Events: `xl_event` for Application events (workbook open/close, sheet change, etc.), plus `xl_on_open` and `xl_on_reload`. Selection-change support was not confirmed in the index — [PyXLL llms.txt](https://www.pyxll.com/llms.txt)
- Custom task panes using Qt (PySide/PyQt), wxPython or tkinter. ActiveX controls can be embedded in sheets — [PyXLL llms.txt](https://www.pyxll.com/llms.txt)
- Full Excel object model via `xl_app()` with win32com or comtypes, explicitly positioned as a VBA replacement. `xl_disable` turns off screen updating and recalculation during macros. Also offers async UDFs, RTD and an LRU cache — [PyXLL llms.txt](https://www.pyxll.com/llms.txt)
- Community package pyxll-utils provides an `@xl_shortcut` decorator that binds via win32com at startup — [pyxll-utils](https://pyxll-utils.readthedocs.io/en/latest/shortcuts.html)

**Office.js (the platform under xlwings Server, xlwings Wasm/Lite, or any JS front end with a Python backend)**
- Custom keyboard shortcuts are supported in Excel on the web, Windows (Version 2102+) and Mac (16.55+). They require a shared runtime (SharedRuntime 1.1). User remapping requires KeyboardShortcuts 1.1 (`Office.actions.replaceShortcuts`, `getShortcuts`, `areShortcutsInUse`). Only "executeFunction" actions are supported — [Microsoft Learn: keyboard shortcuts](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- Shortcut rules: at least one modifier plus exactly one other key, and Shift cannot be the only modifier. Allowed keys are A-Z, 0-9, "-", "_" and "+", so keys such as ' (used by Macabacus Flip Sign) are not allowed. Chords must be simultaneous; sequential KeyTips are a separate feature. Microsoft recommends Ctrl+Shift+Alt+x to avoid conflicts. If there is a conflict, the user picks the winner in a dialog — [Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- On the web, shortcuts do not fire while the task pane has focus. Browser-reserved combinations such as Ctrl+Shift+N, Ctrl+Shift+T, Ctrl+Shift+W, Ctrl+Shift+S and Ctrl+Shift+X cannot be overridden — [Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- `Excel.Range` exposes `formulas`, `numberFormat` and `format` (font, fill, borders, alignment) from ExcelApi 1.1. It also has `getPrecedents`, `getDirectPrecedents`, `getDependents` and `getDirectDependents`, which return WorkbookRangeAreas, can span sheets, and throw ItemNotFound when there are none — [Microsoft Learn: Excel.Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range)
- Undo: since September 2025, most changes made by web add-ins can be undone with Ctrl+Z. Each `context.sync()` is one undo step by default, and several syncs can be grouped (default timeout 20 s, max 120 s). Previously, Office.js calls cleared the undo history — [Microsoft 365 Dev Blog, Sep 2025](https://devblogs.microsoft.com/microsoft365dev/excel-announces-undo-support-for-3rd-party-add-ins/)
- An older (2022) Q&A says custom functions clear the undo stack, which predates the change above — [Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1146975/undo-button-disabled-after-custom-functions-in-exc)

**xlwings (Server / Wasm / Lite)**
- xlwings Wasm limitations: no `App.alert()`, no object handles, no streaming functions. Ribbon and sheet buttons do not handle errors, so the docs recommend task-pane buttons. Only Office.js add-ins are supported. No SSO in Wasm. Pyodide constraints apply: fixed Python and package versions, no multiprocessing or threading, no TCP (SQLite only), CORS issues, a 2 GB memory cap, and no debugger — [xlwings Wasm limitations](https://server.xlwings.org/en/latest/wasm_limitations/)
- xlwings Lite runs on Pyodide and works with Microsoft 365 and Office 2021+. Its code lives in the workbook and is visible to users. To build a distributable add-in you need xlwings Server or xlwings Wasm, not Lite. There are no execution quotas — [xlwings Lite FAQ](https://lite.xlwings.org/faq); [xlwings Lite limitations](https://lite.xlwings.org/limitations)
- xlwings Server Office.js add-in docs (older mirror) — [xlwings officejs_addins (GitHub)](https://github.com/xlwings/xlwings/blob/master/docs/pro/server/officejs_addins.rst?plain=true)

**Python in Excel (=PY)**
- Runs in the Microsoft cloud in an isolated container, with no access to local files, network or Python installs. It reads data through `xl()` (ranges, names, tables, Power Query) and returns output only to its own cell. It cannot access or modify formulas, charts, PivotTables, macros or VBA. Sources are mostly from 2023-24 — [Microsoft Q&A: limits of Python in Excel](https://learn.microsoft.com/en-us/answers/questions/4457428/limits-of-python-implementation-in-excel-365); [xlwings blog: thoughts on Python in Excel](https://www.xlwings.org/blog/my-thoughts-on-python-in-excel); [Microsoft PY function](https://support.microsoft.com/pt-br/Excel/functions/py-function)

### Inferences
- **Feasibility tags for the feature ideas** (E = easy, M = moderate, H = hard or impractical):
  | Feature | PyXLL (COM) | Office.js + Python (xlwings Server/Wasm) | Python in Excel |
  |---|---|---|---|
  | Number/fill/font/border format cycles on a hotkey | E. Cycle state is held in Python or read from the current format | M. Easy API, but shortcut key limits apply and each press is a round trip | H, not possible |
  | Alignment/indent cycles | E | M (`format.horizontalAlignment`, `indentLevel`) | H |
  | AutoColor (classify input/formula/cross-sheet/external link by parsing `.Formula`) | E (read formulas in bulk, regex-classify, write colours) | M (`range.formulas` in bulk; large sheets need chunked syncs) | H |
  | Flip sign / IFERROR wrap / ROUND wrap / parentheses wrap (rewrite formula text) | E | E-M | H |
  | Anchor ($) cycle across a selection (needs a formula tokenizer, e.g. the `formulas` or `openpyxl` tokenizer) | M (tokenizer is the hard part) | M | H |
  | Fast Fill Right/Down (detect extent from adjacent row, skip spacers) | E-M | M | H |
  | Paste Exact (copy formulas without shifting refs) | M (read `.Formula`, write to target; clipboard interaction is awkward) | M (no general clipboard access; keep the copied range in add-in state instead) | H |
  | Multi-slot format paintbrush | E | M | H |
  | Precedent/dependent tracing UI | M (COM `.DirectPrecedents` is same-sheet only; cross-sheet needs parsing) | E-M (`getDirectPrecedents` spans sheets) | H |
  | Undo of add-in actions | H. COM changes clear the undo stack, so you need your own undo log (as Macabacus does) | E (native undo since Sep 2025, one step per sync) | n/a |
  | Ribbon buttons / context menus | E (CustomUI XML) | E (manifest ribbon commands; xlwings Wasm ribbon buttons lack error handling) | H |
  | Task pane UI (settings, palette editor) | E (Qt/wx/tk CTP) | E (HTML task pane) | H |
  | Selection/sheet-change events (e.g. live auto-colour) | E-M (`xl_event`, COM events) | M (Office.js event handlers; must run in the shared runtime) | H |
  | Excel→PowerPoint chart/slide links | M (COM to PowerPoint, Windows only) | H (separate PowerPoint add-in, no direct cross-app link) | H |
  | Mac / Excel on the web support | H (Windows only) | E | E (cloud) |
- Latency: hotkey tools must feel instant, so in-process COM (PyXLL) beats a remote Python server. xlwings Wasm/Lite avoid the network hop but carry Pyodide start-up cost. This is reasoning from the architecture, not a measured benchmark.
- Office.js shortcut rules (A-Z, 0-9, -, _, + only; Ctrl+Shift+Alt advised) mean Macabacus-style defaults such as Ctrl+Alt+' cannot be copied exactly. A remap UI via `replaceShortcuts` is the natural substitute.
- In practice, "Python in Excel" is irrelevant for formatting or shortcut features. It is only useful for analysis cells.

### Gaps
- PyXLL ribbon, shortcut and event doc pages were not fetched directly (the pages I tried returned 404 or I relied on the llms.txt index). Selection-change event support in PyXLL is unconfirmed.
- I did not find the exact Office.js API names, requirement set and Excel builds for undo grouping.
- I did not retrieve which xlwings Server/Wasm Range/Book methods map to formatting (e.g. `font.color`, `number_format`) or whether xlwings exposes Office.js keyboard shortcuts directly.
- No measured latency benchmarks for COM vs Office.js vs xlwings Server per keystroke.
- The requirement sets for `getPrecedents`/`getDependents` were not shown on the Range page excerpt.
- I did not check whether Microsoft's 2025-26 Python in Excel updates added any formatting or automation features.
