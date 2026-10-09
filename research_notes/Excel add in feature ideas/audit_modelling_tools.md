# Excel Audit and Financial-Modelling Add-ins: Feature Landscape (idea fodder for a Python add-in)

Research date: 2026-10-09. About 17 tool calls. The Arixcel and Macabacus sites render with JavaScript, so most of their pages came back empty or as marketing summaries. Where an official page failed, I used third-party listings and say so.

## What specific features does each tool list on its official site/docs?

### Takeaway
Most tools share a core set of features: tracing precedents and dependents, a "formula map" that colours unique, copied and inconsistent formulas, workbook or sheet compare, error, hard-code and hidden-content checks, and a summary or report. Each tool then adds a few features of its own. Examples are OAK's formula Reconstruct, Spreadsheet Detective's AutoNames and sensitivity analysis, PerfectXL's FAST/ICAEW guideline checks and VBA review, Macabacus's 50+ check Model Check, and Arixcel's align-before-compare.

### Cited Findings

**Arixcel Explorer** (official site arixcel.com returned no content to the fetcher; these come from third-party listings)
- Formula pop-up (Ctrl+Q): shows a formula's logical structure, its current value and where its references are, and lets you jump to precedents and back. [G2 / listings via search](https://www.g2.com/products/arixcel-explorer/discuss); [SoftwareOne](https://platform.softwareone.com/product/arixcel-explorer/PCP-1927-2429)
- Multi-cell tracing: traces direct precedents or dependents of several cells at once. Ctrl+Shift+Q lists dependents and navigates to them. [SoftwareOne](https://platform.softwareone.com/product/arixcel-explorer/PCP-1927-2429)
- Formula map: applies a temporary colour scheme that shows formula inconsistencies, hard-coded data and external references. [SoftwareOne](https://platform.softwareone.com/product/arixcel-explorer/PCP-1927-2429)
- Calculation flow: shows how data flows in and out of a calculation block, and which cells are inputs, calculations or outputs. [SoftwareOne](https://platform.softwareone.com/product/arixcel-explorer/PCP-1927-2429)
- Compare: compares workbooks, worksheets or blocks of cells. It detects inserted rows and columns and auto-aligns the sheets before comparing. [SoftwareOne](https://platform.softwareone.com/product/arixcel-explorer/PCP-1927-2429)
- Runs on Windows desktop Excel; free trial available. [winget listing](https://winget.ragerworks.com/package/Arixcel.ArixcelExplorer)

**Macabacus** (auditing subset of a large banking productivity suite)
- Trace multiple precedents and dependents; "Trace In / Trace Out" tools go beyond Excel's native tracing. [Macabacus formula auditing](https://macabacus.com/excel/formula-auditing); [Macabacus Model Check](https://macabacus.com/excel/model-check)
- Step-by-step breakdown of nested formulas. [Macabacus](https://macabacus.com/excel/formula-auditing)
- Instant highlighting of formula errors. [Macabacus](https://macabacus.com/excel/formula-auditing)
- Formula Flow: pattern-based visualisation of formula structure that shows errors and inconsistencies across many cells, and updates live as cells change. [Macabacus](https://macabacus.com/excel/formula-auditing)
- Model Check: "50+ checks" for formula errors, hard-coded values and broken links. The individual checks are not listed publicly. [Macabacus Model Check](https://macabacus.com/excel/model-check)
- File Optimization: removes clutter and reduces file size. 100+ customisable shortcuts, formatting and charting tools, a pre-built table library (DCF, Black-Scholes), and Deck Check for PowerPoint proofing. [Macabacus features](https://macabacus.com/features); [Model Check page](https://macabacus.com/excel/model-check)

**PerfectXL suite** (Risk Finder, Explore, Highlighter, Compare, plus Formula Inspector and Source Inspector)
- Risk Finder checks for: broken formula sequences and incomplete references; circular references, with the exact loop path; hidden and "very hidden" content; Excel errors, with their affected cells and a risk level; external sources, listing the formulas, names and rules that reference them; unlocked cells and "client readiness"; hard-coded numbers in formulas (excluding common constants such as 10, 100 and 1000); and conformance to the FAST and ICAEW modelling standards. [PerfectXL Risk Finder](https://www.perfectxl.com/products/perfectxl-risk-finder/)
- Risk Finder workflow: analyses a file in about 10 seconds and produces a quality matrix. "Fix & Review" lets you jump between the tool and Excel and track fixes. Findings export to Excel, and a one-click PDF quality report is available. [PerfectXL Risk Finder](https://www.perfectxl.com/products/perfectxl-risk-finder/)
- Explore features:
  - Information-flow visualisation, exportable as an image.
  - Circular-reference path view.
  - Defined-names check for duplicate, hidden and unused names.
  - Model cleanup: what is used where, so unused content can be removed.
  - External-sources overview.
  - Separation view for inputs, calculations and outputs.
  - Complexity rating with simplification suggestions.
  - Unique-formula list with locations and associated risks.
  - VBA module review: which modules call which tabs.
  - Full inventory of workbook contents, with export.
  
  [PerfectXL Explore](https://www.perfectxl.com/products/perfectxl-explore/)
- Highlighter features:
  - Temporary cell colouring for consistency.
  - Error overlays.
  - Merged-cell detection.
  - Highlighting of breaks in formula patterns.
  - Unused-data detection.
  - Display of data types and number formats.
  - Up to nine highlight modes, plus custom highlights.
  - A list of highlighted cells.
  - Input, calculation and output view.
  - Dependency view, including unnecessary formulas in defined names.
  - Unlocked-cell marking.
  - Hidden-information overview.
  
  [PerfectXL Highlighter](https://www.perfectxl.com/products/perfectxl-highlighter/)
- Compare is described only as "Spreadsheet comparison: version management for Excel made easy". Formula Inspector and Source Inspector are named on the site without descriptions. [PerfectXL products](https://www.perfectxl.com/products/)

**Operis OAK (Operis Analysis Kit; OAK5 launched July 2020)**
- Map: a visual layout of the workbook that flags formula inconsistencies and departures from best practice. [operisanalysiskit.com](https://www.operisanalysiskit.com/)
- Compare: compares ranges, sheets or whole workbooks and builds an audit trail. Its stated differentiator is helping the user align two workbooks, so that moved rows don't appear as hundreds of false differences. [operisanalysiskit.com](https://www.operisanalysiskit.com/); [search summary of Operis materials](https://www.operis.com/operis-launch-oak5-the-latest-version-of-the-worlds-favourite-financial-modelling-software/)
- Summarize: a report on workbook size, layout, formula complexity and formula risk. [operisanalysiskit.com](https://www.operisanalysiskit.com/)
- Search: finds cells that don't follow modelling good practice. [Operis OAK5 launch](https://www.operis.com/operis-launch-oak5-the-latest-version-of-the-worlds-favourite-financial-modelling-software/)
- Formula Explorer / Formula Walker: walks a formula's precedents and branches with customisable keyboard shortcuts, and compares formulas or branches to find discrepancies. [operisanalysiskit.com](https://www.operisanalysiskit.com/)
- Reconstruct: rebuilds a formula in a safe environment to show its precedent and dependent chain down to the inputs. OAK5 Professional also prunes calculations that are not on a formula's active path, for example the untaken branch of an IF. [operisanalysiskit.com](https://www.operisanalysiskit.com/); [OAK5 launch](https://www.operis.com/operis-launch-oak5-the-latest-version-of-the-worlds-favourite-financial-modelling-software/)
- Desktop Edition: runs Summarize, Map and Compare in 64-bit memory outside Excel, even when Excel is 32-bit. Other features: named-range utilities and a "Review Issues" feature, which is undocumented on the page. [OAK5 launch](https://www.operis.com/operis-launch-oak5-the-latest-version-of-the-worlds-favourite-financial-modelling-software/); [operisanalysiskit.com](https://www.operisanalysiskit.com/)

**Excel Inquire add-in / Spreadsheet Compare (Microsoft)**
- Workbook Analysis report with categories Summary, Workbook, Formulas, Cells, Ranges and Warnings. [Microsoft Support](https://support.microsoft.com/en-US/Excel/analyze-a-workbook-with-spreadsheet-inquire)
- Workbook, Worksheet and Cell Relationship diagrams; Clean Excess Cell Formatting; a Workbook Passwords store for analysing and comparing protected files; Compare Files, which launches Spreadsheet Compare. [Microsoft Support](https://support.microsoft.com/en-US/Excel/analyze-a-workbook-with-spreadsheet-inquire); [Compare workbooks](https://support.microsoft.com/en-us/excel/compare-workbooks-using-spreadsheet-inquire)
- Available only in Windows Excel with Microsoft 365 Apps for enterprise or an equivalent edition, and admins can disable it. [Microsoft Support](https://support.microsoft.com/en-US/Excel/analyze-a-workbook-with-spreadsheet-inquire)

**Spreadsheet Detective** (available since 1997; Excel 2007 to 2019 and Office 365; commercial, with a free evaluation)
- Workbook report: per-sheet summary, potential errors, data flow between sheets, and external references. [spreadsheetdetective.com](https://spreadsheetdetective.com/main/)
- Distinct and copied formula shading on the sheet itself, with inconsistent formulas shaded bold. Condensed formula map: copy-pattern symbols, the original formula in comments, and colours for issues. Full graphical annotations. [spreadsheetdetective.com](https://spreadsheetdetective.com/main/)
- AutoNames: shows the active formula with descriptive names, derived from row and column labels, in place of A1 references, with hyperlinks to precedents. Available as a window, a task pane and a report covering all distinct formulas. [spreadsheetdetective.com](https://spreadsheetdetective.com/main/)
- Precedent and dependent outline reports; per-formula sensitivity analysis that flags problematic sensitivities. [spreadsheetdetective.com](https://spreadsheetdetective.com/main/)
- Audit process control: mark formulas as verified, questionable or erroneous, then produce an audit-status report. [spreadsheetdetective.com](https://spreadsheetdetective.com/main/)
- Worksheet comparison that survives inserted and deleted rows and columns, groups changed copies, pinpoints the changed part of each formula, and can compare sections within one sheet. [spreadsheetdetective.com](https://spreadsheetdetective.com/main/)
- Quick navigator that jumps to a sheet or name as you type. Named-range dialog to verify and edit names and update the formulas that use them. Workbook Discovery finds models and the other models that link to them. Scenario manager and Monte Carlo. Master-key encryption. Compliance Detective, a companion product, tracks metadata, changes and reviewers and scans repositories. [spreadsheetdetective.com](https://spreadsheetdetective.com/main/)

**ExcelAnalyzer (spreadsheetsoftware.com; Galeforce lineage not confirmed on the page)**
- Automatic error and risk checks; version comparison; in-Excel sign-off with comments; a PDF summary of every formula, error and risk; detection of hidden rows, columns and very-hidden sheets; an overall model risk score; "Maps" and "Cells Tested" views. Some vendor pages are flagged as outdated. [spreadsheetsoftware.com](https://spreadsheetsoftware.com/); [ExcelAnalyzer](https://spreadsheetsoftware.com/excelanalyzer/)

**Modano**
- Modular, "Lego-like" reusable model modules; automates cell locking; inserts sections consistently across all sheets; easy time-horizon changes; auditing for unique formulas and cell errors; output workbooks are macro-free; Windows Excel 2007 to 2019. [xltrail blog](https://www.xltrail.com/blog/financial-modeling-with-modano-and-xltrail); [Modano FAQ](https://app.modano.com/faqs)

**FormulaDesk**
- FormulaSpy: understand, troubleshoot and debug complex formulas, with structure and dependency insight. FormulaDesk Math: shows formulas in symbolic mathematical notation. Navigator: sheet-tab management with multi-line and vertical tabs. [formuladesk.com](https://www.formuladesk.com/); [FormulaSpy](https://www.formuladesk.com/formulaspy/)

**Other audit add-ins found incidentally**
- United Addins Spreadsheet Auditor: audit panel, formula decomposition, error tracking and a circular-reference tracker. [unitedaddins.com](https://www.unitedaddins.com/product/spreadsheet-auditor/)
- The Audit Toolbar, for accountants: formatting, analytics and sampling, templating and documenting, cleaning and structuring. [audittoolbar.com](https://www.audittoolbar.com/product-features)

### Inferences
- Draft categorised feature list, with the tools that offer each feature (all claims cited above):
  - **Precedent/dependent navigation:** Arixcel (multi-cell, Ctrl+Q pop-up), Macabacus (Trace In/Out), OAK (Formula Explorer), Spreadsheet Detective (hyperlinked AutoNames and outline reports), Inquire (cell relationship diagram), FormulaSpy.
  - **Formula map / unique-formula colouring:** Arixcel, Macabacus Formula Flow, PerfectXL Highlighter, OAK Map, Spreadsheet Detective shading and condensed map, Modano. PerfectXL Explore adds a unique-formula *list* with locations.
  - **Inconsistency and risk checks:** Macabacus Model Check (50+), PerfectXL Risk Finder (including FAST/ICAEW conformance), OAK Search and Summarize, ExcelAnalyzer (risk score), Inquire Warnings.
  - **Workbook compare:** Arixcel and OAK (align-before-compare), Spreadsheet Detective (insert-tolerant, changed formula parts), PerfectXL Compare, Inquire / Spreadsheet Compare, ExcelAnalyzer.
  - **Named ranges:** PerfectXL Explore (duplicate, hidden, unused), Spreadsheet Detective (verify, edit, update usages), OAK (utilities). Arixcel name management is unverified.
  - **External links:** PerfectXL (overview of formulas, names and rules using each source), Arixcel map, Spreadsheet Detective (inter-workbook flow, Workbook Discovery), Inquire Workbook Relationship, Macabacus (broken links).
  - **Circularity:** PerfectXL Risk Finder and Explore (exact loop path), United Addins circular tracker.
  - **Formula explanation:** Macabacus nested-formula step-through, OAK Reconstruct with active-path pruning, Spreadsheet Detective AutoNames, FormulaDesk Math notation, Arixcel structure pop-up.
  - **Documentation and reporting:** PerfectXL PDF quality report, ExcelAnalyzer PDF, OAK Summarize, Inquire Workbook Analysis, Spreadsheet Detective workbook report and audit-status report.
  - **Hygiene:** hidden and very-hidden content (PerfectXL, ExcelAnalyzer), unlocked cells (PerfectXL), merged cells (PerfectXL), excess formatting and file bloat (Inquire, Macabacus), unused data and names (PerfectXL), VBA inventory (PerfectXL Explore).
  - **Review workflow:** Spreadsheet Detective's verified, questionable and erroneous status; PerfectXL Fix & Review tracking; ExcelAnalyzer sign-off with comments.
- Ideas suited to Python: openpyxl, formulas or pycel-style parsing and graph libraries make AutoName-style formula explanation, Reconstruct-style active-path pruning, align-then-diff comparison, FAST rule checks and HTML or PDF reports feasible outside Excel. Operis's 64-bit Desktop Edition shows there is value in running analysis outside the Excel process.

### Gaps
- Arixcel's official feature pages and documentation could not be fetched because the site is JS-only. The Arixcel features above come from third-party listings, and named-range or external-link management features of Arixcel are unconfirmed.
- I found no current official pages for **Spreadsheet Professional**, **BoardRoom / Model Auditor**, the **FAST toolbar**, or **XLTools**. No feature claims are made for them here. They may be discontinued or renamed, and this needs a targeted search.
- I could not confirm the ExcelAnalyzer link to Galeforce. ExcelAnalyzer's current feature set is partly from pages flagged as outdated.
- Macabacus's 50 Model Check rules, and PerfectXL Compare, Formula Inspector and Source Inspector, have no public feature detail.
- Modano compatibility info (Excel 2007 to 2019) may be stale for 2026.

## Which features are distinctive/unique to one tool vs common across many?

### Takeaway
Tracing, formula-consistency colouring, compare and error or hard-code checks are standard in almost every tool. The features that differentiate tools are deeper analysis or workflow: formula reconstruction with branch pruning (OAK), label-based formula explanation (Spreadsheet Detective AutoNames), sensitivity analysis, standards-conformance checks (PerfectXL FAST/ICAEW), compare that aligns structure first, and reviewer sign-off and status tracking.

### Cited Findings
- Compare that tolerates structural changes is a stated differentiator for OAK (alignment assistance) and Arixcel (auto-detects inserted rows and columns), and Spreadsheet Detective also claims it. [OAK5 launch](https://www.operis.com/operis-launch-oak5-the-latest-version-of-the-worlds-favourite-financial-modelling-software/); [SoftwareOne Arixcel](https://platform.softwareone.com/product/arixcel-explorer/PCP-1927-2429); [Spreadsheet Detective](https://spreadsheetdetective.com/main/)
- Only OAK describes rebuilding a formula with pruning of inactive calculation paths (Reconstruct and Professional tools). [operisanalysiskit.com](https://www.operisanalysiskit.com/); [OAK5 launch](https://www.operis.com/operis-launch-oak5-the-latest-version-of-the-worlds-favourite-financial-modelling-software/)
- Only Spreadsheet Detective offers AutoNames (formulas rewritten with row and column labels), per-formula sensitivity analysis, Monte Carlo and master-key encryption. [Spreadsheet Detective](https://spreadsheetdetective.com/main/)
- Only PerfectXL checks against FAST and ICAEW guidelines, reviews VBA modules and their sheet calls, rates complexity, and offers nine highlight modes. [Risk Finder](https://www.perfectxl.com/products/perfectxl-risk-finder/); [Explore](https://www.perfectxl.com/products/perfectxl-explore/); [Highlighter](https://www.perfectxl.com/products/perfectxl-highlighter/)
- Only FormulaDesk Math renders formulas in symbolic maths notation. [FormulaDesk](https://www.formuladesk.com/)
- Only Modano offers modular build and reuse of model sections. [xltrail](https://www.xltrail.com/blog/financial-modeling-with-modano-and-xltrail)
- Only Macabacus Formula Flow is described as updating live as cells change; its Model Check bundles 50+ checks. [Macabacus](https://macabacus.com/excel/formula-auditing); [Model Check](https://macabacus.com/excel/model-check)

### Inferences
- An opportunity for a new tool is to combine the unique ideas in one place: AutoName explanations, active-path reconstruction, FAST rule checks, align-then-diff comparison and a reviewer status workflow. Tracing and formula maps are table stakes.

### Gaps
- There is no independent, side-by-side benchmark of these tools, such as compare accuracy or speed on large models.

## Any features users particularly praise?

### Takeaway
I found little user-review evidence. What exists praises Macabacus for its formatting and shortcuts, and criticises its price and unreliable undo. I found no Reddit or Wall Street Prep reviews of Arixcel, OAK or PerfectXL.

### Cited Findings
- Capterra reviewers praise Macabacus's formatting tools. One complaint is that some formatting applies only to the active sheet. [Capterra](https://capterra.com/p/10011368/Macabacus/reviews/)
- G2 users report that Macabacus undo is unreliable after certain shortcuts. The basic plan was listed from about $200 per user per year, a figure that may be dated. [G2](https://g2.com/products/macabacus-macabacus/reviews)
- An old (roughly 10-year-old) Wall Street Oasis thread recommends Macabacus, including its free version, for modelling, and calls the paid tools expensive for individuals. [WSO](https://www.wallstreetoasis.com/forum/investment-banking/best-excel-plug-in)
- Spreadsheet Detective cites a five-star ZDNet rating; this is a vendor claim. [Spreadsheet Detective (via search)](https://www.spreadsheetdetective.com/)

### Inferences
- Undo-safety is a recurring pain point. A Python add-in should avoid destroying Excel's undo stack, or use non-destructive overlays such as the "temporary colour schemes" in Arixcel and PerfectXL Highlighter.

### Gaps
- Searches did not surface r/financialmodeling threads, Wall Street Prep reviews, or user opinions on Arixcel, OAK, PerfectXL or the Inquire add-in. A direct Reddit search is still needed.
