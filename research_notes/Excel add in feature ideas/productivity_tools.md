# General-purpose Excel productivity add-ins: feature landscape (as of Oct 2026)

Scope: ASAP Utilities, Kutools for Excel, Ablebits Ultimate Suite, XLTools, Professor Excel Tools, Power-user, RDBMerge-style merge utilities, Peltier Tech Charts (briefly). Purpose: idea fodder for a Python-based Excel add-in.

Source codes used below: ASAP = ASAP Utilities; KU = Kutools; AB = Ablebits Ultimate Suite; XLT = XLTools; PE = Professor Excel Tools; PU = Power-user; RDB = RDBMerge; PT = Peltier Tech.

## What features does each tool list on its official site/docs, and how are they grouped?

### Takeaway
All of these tools group features by task type: select/navigate, range/cell, text, dedupe/compare, merge/split, sheets/workbooks, formulas, format/layout, import/export, and security. Kutools (300+ features, now with an AI group) and Ablebits (70+ tools) are the broadest. ASAP Utilities is organised as about 19 menu groups of small one-click utilities. Professor Excel focuses on modelling hygiene: formula wrapping, layout and fonts, break links and reduce file size. XLTools is a small set of four toolkits: popup calendar, data prep, version control and no-VBA automation. Power-user is PowerPoint-first, with Excel extras for charts, functions and an Excel-to-PowerPoint link.

### Cited Findings

**Kutools for Excel (ExtendOffice)**: 300+ features in 9 groups. The product page lists no version number, only a "What's New" link.
- Group "View": Navigation Pane (browse workbooks, sheets and other elements), Grid Focus (reading layout that highlights the active row and column), Hover View (preview cell contents on hover), Data Form — [Kutools product page](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Range & Cells": Copy Ranges (copy multiple non-adjacent selections), Paste to Visible (paste into filtered/visible cells only), Quickly Create Drop-down List, Merge Same Cells — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Editing": Charts (59 chart styles plus chart tools such as Gantt and Bullet), Super Find, Select Interval Rows & Columns, Batch Insert Smart Checkboxes, Delete Blank Rows, Extract Text, Alternate Row Shading, Convert Hyperlinks, Calculator — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Formulas": Kutools Functions (e.g. Count Visible Cells, Sum Visible Cells), Formula Helper (e.g. add months to a date, count days/hours/minutes between dates), Super Lookup (multi-condition lookup, simple VLOOKUP) — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Workbooks & Sheets": Split Workbook, Compare Worksheets, Combine Worksheets, batch Rename Worksheets, Print Multiple Selections Wizard — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Security": Password Manager, Protect Multiple Worksheets, Encrypt Cells, Encrypt Workbooks — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Mailings": Create Mailing List and Send Emails (bulk personalised email from a sheet) — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Filter & Statistic": Super Filter (multi-criteria), Advanced Sort, Filter by Font Color, Count by Color — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Group "Kutools AI": AI Aide (assistant), AI Data Analysis, AI Generate Formula, AI Explicate Formula, AI Translate Text, Cells Aide, AI Functions — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Review-site listings say its text operations cover 50+ options and its charts include Bubble and Speedometer types. They also say the AI assistant can generate formulas and write VBA. Listings state compatibility with Excel 2013+ and 365 — [Capterra listing](https://www.capterra.co.uk/software/1094448/Kutools-for-Excel)

**Ablebits Ultimate Suite**: "70+ professional tools, 300+ use cases". The installer filename indicates version 2026-1-3877-2566 — [Ablebits Ultimate Suite](https://www.ablebits.com/excel-suite/index.php)
- Merge: Merge 2 Sheets by key columns, Combine Sheets (by matching column headers), Merge Duplicates, Consolidate Sheets, Copy Sheets (many into one), Merge Cells without losing data, Merge Single Table with Several Lookup Tables, Merge Several Tables with a Single Lookup Table, Vlookup Wizard — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Compare: Compare 2 Sheets, Compare Multiple Sheets, Compare 2 Workbooks (same-named sheets) — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Dedupe: Remove Duplicates, Quick Dedupe, Compare 2 Tables, Find Duplicate Cells, Remove Dupes Within Cell (repeated substrings), Find Fuzzy Duplicates (near matches) — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Text: Trim Spaces (including line breaks), Change Case (upper/lower/proper/sentence), Add Text at a position, Extract Text, Remove Characters, Convert Text (characters, line breaks, numbers stored as text), Split Text, Split Names, Substring Tools, Regex Tools — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Manage: Workbook Manager (move/insert/rename sheets across workbooks), Insert Watermarks, Column Manager (move/insert/autofit/delete columns from a pane), Create Table of Contents — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Formulas: Formula Editor (edits long formulas in a tree view) — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Date & Time: Date & Time Wizard (add/subtract dates, age), Date Picker (pop-up calendar), Convert Text to Date — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Transform: Unpivot Table, Create Cards/labels, Split Table (by key column or row count), Split Workbook, Transpose by Key Columns, Transpose, Swap ranges, Flip range, Delete Blanks (rows/cols/sheets), Fill Blank Cells (from above/below) — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Search: Find and Replace (values, formulas, comments, links across workbooks), Select by Value/Color, Filter by selected values, Find Broken Links, Sync Selection across workbooks, Find Similar Cells — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Calculate: Sum/Count by Color, Count Characters/words, Cross-Sheet Operations, Calculate (range arithmetic) — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Utilities: Random Generator, Spell Number, Add/Remove Leading Zeros, Insert Pictures in bulk (linked), Copy Address, Convert Formulas to values, Notes Manager, Shuffle Cells, Select Randomly — [Ablebits](https://www.ablebits.com/excel-suite/index.php)

**ASAP Utilities**: about 19 menu groups. The tools page shows group names only, and the version is shown in the in-app About dialog. The site cites a "Voted Best Microsoft Excel Add-In 2025" award — [ASAP tools overview](https://www.asap-utilities.com/asap-utilities-excel-tools.php)
- Groups: My Tools & Shortcuts (custom favourites menu, user shortcuts), Display & Navigation, Select (by content, format or criteria), Sheets (bulk insert/rename/sort/split/export/hide/protect/print), Range (advanced find/replace, sort, paste, duplicates, named ranges), Fill (missing values, dates, add text/numbers), Columns & Rows (select/hide/delete/colour/join/reshape), Objects & Comments, Format & Layout (clear formatting, print layout, colour duplicates), Numbers & Dates (fix imported numbers, text dates, rounding, padding), Text (case, trim, remove accents/special characters), Formulas (to values, copy formulas without changing references, list formulas), Web Data (clean web content, decode URLs, strip HTML tags, build hyperlinks/HTML tables), Info & List (selection stats, workbook stats, #REF! detection), File & System (rename, back up, export, close files, clean external links/macros), Import (Excel/CSV/text/dBASE, merge multiple files), Export (Excel, PDF, CSV, JSON, delimited text, images, HTML), Start (open the workbook's folder with the file selected, Notepad, Calculator, clipboard viewer), Options (settings, shortcuts, "time and money savings" report) — [ASAP](https://www.asap-utilities.com/asap-utilities-excel-tools.php)
- Meta-tools: "Find and Run a Tool" (searches tool names and descriptions, like a command palette) and "Start Last Tool Again" (repeat last) — [ASAP](https://www.asap-utilities.com/asap-utilities-excel-tools.php)
- Featured tools also include: Fix Excel's incorrect last cell (Ctrl+End reset / sheet shrink), Extend selection to the lowest used row, Fix trailing minus signs, Fix reversed +/- values, Natural sort (e.g. version numbers), Select dates by weekday/month/quarter, Calendar popup date insert, Create sheets from a list, Sort sheet tabs by name/number/colour, Unprotect multiple sheets, List file names and properties from a folder, Remove hidden characters and line breaks, Spell numbers, and 38 extra worksheet functions (e.g. count by colour) — [ASAP popular tools page](https://www.asap-utilities.com/popular.php)

**Professor Excel Tools**: more than 125 features in 7 ribbon groups. It is Windows-only (Office 2010+), and a free trial is available — [Professor Excel Tools](https://professor-excel.com/professor-excel-tools/)
- Copy & Paste: Copy Again, Paste as Link and Transpose, Paste Exact Formula (no reference shifting), Paste Table into One Column (stack) — [PE feature list](https://professor-excel.com/professor-excel-tools-features/)
- Formula Editing: PROFEX functions (e.g. PROFEXSheetName, PROFEXMaxIf), wrap selected formulas in IFERROR/IFNA, bulk absolute/relative reference switch, Values, "Return blanks/zeros" wrapper, Calculation Operation (append +-*/ to existing formulas), wrap in ROUND/ROUNDUP/ROUNDDOWN, Clean Formulas (replace "=+" with "=", drop same-sheet prefixes), trace precedents/dependents for all selected cells, Follow first INDIRECT / INDIRECT Manager — [PE features](https://professor-excel.com/professor-excel-tools-features/)
- Insert Elements: Insert text/sub/superscript, about 250 country flags, symbols (checkmarks, Σ, Δ, circled numbers) — [PE features](https://professor-excel.com/professor-excel-tools-features/)
- Quick Cell Tools: select by type (formulas, constants, errors, blanks, visible, strikethrough, objects/charts), Center Across Selection, Swap Two Cells (optionally updating links), Trim, Force to Number/Text, Clear conditional formatting but keep current colours, Quick Number Format (saved favourite), Show as thousands/millions, Format as Heading / Value Cell (saved styles), Clear All (Ctrl+Del) — [PE features](https://professor-excel.com/professor-excel-tools-features/)
- Workbook Tools: Break Links (external links, validation, connections, pivots, CF, names, hyperlinks), Merge Files/Sheets, Name Manager including hidden names, Table of Contents / Table of Colors, Calculate Selection only, status bar file info (name, last saved, size), Headers & Footers, Compare Sheets, Directory listing, Workbook Info (formula/error counts), Remove PROFEX formulas — [PE features](https://professor-excel.com/professor-excel-tools-features/)
- Display & Layout: Layout Manager (gridlines, zoom, page break preview, formula view, zero display and grouping applied to all or selected sheets), Hidden rows/columns and grouping levels workbook-wide, Font Manager (one font across the workbook), Sheet Manager (including unhiding very hidden sheets), Sort Sheets (by name or tab colour) — [PE features](https://professor-excel.com/professor-excel-tools-features/)
- Finalize: Reduce File (sheet sizes, unused ranges, compress pictures), Export Manager (export sheets to PDF/XLSX, looping over a list of variable values), E-mail sheets as PDF/XLSX, Save selected sheets as PDF/XLSX — [PE features](https://professor-excel.com/professor-excel-tools-features/)

**XLTools**: Windows-only, four toolkits. It is sold as a lifetime licence with a 14-day trial — [XLTools](https://xltools.net/)
- Popup Calendar: date/time picker that can fill a row or column and apply a preferred format — [XLTools](https://xltools.net/)
- Data Prep: clean, transform, combine, query and export; crosstab to flat list; merge sheets; join tables with SQL queries — [XLTools](https://xltools.net/)
- Version Control: workbook backups, change log (who/what/when) for worksheets and VBA projects, version compare and rollback — [XLTools](https://xltools.net/)
- Automation: no-VBA task automation defined in a table, multi-step workflows, custom ribbon buttons — [XLTools](https://xltools.net/)

**Power-user (Power-user Softwares)**: PowerPoint-first, with Excel extras — [Power-user features](https://powerusersoftwares.com/features)
- Excel: UnPivot, advanced charts (Mekko, Sankey, etc.; 100 non-native chart types shared with PowerPoint), table styles, advanced functions, paste visible cells, delete empty cells, merged cells manager, clean formulas, format numbers, swap, flip, transpose, concatenate, add/remove characters — [Power-user features](https://powerusersoftwares.com/features)
- Functions named by the vendor: SumColor, CountColor, CountVisible, IsFormula, IsMerged, LastCell, LastRowNumber, LastColumnNumber, SumFromAllSheets, VlookupMax, SlicerItems, CAGR — [Power-user blog](https://www.powerusersoftwares.com/post/75-best-add-ins-plugins-and-apps-for-excel-free)
- Cross-app: Excel-PowerPoint link (one-click refresh of linked charts, tables and text), Content Library (icons, maps, templates), AI assistant (translate, rephrase, fix grammar), custom keyboard shortcuts for any command in PowerPoint, Excel or Word — [Power-user features](https://powerusersoftwares.com/features); [AppSource](https://appsource.microsoft.com/en-us/product/web-apps/power-user1580162777395.1)

**RDBMerge (Ron de Bruin, free)**
- Merges Excel, CSV and XML files from a folder into a summary workbook. You choose the files (including by filename pattern), sheets and ranges. Data can be stacked below or side by side, with an optional filename column. The output is a new workbook with a merged-data sheet and a log sheet. It is a UserForm on the Data tab. Ron's Windows pages are now hosted on jkp-ads.com, which describes them as "slightly outdated" — [jkp-ads RDBMerge](https://jkp-ads.com/rdb/win/addins/rdbmerge.htm)

**Peltier Tech Charts for Excel (brief)**
- Version 5.0 (released June 2025) has three ribbon tabs: custom charts, standard charts with better defaults (gap width, bar order, square scatter/pie), chart data management, chart formatting, save/export charts, and more tools. It is a VBA add-in for Windows and Mac — [Peltier software page](https://peltiertech.com/peltier-tech-software); [Indezine interview](https://blog.indezine.com/2025/07/peltier-tech-charts-for-excel-conversation-with-jon-peltier.html)
- v4 detail: 7 waterfall variants, 6 dot plots, box plots, histograms, Paretos, clustered-stacked column/bar charts, a series formula editor and bulk series edits, label the last point, colour data labels to match the series, corporate chart styles, and worksheet functions such as PT_Loess — [Peltier documentation](https://peltiertech.com/documentation); [Peltier UDFs](https://peltiertech.com/documentation/user-defined-functions)

**Macabacus (adjacent; finance modelling add-in, not in scope but often compared)**
- A reviewer praises customisable shortcuts and "cycles" (repeated keypresses cycle through formats), but notes that setting up and reordering cycles takes time. Complaints include an unreliable Excel-PowerPoint link and freezing when tracing precedents — [Software Advice Macabacus reviews](https://www.softwareadvice.com/cms/macabacus-profile/reviews/)

### Inferences
- The categories every tool shares make a good baseline list for any add-in: trim/clean text, change case, numbers-as-text to numbers, text-dates to dates, delete blank rows, fill blanks from above, formulas to values, merge cells without data loss, dedupe/colour duplicates, combine sheets/files, split by column, export sheets to separate files, TOC/sheet index, batch rename/sort sheets, count/sum by colour, date picker.
- Several useful features are "workbook hygiene" tools aimed at modellers, mostly from Professor Excel and ASAP: break links, reduce file size, fix last cell, hidden names, very hidden sheets, layout applied to all sheets, font standardisation. Python is well suited to these because they mostly inspect and rewrite workbook structure.
- Number-format cycling, as a keyboard shortcut that steps through a list of formats, is mostly a finance-add-in feature (Macabacus and similar). The general tools offer only "favourite number format" (Professor Excel) or generic "format numbers" (Power-user). This is a clear gap for a general add-in.
- A command-palette style launcher ("Find and Run a Tool") plus "repeat last tool" (ASAP) are low-cost but highly used features. Both appear in ASAP's top-25 usage list (see below).

### Gaps
- The Kutools page shows representative features only, not the full list of 300+. The exact current version and release date were not retrievable because the "What's New" URL returned 404.
- The XLTools detailed feature page returned 404. Only the toolkit-level descriptions above are confirmed.
- Peltier v5's specific new chart types were shown only as images on the site and could not be extracted.

## Which features are most popular or most requested?

### Takeaway
The best hard data is ASAP Utilities' own usage telemetry for 2024. Cleaning and text-fixing utilities dominate it: the top two are space trimming, followed by uppercase, numbers-stored-as-text and delete empty rows. Kutools' marketed "popular" set is checkboxes, duplicates, merge cells without data loss, drop-down extensions, text tools, combine sheets/workbooks and its AI aide. Reddit could not be accessed directly.

### Cited Findings
- ASAP Utilities' 25 most-used tools in 2024 (posted 9 Jan 2025), in order: 1 Delete leading, trailing and excessive spaces; 2 Delete leading and trailing spaces; 3 Change to UPPERcase; 4 Convert unrecognized numbers (text) to numbers; 5 Delete all empty rows; 6 Insert before and/or after each cell; 7 Start Each Word With Uppercase; 8 lowercase; 9 Clean data and formatting; 10 Start first word with Uppercase; 11 First character uppercase, rest lowercase; 12 Advanced character remove or replace; 13 Convert/recognize dates; 14 Formulas to calculated values; 15 Merge column data (join cells); 16 Delete number of leading characters; 17 Find and run a utility; 18 Copy values to empty cells below filled cells (fill down blanks); 19 Count and/or color duplicates; 20 Export worksheets as separate files; 21 Flip sign of numbers; 22 Give each set of duplicates its own colour; 23 Delete number of ending characters; 24 Merge and/or import multiple files; 25 Start the last used utility again — [ASAP news: most-used tools 2024](https://asap-utilities.com/news-detail.php?m=375)
- Kutools' "Popular Features": Batch Insert Smart Checkboxes, Find Duplicates, Kutools AI Aide, Merge Cells Without Losing Data, Drop-down List Extension Toolset, Text Tools, Combine Data from Multiple Worksheets/Workbooks — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- An ASAP reviewer valued being able to assign shortcuts to favourite utilities. Another said finding the right tool takes time, which is a discoverability problem — [Capterra ASAP reviews](https://www.capterra.ca/reviews/172223/asap-utilities-for-excel)
- AlternativeTo suggests Professor Excel Tools and Essential Excel Add-In as free alternatives. One user claims DataXL offers "almost all [Kutools] features for FREE" (a single opinion) — [AlternativeTo Kutools](https://alternativeto.net/software/kutools-for-excel)
- Macabacus reviewers single out formatting tools, shortcuts and Excel-PowerPoint linking — [Software Advice](https://www.softwareadvice.com/cms/macabacus-profile/reviews/)

### Inferences
- Demand is concentrated on boring, high-frequency cleanup tasks: trim, case, text to number, text to date, blank rows, fill down, formulas to values. These should be fast, single-key or palette actions in a Python add-in.
- Discoverability is a known problem with large suites (ASAP review comment; ASAP itself pushes "Find and Run"), so a searchable command palette plus favourites and repeat-last should ship early.

### Gaps
- Reddit r/excel is blocked for this research agent's fetch and search tools, so no Reddit sentiment is cited. A manual check of r/excel threads (e.g. "Kutools worth it", "favourite add-in") is recommended.
- Kutools and Ablebits publish no usage statistics, so their "popular" labels are marketing, not telemetry.

## What features are unique to one tool?

### Takeaway
Each tool has a few distinctive capabilities. Below are features found in only one tool's official listing among those reviewed. This was not exhaustively verified across all 300+ Kutools features.

### Cited Findings
- XLTools: version control for workbooks and VBA (change log, compare, rollback), SQL joins on Excel tables, and a table-driven no-VBA automation builder with custom ribbon buttons — [XLTools](https://xltools.net/)
- Ablebits: Find Fuzzy Duplicates, Regex Tools, Remove Dupes Within Cell, Formula Editor tree view, Sync Selection across workbooks, Merge with multiple lookup tables, Create Cards/labels, Insert Watermarks, Column Manager pane — [Ablebits](https://www.ablebits.com/excel-suite/index.php)
- Kutools: Grid Focus/reading layout, Hover View, Encrypt Cells, Password Manager, mail merge and bulk email from a sheet, Smart Checkboxes, Super Filter, Print Multiple Selections Wizard, and a broad AI group (Explicate Formula, Generate Formula, Translate) — [Kutools](https://www.extendoffice.com/product/kutools-for-excel.html)
- Professor Excel: Paste Exact Formula, Paste Table into One Column, bulk wrap in IFERROR/ROUND/"return blanks", Calculation Operation appended to formulas, INDIRECT Manager, Layout Manager across sheets, Font Manager, Reduce File, Export Manager driven by variable list, Clear CF but keep colours, Insert country flags, status-bar file info — [PE features](https://professor-excel.com/professor-excel-tools-features/)
- ASAP: Web Data group (decode URLs, strip HTML, build HTML tables), dBASE import, JSON export, Fix incorrect last cell, Fix trailing minus, natural/version sort, "time and money saved" report, Start folder/Notepad/clipboard viewer launcher — [ASAP overview](https://www.asap-utilities.com/asap-utilities-excel-tools.php); [ASAP popular](https://www.asap-utilities.com/popular.php)
- Power-user: Excel-PowerPoint live link, Mekko/Sankey charts, functions like SlicerItems, SumFromAllSheets, VlookupMax and IsMerged, custom shortcuts for any Office command — [Power-user features](https://powerusersoftwares.com/features); [Power-user blog](https://www.powerusersoftwares.com/post/75-best-add-ins-plugins-and-apps-for-excel-free)
- RDBMerge: a merge log sheet listing which files and ranges were merged, and side-by-side vs stacked merge layout — [jkp-ads RDBMerge](https://jkp-ads.com/rdb/win/addins/rdbmerge.htm)
- Peltier: series formula editor and bulk series edits, "label last point", and LOESS smoothing UDF — [Peltier documentation](https://peltiertech.com/documentation); [Indezine](https://blog.indezine.com/2025/07/peltier-tech-charts-for-excel-conversation-with-jon-peltier.html)

### Inferences
- Python-favourable differentiators for a new add-in, because they are hard in VBA but natural in Python: fuzzy dedupe (rapidfuzz), regex tools, SQL over sheets (duckdb/pandas), version diff/compare of workbooks, folder merge with a log (pandas), JSON/Parquet export, LOESS and statistical chart helpers, and natural sort.
- Features combining the gaps above could set an add-in apart: number-format, font-colour and fill cycling on hotkeys (Macabacus-style, absent from general tools), a command palette, and a workbook hygiene audit (links, names, hidden sheets, file bloat).

### Gaps
- Uniqueness is based on official feature pages reviewed here. Kutools' full 300+ list may overlap with items listed as unique to others, e.g. Kutools likely also has a date picker, TOC and reading layout variants. This was not verified.
- Power-user's site did not confirm a number-format cycling feature for Excel.
