# Excel compatibility

This matrix tracks every area of Excel that Simple Cell aims to match. A feature is ✅ only when it is implemented **and** covered by an automated test. **Simple Cell does not claim full Excel compatibility.** It also has not been verified against Excel itself: no copy of Excel was available to the project. Formula results are checked against Microsoft's documented examples and against LibreOffice Calc.

| Status | Meaning |
|---|---|
| ✅ | Fully implemented and tested |
| 🟡 | Partially implemented (limits listed) |
| 🔒 | Preserved but not editable |
| 📋 | Planned |
| ❌ | Unsupported |

## Coverage summary

| Area | Status | Coverage |
|---|---|---|
| Formulas | 🟡 | 499 of 526 worksheet functions calculated; 481 with test cases ([FORMULA_COVERAGE.md](FORMULA_COVERAGE.md)) |
| Calculation semantics | 🟡 | See below |
| Keyboard shortcuts | 🟡 | 76 command shortcuts registered + 25 built-in navigation/editing keys (including ribbon KeyTips) done; 1 planned (Alt+F1 chart, M5) ([SHORTCUT_REFERENCE.md](SHORTCUT_REFERENCE.md)) |
| File formats | 🟡 | XLSX, XLSM, XLTX, ODS, CSV, TSV, JSON, HTML, XML 2003 ✅; XLS ✅ through LibreOffice; XLSB open only, untested; content not yet shown (charts, CF, comments…) 🔒 kept ([FILE_FORMAT_SUPPORT.md](FILE_FORMAT_SUPPORT.md)) |
| Formatting | 🟡 | Cell formatting, Format Painter ✅; conditional formatting ✅ (all main rule types, shown, edited, saved); cell styles 🟡 (built-in styles applied as formatting; named styles in files are not linked) |
| Editing | 🟡 | Insert/Delete Cells with a shift, Paste Special (all kinds, operations, skip blanks, transpose) ✅; Go To Special ✅; Flash Fill 🟡 (words, initials, delimiters, case, joins; not digits inside numbers); grouping/outline ✅ (Auto Outline and Data ▸ Subtotal 📋) |
| Tables | 🟡 | Create (Ctrl+T), styles, total row, filters, rename, convert to range ✅; calculated columns and auto-expansion 📋 |
| PivotTables | 📋 | |
| Charts | 🟡 | Column, bar, line, area, pie, doughnut, scatter (clustered, stacked, 100%) ✅ shown, created, edited and saved; other types kept 🔒; sparklines (line, column, win/loss) ✅ |
| Macros | 📋 | VBA will be preserved and shown read-only; native JavaScript macros planned ([MACRO_SECURITY.md](MACRO_SECURITY.md) when available) |
| Automation | 📋 | |
| Data connectors | 📋 | |
| Printing and PDF | 🟡 | Page setup, print area/titles, breaks, headers/footers, fit to page, preview, PDF, charts on pages ✅; pictures on printed pages 📋 |

## Calculation semantics

| Feature | Status | Notes |
|---|---|---|
| Operator precedence (`:` ` ` `,` then negation, `%`, `^`, `* /`, `+ -`, `&`, comparisons) | ✅ | `=-2^2` is 4, as in Excel |
| Relative, absolute and mixed references; cross-sheet and 3-D references | ✅ | |
| Whole-column and whole-row references | ✅ | |
| Named ranges and named formulas (workbook and sheet scope) | ✅ | |
| LET, LAMBDA, named LAMBDAs, LAMBDA helpers (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY) | ✅ | |
| Dynamic arrays, spilling, `#SPILL!`, the `A1#` spill reference | ✅ | |
| Implicit intersection (`@`, legacy formulas) | ✅ | Formulas from files without the dynamic-array flag calculate as legacy formulas |
| Legacy array formulas (Ctrl+Shift+Enter) | ✅ | |
| Structured table references (`Table1[Col]`, `[@Col]`, `#Headers`…) | ✅ | Typing over a header renames the column; formulas using the old name are not rewritten yet |
| Type coercion (text numbers and dates, logicals, empty cells) | ✅ | |
| Error values: `#NULL!` `#DIV/0!` `#VALUE!` `#REF!` `#NAME?` `#NUM!` `#N/A` `#SPILL!` `#CALC!` and the rest | ✅ | |
| Automatic and manual calculation, F9, Ctrl+Alt+F9 | ✅ | |
| Circular reference detection; iterative calculation | ✅ | |
| Volatile functions (NOW, TODAY, RAND, OFFSET, INDIRECT, CELL…) | ✅ | |
| 1900 date system with the 1900 leap-year bug; 1904 date system | ✅ | |
| Number formats (sections, conditions, colours, dates, elapsed time, fractions, scientific, scaling) | ✅ | |
| Locale-specific separators (`;` and `,` as the decimal point) | 📋 | Formulas are entered with `,` separators (en-US) |
| Precision as displayed | ❌ | Stored and round-tripped; calculations always use full precision |
| External workbook references (`[Book.xlsx]Sheet!A1`) | 🔒 | Kept; formulas using them show the saved result |
| Unsupported functions | 🔒 | Kept; the saved result is shown and flagged, never replaced by 0 |

## Conditional formatting

| Feature | Status | Notes |
|---|---|---|
| Cell value rules (greater/less than, between, equal…), formula rules | ✅ | Blank cells compare as 0, as in Excel |
| Top/bottom N and N%, above/below average (with standard deviations), duplicate/unique values | ✅ | Ties included in top/bottom |
| Text contains / begins / ends, dates occurring (yesterday … next month), blanks, errors | ✅ | Excel's own formulas are generated for new rules |
| 2- and 3-colour scales, data bars, icon sets (all 20 sets) | ✅ | Min, max, number, percent, percentile and formula thresholds |
| Rule priority, Stop If True, Manage Rules (reorder, delete), Clear Rules (selection or sheet) | ✅ | |
| Rules follow inserted and deleted rows and columns | ✅ | |
| Custom formats in the rule dialog (any font, fill and border) | 🟡 | Seven presets (Excel's six plus bold); custom formats are kept from files |
| Data bar options from Excel 2010 (negative bars, axis, solid fill) | 🔒 | Kept in the file; drawn as Excel 2007 gradient bars |
| Copying or filling cells carries their conditional formats | 📋 | |
| ODS conditional formats | 🔒 | Kept and saved, not shown |

## Data validation

| Feature | Status | Notes |
|---|---|---|
| Whole number, decimal, date, time, text length (all eight operators), list, custom formula | ✅ | Limits can be formulas (`=$D$1`) |
| Lists from typed items or a range; in-cell dropdown (arrow button or Alt+Down) | ✅ | Matching is case-insensitive, as in Excel |
| Ignore blank, input message, Stop / Warning / Information alerts with custom titles and messages | ✅ | Stop undoes the entry (Retry or Cancel); Warning asks Yes/No/Cancel |
| Circle Invalid Data, Clear Validation Circles | ✅ | Circles disappear as cells are corrected |
| Rules follow inserted and deleted rows and columns; saved in .xlsx | ✅ | |
| Validation of pasted or filled values | ❌ | As in Excel, pasting replaces rules rather than checking values; Simple Cell doesn't check pasted values either |
| ODS content validations | 🔒 | Kept and saved, not enforced |

## Tables and filters

| Feature | Status | Notes |
|---|---|---|
| Create table (Ctrl+T / Ctrl+L, Insert ▸ Table, Home ▸ Format as Table) with or without headers | ✅ | Without headers a whole sheet row is inserted for the header (Excel shifts only the table's columns) |
| Table styles Light 1–21, Medium 1–28, Dark 1–11 from the theme colours; banded rows/columns, first/last column | 🟡 | Generated approximations of Excel's built-in styles; custom table styles from files show as Medium 2 |
| Total row (Total label, SUBTOTAL 109 or 103) | ✅ | Per-column total functions are kept from files; choosing them in the UI is 📋 |
| Rename, Filter Button on/off, Convert to Range (keeps the look; structured references become ranges) | ✅ | |
| AutoFilter on tables and ranges (Ctrl+Shift+L): value checklist with search, blanks, conditions with wildcards, top N | ✅ | A row is hidden if any filter on the sheet hides it, as in Excel |
| Sort A→Z / Z→A from the filter menu; Clear, Reapply | ✅ | |
| SUBTOTAL/AGGREGATE recalculate when rows are filtered, hidden or shown | ✅ | |
| Filter by colour, by date groups (year/month), dynamic date filters | 🔒 | Kept from files and saved; not applied by Simple Cell |
| Calculated columns, auto-expanding tables, slicers | 📋 | |

## Notes and hyperlinks

| Feature | Status | Notes |
|---|---|---|
| Notes: new/edit (Shift+F2, context menu), delete, red corner marker, shown on hover | ✅ | Edited in a dialog rather than in place |
| Notes follow inserted/deleted rows and columns; Clear All removes them | ✅ | Copying or sorting cells doesn't carry notes yet |
| Threaded comments (Excel 365) | 🟡 | Shown as notes; replies can't be added |
| Hyperlinks: insert/edit (Ctrl+K), web, email and in-workbook addresses, ScreenTip, remove | ✅ | Excel's blue underlined style is applied to new links |
| Following links | ✅ | From the link chip under the selected cell. Only http, https and mailto addresses open outside the app; places in the workbook are selected |
| HYPERLINK() function | ✅ | Calculates its display value; clicking it is 📋 |

## Keyboard shortcuts

See [SHORTCUT_REFERENCE.md](SHORTCUT_REFERENCE.md) for the full list. Shortcuts follow Excel for Windows, and Excel for Mac on macOS. A Google Sheets preset changes the shortcuts that differ. Every shortcut can be rebound. Conflicting bindings are detected in the shortcut dialog and in a unit test of the registry.

## File formats

Details, limits and tests: [FILE_FORMAT_SUPPORT.md](FILE_FORMAT_SUPPORT.md).

| Format | Read | Write | Notes |
|---|---|---|---|
| XLSX, XLSM, XLTX, XLTM | ✅ | ✅ | Unknown parts kept and written back; ranges in kept content follow inserted/deleted rows and columns; VBA kept in .xlsm, never run |
| ODS | ✅ | ✅ | OpenFormula translation; data styles ⇄ number formats; shapes, pictures and charts kept |
| XLS | ✅ | ✅ | Through LibreOffice when installed |
| XLSB | 🟡 | ❌ | Opened through LibreOffice when installed; untested (no fixture can be made without Excel) |
| CSV, TSV, TXT | ✅ | ✅ | UTF-8, UTF-16 and Windows-1252; delimiter detection; CSV-injection guard on export |
| JSON, HTML, XML Spreadsheet 2003 | ✅ | ✅ | |
| PDF | — | ✅ | Save as PDF from the Print dialog (text stays selectable) |
| Pictures, shapes, PivotTables, chart types other than the seven drawn | 🔒 | 🔒 | Kept and saved unchanged; shown and editable in later milestones |
