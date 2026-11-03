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
| Keyboard shortcuts | 🟡 | 64 command shortcuts registered + 24 built-in navigation/editing keys done; 11 planned ([SHORTCUT_REFERENCE.md](SHORTCUT_REFERENCE.md)) |
| File formats | 🟡 | XLSX, XLSM, XLTX, ODS, CSV, TSV, JSON, HTML, XML 2003 ✅; XLS ✅ through LibreOffice; XLSB open only, untested; content not yet shown (charts, CF, comments…) 🔒 kept ([FILE_FORMAT_SUPPORT.md](FILE_FORMAT_SUPPORT.md)) |
| Formatting | 🟡 | Cell formatting ✅; conditional formatting ✅ (all main rule types, shown, edited, saved); cell styles, Format Painter 📋 |
| Tables | 📋 | Structured references parse and calculate; creating and editing tables is planned |
| PivotTables | 📋 | |
| Charts | 📋 | |
| Macros | 📋 | VBA will be preserved and shown read-only; native JavaScript macros planned ([MACRO_SECURITY.md](MACRO_SECURITY.md) when available) |
| Automation | 📋 | |
| Data connectors | 📋 | |
| Printing and PDF | 📋 | |

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
| Structured table references (`Table1[Col]`, `[@Col]`, `#Headers`…) | 🟡 | Calculated, including tables read from XLSX; creating and editing tables is M4 |
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
| PDF | — | 📋 | M5 |
| Charts, pictures, shapes, comments, data validation, hyperlinks, PivotTables in files | 🔒 | 🔒 | Kept and saved unchanged; shown and editable in later milestones |
