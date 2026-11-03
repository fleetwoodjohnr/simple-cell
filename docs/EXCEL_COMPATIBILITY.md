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
| File formats | 🟡 | CSV/TSV ✅; XLSX, XLSM, ODS 📋 next milestone ([FILE_FORMAT_SUPPORT.md](FILE_FORMAT_SUPPORT.md)) |
| Formatting | 🟡 | Cell formatting ✅; conditional formatting, cell styles, Format Painter 📋 |
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
| Implicit intersection (`@`, legacy formulas) | ✅ | Legacy formulas are recognised when files are read (next milestone) |
| Legacy array formulas (Ctrl+Shift+Enter) | ✅ | |
| Structured table references (`Table1[Col]`, `[@Col]`, `#Headers`…) | 🟡 | Calculated; tables themselves arrive with XLSX import and the table feature |
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

## Keyboard shortcuts

See [SHORTCUT_REFERENCE.md](SHORTCUT_REFERENCE.md) for the full list. Shortcuts follow Excel for Windows, and Excel for Mac on macOS. A Google Sheets preset changes the shortcuts that differ. Every shortcut can be rebound. Conflicting bindings are detected in the shortcut dialog and in a unit test of the registry.

## File formats

| Format | Read | Write | Notes |
|---|---|---|---|
| CSV | ✅ | ✅ | UTF-8, UTF-16 and Windows-1252; delimiter detection; numbers and dates recognised; CSV-injection guard on export |
| TSV / TXT | ✅ | ✅ | |
| XLSX, XLSM, XLTX | 📋 | 📋 | Next milestone (M3), with preservation of unknown parts |
| ODS | 📋 | 📋 | M3 |
| XLS, XLSB | 📋 | 📋 | M3, through LibreOffice when installed |
| JSON, HTML, XML Spreadsheet 2003 | 📋 | 📋 | M3 |
| PDF | — | 📋 | M5 |
