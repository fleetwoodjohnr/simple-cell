# File format support

✅ means implemented **and** covered by an automated test (the test is named). Nothing here has been checked against Microsoft Excel itself. No copy of Excel was available to the project. Instead, files are checked three ways: Simple Cell reads what LibreOffice Calc 26.2 writes, LibreOffice reads what Simple Cell writes, and Simple Cell round-trips its own files.

| Format | Open | Save | How | Tests |
|---|---|---|---|---|
| Excel workbook (.xlsx) | ✅ | ✅ | Own reader and writer | unit `xlsx` (8), e2e `files › open an XLSX…` |
| Excel macro-enabled workbook (.xlsm) | ✅ | ✅ | Same; the VBA project is kept byte for byte | unit `xlsx › keeps parts…` |
| Excel template (.xltx, .xltm) | ✅ | ✅ | A template opens as a new, untitled workbook, as in Excel | e2e `files › legacy XLS…` |
| OpenDocument spreadsheet (.ods) | ✅ | ✅ | Own reader and writer | unit `ods` (5), e2e `files › open an XLSX…` (saves and reopens as .ods) |
| Excel 97–2003 (.xls) | ✅ | ✅ | Through LibreOffice, if installed: converted to .xlsx on open, and from .xlsx on save | e2e `files › legacy XLS…` |
| Excel binary workbook (.xlsb) | 🟡 | ❌ | Opened through LibreOffice, if installed. LibreOffice can't write .xlsb, so it is saved as another format. | Untested: an .xlsb fixture can't be made without Excel |
| CSV, TSV, TXT | ✅ | ✅ | Own | unit `interchange`, e2e `core › save as CSV…` |
| JSON | ✅ | ✅ | Own | unit `interchange › JSON…` |
| HTML table | ✅ | ✅ | Own | unit `interchange › HTML…` |
| XML Spreadsheet 2003 | ✅ | ✅ | Own | unit `interchange › XML Spreadsheet 2003…` |
| PDF | — | ✅ | Print dialog ▸ Save as PDF, through Electron's PDF engine | e2e `print` (pdftotext checks the text) |

## XLSX, XLSM, XLTX, XLTM

### Read and written

- **Cells.** Numbers, text, booleans, errors, and dates in both the 1900 and 1904 systems. Rich-text runs (mixed fonts in one cell) are shown as plain text and written back unchanged while the cell's text is unchanged.
- **Formulas.**
  - Formula kinds: ordinary, shared (fill-down blocks share one parsed formula), legacy array (Ctrl+Shift+Enter) and dynamic arrays. Dynamic arrays are written with the `metadata.xml` dynamic-array flag, so Excel spills them rather than adding `@`.
  - Prefixes: `_xlfn.`/`_xlpm.` for newer functions, plus `ANCHORARRAY` for `A1#` and `SINGLE` for `@`.
- **Saved results.**
  - On open, Simple Cell recalculates and compares its results with the ones saved in the file. Any difference is reported in the file's compatibility notes.
  - Formulas using a function Simple Cell doesn't calculate keep the saved result and are flagged. They are never replaced with 0.
- **Styles.** Fonts, fills, borders, alignment, wrap, rotation, indent, protection flags, number formats (built-in and custom) and theme colours with tint.
- **Layout.** Column widths, row heights, hidden rows and columns, merged cells, frozen panes, gridlines, zoom, and the active sheet.
- **Names and tables.** Defined names at workbook and sheet scope, and hidden and very-hidden sheets. Tables keep their name, columns, total row, style options and filter criteria. A table Simple Cell didn't change is written back exactly as it was read; a changed table is regenerated, and its extension elements are dropped.
- **Notes and hyperlinks.** Notes are read from Excel's legacy comments part and shown. A sheet whose notes change gets a new comments part and VML shapes; shapes that aren't notes, such as form controls, are kept. Threaded comments (Excel 365) are shown through their legacy note copy; changing a sheet's notes turns its threads into notes, with a warning. Hyperlinks are read and written, both external addresses (stored as relationships) and places in the workbook. Tested: unit `notes` (3 tests, one with LibreOffice reading the result).
- **AutoFilter.** Value lists, blanks, custom conditions and top 10 are read, applied and written, for sheets and tables. Other filter kinds are kept as they are but not re-applied: date grouping, colour, icon and dynamic filters. Excel's hidden `_FilterDatabase` name is regenerated. Tested: unit `xlsx › tables and filters…`, which includes LibreOffice reading the result.
- **Conditional formatting.** Every rule type of the main schema is read, shown and written back: cell value, formula, top/bottom, above/below average, duplicate/unique, text, dates, blanks/errors, colour scales, data bars and icon sets. Differential formats keep their original indexes, which tables and PivotTables also use. The Excel 2010 data-bar extensions stay linked to their rules. Tested: unit `xlsx › conditional formatting…` (3 tests, one with LibreOffice reading the result).
- **Data validation.** Every rule type (whole number, decimal, list, date, time, text length, custom) with its operator, formulas, input message and error alert. Tested: unit `xlsx › data validation…`, which includes LibreOffice reading the result.

### Kept unchanged (not yet shown or editable)

Simple Cell keeps every part and element of the file it doesn't understand, and writes them back in schema order. Examples:
- Charts, pictures and shapes; threaded comment parts (until the sheet's notes change).
- Conditional formats and data validation that exist only in the Excel 2010 extension list (for example icon sets with custom icons, or lists that refer to other sheets in old files), and hyperlinks.
- Sheet protection, sort state, the printer-settings part. (Page setup, margins, headers/footers and breaks are read into the model and written back.)
- PivotTables and their caches, external links.
- `customXml`, `extLst` extensions and the VBA project.

When the file is opened, its compatibility notes list what was kept but not shown.

When you insert or delete rows or columns, the ranges inside kept content are adjusted, and undo restores them. This covers:
- Other `sqref`/`ref` ranges. Conditional formats, data validation, notes and hyperlinks are part of the model, so they move like cells.
- Drawing anchors, comment positions and note anchors.
- Chart series formulas in every chart in the workbook.

Tested: unit `xlsx › keeps parts and elements…`.

### Warnings

| Situation | What happens |
|---|---|
| Saving a workbook with macros as .xlsx | Warned; the VBA project is removed (as Excel does). Save as .xlsm to keep it. |
| PivotTables, charts, pictures… | Kept and saved; the open notes say they aren't shown or editable yet |
| "Precision as displayed" | Kept; Simple Cell calculates with full precision (noted on open) |
| Links to other workbooks | Kept; formulas using them show the saved values |
| Formulas that can't be parsed | Kept as values, with a warning |

### Security

- Size limits: the entry count, the total inflated size (counted while inflating, so it isn't taken from the header) and the compression ratio are all limited. Zip bombs are refused with a clear message. Tested: unit `xlsx › refuses zip bombs…`.
- `DOCTYPE` declarations are stripped before XML is parsed, so external entities (XXE) are never resolved.
- VBA is never run. See [MACRO_SECURITY.md](MACRO_SECURITY.md) when available (M7).

## ODS

### Read and written

- **Cells.** Values of every ODF type (float, percentage, currency, date, time, boolean, string) and errors.
- **Formulas.** Translated between OpenFormula and Excel syntax:
  - References: `[.A1:.B2]` and `[$Sheet.A1]`.
  - Separators: `;` for arguments and `|` for array rows.
  - Operators: `~` (union) and `!` (intersection).
  - Function names: newer Excel functions are written with LibreOffice's `COM.MICROSOFT.` prefix, and the `LEGACY.*` names are mapped.
- **Styles and number formats.** Cell styles with their inheritance chain (font, colour, fill, borders, alignment, wrap, rotation). Data styles are converted to and from Excel number-format codes; a code that ODF can't express is warned about and saved as General.
- **Layout.** Repeated rows and columns (large empty repeats aren't expanded into memory), column widths, row heights, hidden rows and columns, merged cells, frozen panes, gridlines, zoom and the active sheet.
- **Workbook settings.** Named ranges and named expressions, the 1904 date system and iteration settings.

### Kept unchanged

The following are written back as they were:
- Shapes, pictures and embedded charts (the `Pictures/` and `Object N/` entries).
- Conditional formats, content validations and forms. Annotations (notes) and text links are read into the model and written back from it.
- `styles.xml` and `meta.xml`, apart from the date and generator.

Tested: unit `ods › keeps shapes, pictures…`.

Unlike XLSX, these kept elements are **not** adjusted when rows or columns are inserted or deleted.

## XLS and XLSB (through LibreOffice)

Simple Cell has no reader or writer of its own for the binary formats. If LibreOffice is installed, Simple Cell uses it in headless mode to convert:
- **On open:** XLS and XLSB are converted to XLSX and then read.
- **On save:** XLS is written as XLSX first, then converted.

Every save warns that the binary format drops newer features:
- Dynamic arrays.
- More than 65,536 rows or 256 columns.
- Newer functions.

Without LibreOffice, these files can't be opened, and the message says so.

## CSV and TSV

- **Encoding.** UTF-8 (with or without a byte order mark), UTF-16 LE and BE, and Windows-1252 are detected.
- **Parsing.** The delimiter (`,` `;` tab `|`) is sniffed. Quoted fields with embedded delimiters, quotes and line breaks are handled. Numbers, percentages, currency and dates are recognised as Excel does when typing.
- **Saving.** Only the current sheet is saved, with a warning when the workbook has more than one. Values are written as displayed.
- **Formula injection.** The CSV-injection guard is on by default: values starting with `=`, `+`, `-` or `@` that would run as formulas in other programs are written with a leading apostrophe. It can be turned off in Settings.

## JSON

- **Open.**
  - Accepted shapes: an array of objects (keys become the header row), an array of arrays, or an object holding one such array (`{"data": [...]}`).
  - ISO dates become date cells.
  - Nested values are written into cells as JSON text, with a note.
- **Save.**
  - If the first row holds unique text headers, the current sheet is saved as an array of objects; otherwise as an array of arrays.
  - Dates are written as ISO strings and formulas as their results.

## HTML

- **Open.** Every `<table>` in the page becomes a sheet. `colspan` and `rowspan` become merged cells, `<th>` and `<b>` become bold, and values are recognised as when typing. Scripts and styles in the page are ignored.
- **Save.** The visible sheets are written as formatted tables showing the displayed values:
  - Font, colours, fill, alignment and borders.
  - Merges as spans.
  - Hidden rows and columns left out.

## XML Spreadsheet 2003

- **Open and save.**
  - Cell types `Number`, `String`, `Boolean`, `DateTime` and `Error`.
  - Formulas in R1C1 notation, translated to and from Simple Cell's formulas.
  - Styles (font, fill, alignment, borders, number format), column widths, row heights, hidden rows and columns, merges, named ranges and frozen panes.
- **Not supported.** The format has no charts, conditional formatting or tables. Saving warns when the workbook came from a file that had such content.
