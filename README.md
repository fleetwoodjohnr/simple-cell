# Simple Cell

A fast, private spreadsheet for Windows, macOS and Linux, in the Simple family of apps. It works offline, with no account and no telemetry.

> **Status: early development (0.1.0, not released).** What works today, with the tests that prove it, is listed in [IMPLEMENTATION_STATUS.md](IMPLEMENTATION_STATUS.md). Excel compatibility is tracked feature by feature in [docs/EXCEL_COMPATIBILITY.md](docs/EXCEL_COMPATIBILITY.md). Simple Cell does **not** claim full Excel compatibility.

## What works today

- **Grid and editing.** A canvas grid that scrolls to Excel's limits (1,048,576 rows × 16,384 columns).
  - Typing numbers, dates, percentages and currency, with Excel's automatic formats.
  - Formula bar and Name Box, multiple sheets, undo and redo.
- **Formulas.** Excel operator precedence, dynamic arrays with spilling, LET and LAMBDA, names, cross-sheet references, and circular reference detection.
  - 499 of Excel's 526 worksheet functions are implemented and checked against Microsoft's documented examples and LibreOffice ([docs/FORMULA_COVERAGE.md](docs/FORMULA_COVERAGE.md)).
- **Conditional formatting.** Highlight rules, top/bottom rules, colour scales, data bars, icon sets and formula rules, kept in .xlsx files.
- **Tables and filters.** Ctrl+T tables with styles and a total row; filters by value, condition or top N.
- **Notes and links.** Cell notes and hyperlinks, kept in .xlsx and .ods files.
- **Data validation.** Dropdown lists, number, date, text-length and formula rules, input messages and error alerts.
- **PivotTables and data tools.** PivotTables with a field list, Remove Duplicates, Text to Columns, Goal Seek and Descriptive Statistics.
- **Charts.** Column, bar, line, area, pie, doughnut and scatter charts, created from a range (Alt+F1), edited, printed, and read from and saved to .xlsx.
- **Printing and PDF.** Page setup, print areas and titles, page breaks, headers and footers, fit to page, a preview, and Save as PDF.
- **Macros.** JavaScript macros in a sandbox that can only touch the workbook, with a trust prompt for macros from files. VBA in .xlsm files is kept and can be read, but is never run. See [docs/MACRO_SECURITY.md](docs/MACRO_SECURITY.md).
- **Shortcuts.** Excel keyboard shortcuts for navigation, editing and formatting, plus a Google Sheets preset ([docs/SHORTCUT_REFERENCE.md](docs/SHORTCUT_REFERENCE.md)).
- **Files.** Excel (.xlsx, .xlsm, templates) and OpenDocument (.ods), read and written by Simple Cell itself. Charts, comments, conditional formatting and other content it can't show yet are kept and saved unchanged.
  - .xls through LibreOffice when it is installed.
  - CSV, TSV, JSON, HTML tables and XML Spreadsheet 2003.
  - Details in [docs/FILE_FORMAT_SUPPORT.md](docs/FILE_FORMAT_SUPPORT.md).
- **Safety.** AutoSave and crash recovery; a warning, with the option to save as .xlsx instead, before saving to a format that would lose content.

## Build and run

```bash
npm install
npm run dev                 # development with hot reload
npm run build && npm start  # production build
npm test                    # unit tests (vitest)
npm run e2e:headless        # Playwright drives the real app (Linux: under xvfb)
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the project's conventions and [ARCHITECTURE.md](ARCHITECTURE.md) for how it fits together.

## License

MIT. Bundled fonts are under the SIL Open Font License and Apache-2.0 (see ARCHITECTURE.md).
