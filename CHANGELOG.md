# Changelog

## 0.1.0 (unreleased)

- First development version:
  - Canvas grid, cell editing, formula bar, Name Box, sheet tabs and status bar.
  - Calculation engine with 499 Excel functions, dynamic arrays and LAMBDA.
  - CSV and TSV files, AutoSave and crash recovery, Excel keyboard shortcuts.
  - Excel (.xlsx, .xlsm, .xltx, .xltm) and OpenDocument (.ods) files, read and written by Simple Cell itself, keeping the parts it can't show yet.
  - .xls through LibreOffice; .xlsb opened through LibreOffice.
  - JSON, HTML tables and XML Spreadsheet 2003.
  - Saving to a format that would lose content offers to save as .xlsx instead.
  - Conditional formatting: highlight rules, top/bottom, colour scales, data bars, icon sets and formula rules, shown on the grid and kept in .xlsx files.
  - Data validation: lists with an in-cell dropdown, number, date, text-length and formula rules, input messages, error alerts and Circle Invalid Data.
  - Tables (Ctrl+T) with styles, a total row and Table Design options; filters on tables and ranges (Ctrl+Shift+L).
  - Notes (Shift+F2) and hyperlinks (Ctrl+K), kept in .xlsx and .ods files.
  - Insert and Delete Cells with a shift, Paste Special, Format Painter, Go To Special, Cell Styles, Flash Fill, grouping and outline.
  - Printing and PDF: page setup, print area and titles, page breaks, headers and footers, fit to page, preview.
  - Charts: column, bar, line, area, pie, doughnut and scatter, created from data (Alt+F1), edited, moved and printed, and kept in .xlsx files; line, column and win/loss sparklines.
  - PivotTables (create, lay out, summarize, refresh; kept in .xlsx), Remove Duplicates, Text to Columns, Goal Seek, Descriptive Statistics.
  - JavaScript macros in a sandbox (no network, no files) with a trust prompt for files; read-only VBA viewer.
  - Faster recalculation of long dependency chains (100,000 cells: 35 s → under 0.5 s).
  - Accessibility fixes from an axe WCAG 2.2 AA audit: contrast of the default and cyan accents, button names that include their visible captions, the sheet tab list, and the macro editor in dark mode.
  - Fixed: keys typed right after pressing Enter in the Name Box could be lost when the screen was slow to redraw (the Name Box took focus back from the grid).
  - Fixed: closing an unsaved workbook crashed the window instead of asking to save (the dialog re-rendered forever).
  - Fixed: after Alt+Enter, F4 or a point-mode arrow while editing, letters typed quickly could land in the wrong order.
  - Accessibility: every dialog passes axe WCAG 2.2 AA in light and dark; the Insert Function list and the command palette work from the keyboard.
  - Fixed: keys pressed right after closing a dialog, or right after opening a validation list or filter menu, could be lost when the window was slow to redraw (focus moved a frame late).
  - Solver (Data ▸ Solver): maximize, minimize or reach a value subject to constraints, with the Simplex LP method (including integer and binary variables) and a nonlinear method; Answer Report; the model is saved where Excel's Solver keeps it.
  - Fixed: closing a workbook straight after typing into it could close it without asking to save.
  - AutoFit row height and AutoFit column width (Home ▸ Format) for the selection; double-clicking a row border now fits the row to its text (font size and wrapped lines) instead of resetting it.
  - Data Tables (Data ▸ What-If Analysis ▸ Data Table) with one or two input cells, and the calculation option "Automatic except data tables".
  - Scenario Manager: named sets of input values, shown with one click, compared in a Scenario Summary sheet.
  - macOS: ⌘W closes the workbook, not the whole window.
  - Fixed: an address typed quickly into Insert Link could be cleared again.
