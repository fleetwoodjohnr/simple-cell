# Implementation status

_Last updated: 2026-11-03 · Milestones M0–M5 complete (with the limits listed) · version 0.1.0 (unreleased) · tested on Linux only (Windows/macOS: CI workflow written, not yet run — there is no GitHub remote yet)._

"Done" means implemented **and** covered by a passing automated test. A button existing in the UI does not count.

## Test results (current)

| Suite | Result |
|---|---|
| `npm run typecheck` | ✅ clean |
| `npm test` (vitest) | ✅ 575 / 575 (481 function tests covering 973 formulas, plus engine, parser, number-format, registry and file-format tests; the 4 LibreOffice re-read tests, and the LibreOffice part of the data validation test, run when LibreOffice is installed) |
| `npm run crosscheck` (every function test formula also calculated by LibreOffice Calc 26.2) | ✅ 878 same result; 92 differ where Simple Cell matches Excel's documented value; 0 where only LibreOffice matches ([report](docs/LIBREOFFICE_CROSSCHECK.md)) |
| `npm run e2e:headless` (Playwright + Electron, Linux/xvfb) | ✅ 18 / 18 (every test also fails on a CSP violation or an uncaught renderer error). Two intermittent failures were seen on 2026-11-03, once each in `smoke` and `files › legacy XLS…`. Neither reproduced in repeated runs, and the cause is not yet known. |
| Windows, macOS | ⏳ not run: needs the GitHub Actions workflow (`.github/workflows/ci.yml`) and a remote |

## M0 — Scaffold ✅

- **Shell from Simple Docs.** Electron main process, secure IPC with granted paths, atomic saves, settings, recents, recovery and version-history store, and the updater (install on request only). Renamed for Simple Cell.
- **Renderer setup.** Strict CSP in packaged builds, Tailwind theme tokens shared with the family (light/dark/accent), and the app icon.
- **Packaging.** `electron-builder.yml` (AppImage, deb, rpm, NSIS, dmg/zip) and the CI workflow adapted from Simple Docs.

## M1 — Core grid and editing ✅

### Done

| Feature | Test |
|---|---|
| Canvas grid: frozen panes, merged cells, fills, borders, text overflow, wrap, #### for numbers that don’t fit, HiDPI, zoom | e2e `smoke`, visual check |
| Selection: cells, ranges, rows, columns, multiple ranges (Ctrl+click), Shift+click, drag; Ctrl+Arrow, Ctrl+Shift+Arrow, Home, Ctrl+Home/End, PageUp/Down | e2e `core › insert and delete rows…` |
| Editing: type to replace, F2 to edit, Esc cancels, Enter/Tab move (and cycle inside a selection), Alt+Enter line break, Ctrl+Enter fills the selection, Ctrl+Shift+Enter array formula | e2e `core › keyboard…` |
| Formula point mode (arrow keys insert references), F4 cycles $ references, coloured reference highlights, function autocomplete and argument hints | e2e `core › keyboard…` |
| Typed values get Excel’s automatic formats ($1,234.50, 25%, 1/15/2024, 1:30 PM); leading apostrophe keeps text | unit `engine › keeps typed formats`, e2e `smoke` |
| Formula bar and Name Box (go to a reference or name; type a new name to define it) | e2e `smoke`, `core` |
| Undo/redo of every edit, including inserting and deleting rows | unit `engine › undo and redo`, `insert and delete rows…`; e2e `smoke` |
| Insert/delete rows and columns with reference adjustment (#REF! for deleted cells) | unit `engine`, e2e `core` |
| Sheets: add, rename (double-click), delete, move (drag or menu), tab colour, hide/unhide; cross-sheet references follow renames | unit `engine › cross-sheet…`, e2e `core` |
| Clipboard: copy/cut/paste inside the app (formulas move relative), paste TSV from other apps, copy out as TSV and HTML | e2e `core › copy and paste…` |
| Fill down/right/up/left (Ctrl+D, Ctrl+R); AutoFill with the fill handle (numbers, dates, weekdays, months, "Item 1") | e2e `core` (fill down); AutoFill: unit tests pending |
| Formatting: bold/italic/underline/strike, font, size, colours, borders, alignment, wrap, indent, merge, number formats, Format Cells dialog | e2e `core › keyboard…` (bold); others: manual check, tests pending |
| Column width/row height (drag, double-click autofit, dialog), hide/unhide, freeze panes, gridlines, headings, zoom | manual check; tests pending |
| Status bar: Average, Count, Sum, Min, Max of the selection; Ready/Enter/Edit mode; circular reference warning | e2e `core › insert and delete rows…` |
| CSV/TSV open and save (UTF-8/UTF-16/Windows-1252 detection, delimiter sniffing, quoted fields); CSV-injection guard on export | e2e `core › save as CSV…` |
| AutoSave and crash recovery (renderer crash → restore) | e2e `core › unsaved work survives a renderer crash` |
| Command palette (Alt+Q), keyboard shortcut list with rebinding, conflict detection, import/export, Excel and Google Sheets presets | manual check; tests pending |

### Known limitations


## M2 — Calculation engine ✅

### Done

| Feature | Test |
|---|---|
| **Parser.** Excel precedence (unary minus before `^`), range, intersection and union operators, cross-sheet and 3-D references, names, structured table references, `A1#` spill references, `@`, array constants, `_xlfn.`/`_xlpm.` file prefixes. | unit `parser` |
| **Evaluation.** Excel coercion rules and the full error set; dynamic arrays with spilling and `#SPILL!`; implicit intersection for legacy formulas; scalar functions lifted over arrays; LET, LAMBDA, named LAMBDAs, eta-reduced functions (`GROUPBY(…, SUM)`). | unit `engine`, `functions` |
| **Recalculation.** Incremental, with single-cell and column-indexed range dependency indexes and an iterative topological ordering (a 5,000-cell chain is tested). Volatile functions, INDIRECT and OFFSET, circular references (with optional iteration), manual mode, F9 and Ctrl+Alt+F9. | unit `engine` |
| **Number formats.** The full Excel format-code language, and date serials with the 1900 leap-year bug and the 1904 system. | unit `numfmt` |
| **Function catalog.** All 526 Excel functions from Microsoft's list (`scripts/gen-catalog.mjs`). It drives autocomplete, argument hints, the Insert Function dialog and the coverage report. | unit `docs` |
| **499 functions calculated, 481 with test cases** (973 formulas, mostly Microsoft's documented examples). Full categories: math & trig, logical, text (except TRANSLATE and DETECTLANGUAGE), date & time, financial, engineering, database. See [FORMULA_COVERAGE.md](docs/FORMULA_COVERAGE.md). | unit `functions` |
| **Independent check.** Every test formula is also run through LibreOffice Calc. | `npm run crosscheck` |
| **Formula editing.** Point mode, F4, coloured reference highlights, function autocomplete, argument hints, Insert Function dialog. | e2e `core › keyboard…` |
| **Generated docs.** `docs/FORMULA_COVERAGE.md` and `docs/SHORTCUT_REFERENCE.md` are generated from the registries; a test fails when they are stale, and another checks the shortcut registry for conflicts. | unit `docs` |

### Known limitations

- **Not calculated (27 functions).** Formulas using them keep the saved result.
  - FORECAST.ETS (4 functions).
  - Online services: WEBSERVICE, IMPORTCSV, IMPORTTEXT, STOCKHISTORY, TRANSLATE, DETECTLANGUAGE, COPILOT, PY, FIELDVALUE.
  - FILTERXML, IMAGE, GETPIVOTDATA, PIVOTBY, RTD, CALL, REGISTER.ID.
  - The 7 CUBE functions.
- **Partial: GROUPBY.** Subtotal depth beyond 1 and the field-relationship argument are not implemented.
- **Partial: database functions.** Value criteria work; computed (formula) criteria do not.
- **Known deviation: DURATION and MDURATION.** They differ from Excel's documented examples by up to 2×10⁻⁶ (relative).
- **Not done yet.**
  - A Trace Precedents/Dependents UI (the engine API exists).
  - Error-indicator triangles.
  - The Evaluate Formula dialog.
  - Locale-specific formula separators.

## M3 — Files ✅

Details: [docs/FILE_FORMAT_SUPPORT.md](docs/FILE_FORMAT_SUPPORT.md).

### Done

| Feature | Test |
|---|---|
| XLSX/XLSM/XLTX/XLTM read and write: values, ordinary/shared/array/dynamic-array formulas, styles, number formats, theme colours, merges, sizes, hidden rows/columns, panes, names, tables, sheet visibility | unit `xlsx › reads a workbook written by LibreOffice…`, `round-trips…`, `dynamic arrays spill…`; e2e `files › open an XLSX, edit, save, reopen…` |
| LibreOffice opens the XLSX and ODS files Simple Cell writes, with the same results (including XLOOKUP and TEXTJOIN) | unit `xlsx › LibreOffice reads…`, `ods › LibreOffice reads…` |
| Preservation: drawings, charts, customXml, extLst, VBA and unknown elements are written back; ranges inside them follow inserted/deleted rows and columns, with undo | unit `xlsx › keeps parts and elements…` |
| Saved results compared after open; differences reported. Unimplemented functions keep the saved value | unit `xlsx › reports formulas whose saved results differ…`, `formulas using functions Simple Cell lacks…` |
| Zip-bomb and non-zip files refused; DOCTYPE stripped | unit `xlsx › refuses zip bombs…` |
| ODS read and write: OpenFormula translation, data styles ⇄ number formats, repeated rows/columns, merges, names, panes; shapes/pictures kept | unit `ods` (5 tests) |
| XLS open and save, XLSB open, through LibreOffice | e2e `files › legacy XLS opens through LibreOffice…` (XLS only) |
| Templates open as a new untitled workbook | e2e `files › legacy XLS…` |
| JSON, HTML table, XML Spreadsheet 2003 (R1C1 formulas) open and save | unit `interchange` |
| Lossy save warning lists what changes and offers "Save as .xlsx instead"; AutoSave writes only to .xlsx/.xlsm/.ods | e2e `files › open an XLSX…`, `core › save as CSV…` |
| Canvas grid loads the fonts a workbook uses before measuring text | visual check |

### Known limitations

- **Shown read-only or not at all.** Charts, pictures, hyperlinks, PivotTables and sheet protection from files are kept and saved, but not shown or enforced. They arrive in M4–M6.
- **ODS structure edits.** Kept ODS content (shapes, conditional formats) is not adjusted when rows or columns are inserted or deleted.
- **XLSB.** Untested, because no XLSB file can be produced without Excel. It can't be saved, since LibreOffice doesn't write it.
- **Not checked against Excel itself.** Files are checked against LibreOffice and by round trip.

## M4 — Editing depth ✅

### Done

| Feature | Test |
|---|---|
| Conditional formatting: every main rule type, evaluated in the core for the cells on screen (including empty cells); fills, fonts, borders and number formats from the rule; colour scales, data bars and 20 icon sets drawn on the grid | unit `condfmt` (9), e2e `condfmt` |
| Home ▸ Conditional menu with Excel's presets, rule dialog, formula rules, Manage Rules (reorder, delete), Clear Rules; undoable | e2e `condfmt` |
| Conditional formats in XLSX read and written (dxf indexes kept, Excel 2010 data-bar links kept); LibreOffice reads them | unit `xlsx › conditional formatting…` |
| Tables: Ctrl+T, 60 generated table styles, total row, Table Design tab (name, options, style gallery, convert to range); structured references; header edits rename columns | unit `tables` (6), e2e `tables` |
| AutoFilter for tables and ranges: header buttons, value checklist with search, blanks, conditions, top N, sort, clear, reapply; SUBTOTAL follows row visibility; XLSX round trip with `_FilterDatabase` | unit `tables`, `xlsx › tables and filters…`; e2e `tables` |
| Notes (Shift+F2, hover popup, red marker) and hyperlinks (Ctrl+K, link chip, open web/mail links, jump to places); follow structure edits; XLSX (comments + VML + hyperlinks) and ODS (annotations, text links) round trips | unit `notes` (3), e2e `notes` |
| Insert/Delete Cells with a shift (Ctrl+Plus/Minus dialog; Excel's reference rules) | unit `engine › insert and delete cells`, e2e `core › insert and delete rows…` |
| Paste Special (10 kinds, add/subtract/multiply/divide, skip blanks, transpose; Ctrl+Alt+V) and Format Painter (double-click keeps it on) | unit `engine › paste special`, e2e `editing` |
| Go To Special (blanks, constants/formulas by type, notes, conditional formats, validation, visible cells, last cell) and Find & Select shortcuts | unit `engine › go to special` |
| Cell Styles gallery: Excel's built-in styles applied as formatting | e2e `editing` |
| Grouping and outline: Group/Ungroup (Shift+Alt+→/←), outline gutter with bars, +/− and level buttons, nested groups keep their state; XLSX outline levels and collapsed flags | unit `outline`, e2e `outline` |
| Ribbon KeyTips (Alt, then letters; Esc steps back) on Windows and Linux | e2e `keytips` |
| Flash Fill (Ctrl+E): learns from examples (words by position, initials, text before/after a delimiter, case changes, joined columns) | unit `flashfill` (3), e2e `editing` |
| Data validation: all rule types, list dropdown (arrow and Alt+Down), input messages, Stop/Warning/Information alerts, Circle Invalid Data, Data ▸ Validation dialog; follows structure edits; XLSX round trip | unit `validation` (5), `xlsx › data validation…`; e2e `validation` |

### Known limitations

- Calculated table columns, table auto-expansion.
- Flash Fill for patterns inside numbers (phone numbers, digit groups) — the learner works on whole words and delimiters only.
- Split view; Subtotal (Data ▸ Subtotal) and Auto Outline.
- Command KeyTip letters matching Excel's exactly (they are assigned from command names).

## M5 — Charts, printing and PDF ✅

### Done

| Feature | Test |
|---|---|
| Page setup per sheet: orientation, paper (Letter, Legal, Tabloid, Executive, A3, A4, A5), margins, scale or fit to pages, gridlines, headings, centring, page order, manual breaks, header/footer codes (&P &N &A &F &D &T, left/centre/right) | unit `print` (3) |
| Print area and print titles (Excel's built-in names); Page Layout tab | unit `print` |
| Pagination like Excel's (down then over, titles repeated, hidden rows/columns skipped, fit to page) | unit `print › paginates…` |
| Print dialog with live page previews; Save as PDF (Electron's PDF engine, text stays text); Print through the system dialog | e2e `print` (checks the PDF with pdftotext and pdfinfo) |
| XLSX keeps page setup (printer-settings link, other header variants) and writes it back | unit `print › XLSX keeps page setup…` |
| Charts: column, bar, line, area, pie, doughnut, scatter; clustered/stacked/100%; titles, legend, data labels, axis titles; drawn with ECharts over the grid | unit `charts` (4), e2e `charts` |
| Insert ▸ Chart and Alt+F1 from a range (Excel's choice of series and categories); select, drag to move, resize, Delete, double-click to edit; undo | e2e `charts` |
| Series references follow inserted/deleted rows and columns; hidden cells aren't plotted; data updates live | unit `charts` |
| XLSX charts read from DrawingML and drawn; unchanged charts written back byte for byte; moved ones get a new anchor; new or edited ones a new chart part; pictures in the same drawing kept; LibreOffice reads the charts Simple Cell writes | unit `charts-xlsx` (3) |
| Charts on printed pages and in PDFs (SVG) | e2e `charts` |
| Sparklines: line, column, win/loss (markers, high/low/first/last/negative points), Insert ▸ Sparklines, clear; follow structure edits; XLSX extension list round trip (other extensions kept); LibreOffice reads them | unit `sparklines` (2), e2e `charts › sparklines…` |

### Known limitations

- Chart types beyond the seven (combo charts show their first type; radar, stock, surface, bubble and Excel 2016 chartex charts are kept, not shown); chart formatting beyond colours (fonts, axis scales, number formats) isn't read or edited.
- Printing: icons from conditional formatting and pictures on printed pages; a different paper size per sheet in one PDF (the first sheet's size is used).

## Next

- **M5 onward.** See the plan in ARCHITECTURE.md.
