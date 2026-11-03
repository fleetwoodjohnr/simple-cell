# Implementation status

_Last updated: 2026-11-02 · Milestones M0–M1 complete, M2 in progress · version 0.1.0 (unreleased) · tested on Linux only (Windows/macOS: CI workflow written, not yet run — there is no GitHub remote yet)._

"Done" means implemented **and** covered by a passing automated test. A button existing in the UI does not count.

## Test results (current)

| Suite | Result |
|---|---|
| `npm run typecheck` | ✅ clean |
| `npm test` (vitest) | ✅ 25 / 25 |
| `npm run e2e:headless` (Playwright + Electron, Linux/xvfb) | ✅ 6 / 6 (every test also fails on a CSP violation or an uncaught renderer error) |
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

- **Insert/delete cells.** Ctrl+Plus and Ctrl+Minus on a cell range insert or delete entire rows. Shifting cells right or down comes with M4.
- **Paste Special.** Values, formulas and formats are available from the context menu. The full Paste Special dialog comes with M4.
- **Printing and PDF.** Not available yet (M5).
- **File formats.** Only CSV/TSV can be saved. XLSX/ODS reading and writing is M3.

## M2 — Calculation engine 🟡 in progress

### Done

- **Parser.** Excel precedence (unary minus before `^`), range, intersection and union operators, cross-sheet and 3-D references, names, structured table references, `A1#` spill references, `@`, array constants, and `_xlfn.`/`_xlpm.` file prefixes. Test: unit `parser`.
- **Evaluation.** Excel coercion rules, the full error set, dynamic arrays with spilling and `#SPILL!`, implicit intersection for legacy formulas, lifting of scalar functions over arrays, and LET, LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY. Test: unit `engine`.
- **Recalculation.** Incremental, with dependency indexes and an iterative ordering (a 5,000-cell chain is tested). Also covered:
  - volatile functions, INDIRECT and OFFSET;
  - circular reference detection, with optional iteration;
  - manual calculation mode.

  Test: unit `engine`.
- **Number formats.** The Excel format-code engine (sections, conditions, colours, dates and times, elapsed time, fractions, scientific, thousands and scaling) and date serials with the 1900 leap-year bug. Test: unit `numfmt`.
- **Function catalog.** All 525 Excel functions from Microsoft's list (`scripts/gen-catalog.mjs`). Formulas using an Excel function that isn't implemented keep the value saved in the file, flagged; they are never shown as 0.
- **Implemented functions (231).** Math & trig 80/80, logical 19/19, text 48/50, date & time 25/25, information 21/22, lookup & reference 36/41.

### Still open in M2

- **Remaining functions.** Statistical, financial, engineering, database, compatibility and web functions.
- **Test coverage.** A table of test cases for each function.
- **Independent check.** A LibreOffice cross-check.
- **Generated docs.** `docs/FORMULA_COVERAGE.md` and `docs/SHORTCUT_REFERENCE.md`.
