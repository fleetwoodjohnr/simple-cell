# Architecture

Simple Cell is an Electron app in the Simple family. It shares its shell with [Simple Docs](../simple-docs): title bar, File backstage, settings, recents, crash recovery, the GitHub-Releases updater, the secure IPC model, CI, and installers.

```
src/main/        Electron main process: window, secure IPC (validated arguments; reads/writes only paths the user granted),
                 atomic saves, settings/recents/recovery/history store, updater, LibreOffice bridge (XLS/XLSB only)
src/preload/     contextBridge API typed in src/shared/ipc-types.ts (no raw ipcRenderer in the page)
src/core/        The document core: pure TypeScript with no DOM and no UI. It runs in a Web Worker in the app and
                 in-process in unit tests.
  model/         Workbook, Sheet (sparse row-major cells), StyleTable (interned styles), names, tables; adjust.ts
                 rewrites references when rows/columns are inserted, deleted or moved
  engine/        Formula lexer/parser/printer (parser.ts), values and Excel coercion (values.ts), evaluator
                 (evaluate.ts), recalculation engine (engine.ts), functions/<category>.ts, catalog.json (every Excel
                 function, generated from Microsoft's documentation)
  numfmt/        Excel number format codes and date serials (1900 and 1904 systems)
  formats/       File adapters: bytes ⇄ Workbook plus compatibility warnings. zip.ts (fflate with zip-bomb limits),
                 xml.ts (xmldom, DOCTYPE stripped), opc.ts (relationships, content types); xlsx/ (read, write,
                 shift.ts keeps preserved content in step with structure edits), ods/ (read, write, OpenFormula
                 translation), csv, json, html, xml2003; registry.ts picks the adapter by format
  doc.ts         Transactions (every edit is undoable), typed input, structure edits, copy/fill
  host.ts        Everything the UI can ask of an open workbook; worker.ts exposes it over postMessage
src/renderer/app/Charts.tsx  ECharts adapter: chart option builder, on-grid chart overlay, editor; SVG for printing
src/grid/        Canvas grid: layout (axis.ts, layout.ts), painting (paint.ts), interaction (Grid.tsx, keys.ts)
src/renderer/    React UI: title bar, ribbon, formula bar, sheet tabs, status bar, dialogs, command registry
tests/           unit/ (vitest, against the core) · e2e/ (Playwright drives the real Electron app)
```

## Key decisions

**Electron rather than Tauri.** The brief preferred Tauri 2. Simple Cell uses Electron for four reasons:
- **Shared shell.** It is the stack of the existing Simple Docs app, whose shell, updater and 3-OS CI carry over unchanged.
- **End-to-end tests on every OS.** Playwright drives the real app on Windows, macOS and Linux. Tauri's WebDriver does not support macOS.
- **Identical rendering.** Every OS draws the canvas grid with the same Chromium, so text metrics and column widths match. On Linux, Tauri would use WebKitGTK.
- **Built-in PDF export.** `webContents.printToPDF` provides it.

The cost is a larger installer (about 100 MB, against Tauri's roughly 10 MB).

**Own calculation engine, in TypeScript, off the UI thread.** The engine is purpose-built rather than a dependency:
- **Why not the alternatives.** IronCalc (Rust/WASM) has its own data model and an XLSX writer that drops charts and pivots. HyperFormula is GPL.
- **What owning it gives us.** Direct control over spill ranges, LAMBDA/LET, structured references, implicit intersection, and keeping cached results for functions it can't calculate.
- **Where it runs.** The core lives in a Web Worker. The UI keeps only a cache of the cells on screen and receives patches after each edit.

**References are stored relative to their cell (R1C1 style).**
- A filled column of formulas shares one syntax tree.
- Copying a formula needs no rewriting.
- References that would fall off the sheet become `#REF!`.
- Inserting or deleting rows rewrites only the references that cross the change.

**Recalculation.** Changing a cell marks its dependents dirty, found through two indexes:
- a single-cell index, and
- a column-indexed range index; ranges wider than 64 columns go in a separate list.

The dirty formulas are then ordered by an iterative depth-first search, so long chains can't overflow the stack.

A formula that reads a cell which is still dirty (INDIRECT, OFFSET, spilled ranges) computes that cell first. Reading a cell that is already being computed is a circular reference: it is reported, and iterated if iterative calculation is on.

**Never a made-up result.** A formula that uses an Excel function Simple Cell doesn't implement keeps the value saved in the file, flagged as not recalculated. Without a saved value it shows `#NAME?`. Formulas referring to other workbooks also keep their saved value.

**Files keep what Simple Cell can't show yet.** Reading an XLSX keeps every package part, unknown worksheet and workbook elements, and the original style records, in a `carry` object on the workbook. Writing regenerates what Simple Cell models (cells, formulas, styles, layout, names, tables) and writes the carried content back in schema order. Package parts no longer reachable from the relationships are dropped. Inserting or deleting rows or columns rewrites ranges inside the carried XML (conditional formats, validations, drawing anchors, chart series) through the same transaction as the edit, so undo restores them. ODS works the same way with its own carry. Changes to preserved ODS content aren't adjusted yet.

**Every edit is a transaction.** Each change records how to revert it, which gives undo/redo and recovery. The same log is where a macro recorder and collaboration will connect.

## Security

- **Window isolation.** `contextIsolation` and `sandbox` are on and `nodeIntegration` is off.
- **Content-Security-Policy.** The packaged renderer has a strict policy: no inline scripts and no `eval`. Every E2E test fails on a policy violation.
- **No navigation.** Navigation and new windows are blocked.
- **File access.** The main process reads and writes only files the user granted through a dialog, recents, drag-and-drop or the command line.
- **CSV export.** By default, text starting with `= + - @` gets a leading apostrophe so other applications don't run it as a formula (CSV injection).

## Third-party licenses

| Component | License |
|---|---|
| Electron, React, Radix UI, zustand, Tailwind CSS, fflate, @xmldom/xmldom | MIT |
| lucide-react | ISC |
| Inter (UI font); Carlito, Liberation Sans/Serif/Mono (bundled metric-compatible fonts) | SIL OFL 1.1 |
| Caladea (bundled) | Apache-2.0 |
| electron-updater | MIT |
| Apache ECharts (charts) | Apache-2.0 |
| LibreOffice (optional, not bundled; used for XLS/XLSB) | MPL-2.0 |
