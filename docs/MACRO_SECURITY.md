# Macro security

Simple Cell has two kinds of macro, and it treats them very differently.

| | JavaScript macros (Simple Cell's own) | VBA (from Excel .xlsm files) |
|---|---|---|
| Runs | Yes, in a sandbox, only when you start it | **Never** |
| Kept in files | .xlsx/.xlsm (a `customXml` part) and .ods (`SimpleCell/macros.xml`) | .xlsm, byte for byte; removed with a warning when saving as .xlsx |
| Can see | Only the open workbook, through the macro API | — |
| Shown | In the macro editor (Automate ▸ Macros, Alt+F8) | Read-only source viewer (Automate ▸ VBA, Alt+F11) |

## JavaScript macros

### When they run

- Macros never run when a workbook opens, and there is no auto-run event.
- A macro runs only when you choose Run in the macro editor, or start it by name.
- A workbook opened from a file asks first: the editor shows a warning and **Run is disabled** until you choose **Enable macros**.
  - Enabling stores a SHA-256 hash of that exact set of macros in your settings.
  - If anyone changes a macro, the hash no longer matches and Simple Cell asks again.
- Macros you write or edit yourself in the editor count as trusted.

Tested: e2e `macros › … trust before running a file’s macros`.

### The sandbox

Each run gets a new dedicated Web Worker, created from a `blob:` URL, which is closed when the run ends. The worker can't do any of the following:

| Can't… | Because |
|---|---|
| Reach the network | The app's Content-Security-Policy (`connect-src 'self' data: blob:`) applies to the worker. The prelude also removes `fetch`, `XMLHttpRequest`, `WebSocket`, `EventSource`, `importScripts`, `Request`/`Response` and `BroadcastChannel` as non-writable, non-configurable properties, so a macro can't put them back. |
| Read or write your files | A worker has no Node.js. The renderer has `nodeIntegration` off and `contextIsolation` and `sandbox` on, and the app's file bridge (`window.simpleCell`) doesn't exist inside a worker. |
| Touch the app's window | A worker has no DOM, `window` or `document`. |
| Store data or start other code | `indexedDB`, `caches`, `Worker`, `SharedWorker` and `navigator` are removed. |
| Run forever | Each run is stopped after 30 seconds, and the editor's **Stop** button ends it at once (`Worker.terminate()`). |

What a macro *can* do is send messages to the app asking for workbook operations. The app answers only this fixed list, and checks every argument:

| Operation | Checks |
|---|---|
| `sheets`, `getSelection`, `usedRange` | — |
| `getValues`, `getFormulas` | The sheet must exist; the address must be a valid range of at most 100,000 cells |
| `setValues` | Same, and the values must be a 2-D array matching the range (or a single value). Values are entered as if typed, as one undoable step |
| `format` | Only bold, italic, underline, strike, colour, fill, number format, size, font, alignment and wrap |
| `clear`, `select`, `addSheet`, `calculate` | The sheet and range must be valid |

Tested: e2e `macros`. One macro checks that `fetch`, `XMLHttpRequest`, `WebSocket`, `importScripts`, `document`, `window`, `require` and `process` are all undefined, and that putting `fetch` back fails.

### The API

```js
async function main(workbook) {
  const sheet = workbook.activeSheet          // or workbook.sheet('Data')
  const [[a]] = await sheet.range('A1').getValues()
  await sheet.range('B1').setValue(a * 2)
  await sheet.range('B2:C3').setValues([[1, 2], [3, 4]])
  await sheet.range('B1').format({ bold: true, fill: '#FFF2CC' })
  console.log(await workbook.sheets())
}
```

Every call returns a promise. `console.log` writes to the editor's output pane. An error stops the run and its message is shown; syntax errors also give their line number.

## VBA

- **Simple Cell never runs VBA.** It has no VBA interpreter.
- An .xlsm's `vbaProject.bin` is written back unchanged when you save as .xlsm, so Excel's macros keep working in Excel. Saving as .xlsx removes the VBA project, with a warning, as Excel does.
- Automate ▸ VBA shows each module's source read-only. It is read with Simple Cell's own compound-file reader and MS-OVBA decompression. Tested: unit `macros › VBA source is read…`, against a compound file built by the test, because no Excel-made sample is bundled.
- Opening a file with VBA shows a notice that it contains macros, which Simple Cell keeps but doesn't run.

## Not yet available

- **Macro recorder** (planned).
- **Python macros** (planned, via Pyodide in the same kind of sandbox).
- **Assigning macros to buttons and shortcuts.** Macros run from the editor, or by name.
- **Permission prompts for network access or files.** Macros have no access to either. When such permissions arrive, they will be granted per domain or per file and shown before each run.

## Reporting a problem

If you find a way for a macro to reach anything beyond its workbook, please report it privately to the maintainers rather than in a public issue.
