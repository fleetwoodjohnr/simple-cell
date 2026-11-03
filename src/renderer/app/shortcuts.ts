// Keyboard shortcut registry, from Microsoft's "Keyboard shortcuts in Excel" (Windows and Mac pages).
// `command` entries are dispatched by the command registry; `builtin` entries are handled by the grid and editor
// (navigation, entering data) and are listed here so the shortcut reference and its tests cover them too.
// Key format: Mod = Ctrl (Windows/Linux) or ⌘ (macOS); Ctrl = the Control key on macOS; chords are joined with "-".

export type ShortcutContext = 'grid' | 'editing' | 'global'
export type ShortcutStatus = 'done' | 'planned'

export interface ShortcutDef {
  command: string
  win: string[]
  mac?: string[]
  /** Google Sheets preset (when it differs). */
  sheets?: string[]
  context: ShortcutContext
  status?: ShortcutStatus
}

export interface BuiltinShortcut {
  keys: string
  mac?: string
  action: string
  context: ShortcutContext
  status: ShortcutStatus
}

export const SHORTCUTS: ShortcutDef[] = [
  // Files and workbooks
  { command: 'file.new', win: ['Mod-n'], context: 'global' },
  { command: 'file.open', win: ['Mod-o', 'Mod-F12'], mac: ['Mod-o'], context: 'global' },
  { command: 'file.save', win: ['Mod-s', 'Shift-F12'], mac: ['Mod-s'], context: 'global' },
  { command: 'file.saveAs', win: ['F12', 'Mod-Shift-s'], mac: ['Mod-Shift-s'], context: 'global' },
  { command: 'file.close', win: ['Mod-w', 'Mod-F4'], mac: ['Mod-w'], context: 'global' },
  { command: 'file.print', win: ['Mod-p', 'Mod-F2'], mac: ['Mod-p'], context: 'global' },
  { command: 'tab.next', win: ['Mod-Tab', 'Mod-F6'], mac: ['Ctrl-Tab'], context: 'global' },
  { command: 'tab.prev', win: ['Mod-Shift-Tab', 'Mod-Shift-F6'], mac: ['Ctrl-Shift-Tab'], context: 'global' },
  // Editing
  { command: 'edit.undo', win: ['Mod-z', 'Alt-Backspace'], mac: ['Mod-z'], context: 'grid' },
  { command: 'edit.redo', win: ['Mod-y', 'F4'], mac: ['Mod-y', 'Mod-Shift-z'], sheets: ['Mod-y', 'Mod-Shift-z'], context: 'grid' },
  { command: 'edit.copy', win: ['Mod-c', 'Mod-Insert'], mac: ['Mod-c'], context: 'grid' },
  { command: 'edit.cut', win: ['Mod-x'], context: 'grid' },
  { command: 'edit.paste', win: ['Mod-v', 'Shift-Insert'], mac: ['Mod-v'], context: 'grid' },
  { command: 'edit.pasteValues', win: ['Mod-Shift-v'], mac: ['Mod-Shift-v'], context: 'grid' },
  { command: 'edit.fillDown', win: ['Mod-d'], context: 'grid' },
  { command: 'edit.fillRight', win: ['Mod-r'], context: 'grid' },
  { command: 'edit.find', win: ['Mod-f', 'Shift-F5'], mac: ['Mod-f'], context: 'grid' },
  { command: 'edit.replace', win: ['Mod-h'], mac: ['Ctrl-h'], sheets: ['Mod-h'], context: 'grid' },
  { command: 'edit.goto', win: ['Mod-g', 'F5'], mac: ['Ctrl-g', 'F5'], context: 'grid' },
  { command: 'edit.selectAll', win: ['Mod-a', 'Mod-Shift-Space'], mac: ['Mod-a'], context: 'grid' },
  { command: 'edit.insertDate', win: ['Mod-;'], mac: ['Ctrl-;'], context: 'grid' },
  { command: 'edit.insertTime', win: ['Mod-Shift-;'], mac: ['Mod-;'], context: 'grid' },
  { command: 'edit.editCell', win: ['F2'], mac: ['Ctrl-u', 'F2'], context: 'grid' },
  { command: 'edit.insertCells', win: ['Mod-Shift-=', 'Mod-Plus'], mac: ['Ctrl-Shift-=', 'Mod-Shift-='], sheets: ['Mod-Alt-='], context: 'grid' },
  { command: 'edit.deleteCells', win: ['Mod-Minus'], mac: ['Mod-Minus', 'Ctrl-Minus'], sheets: ['Mod-Alt-Minus'], context: 'grid' },
  { command: 'edit.clearFormats', win: [], mac: [], sheets: ['Mod-\\'], context: 'grid' },
  // Formatting
  { command: 'fmt.bold', win: ['Mod-b', 'Mod-2'], mac: ['Mod-b'], context: 'grid' },
  { command: 'fmt.italic', win: ['Mod-i', 'Mod-3'], mac: ['Mod-i'], context: 'grid' },
  { command: 'fmt.underline', win: ['Mod-u', 'Mod-4'], mac: ['Mod-u'], context: 'grid' },
  { command: 'fmt.strike', win: ['Mod-5'], mac: ['Mod-Shift-x'], sheets: ['Alt-Shift-5'], context: 'grid' },
  { command: 'fmt.cells', win: ['Mod-1'], context: 'grid' },
  { command: 'fmt.font', win: ['Mod-Shift-f', 'Mod-Shift-p'], mac: ['Mod-Shift-f'], context: 'grid' },
  { command: 'fmt.general', win: ['Mod-Shift-`'], mac: ['Ctrl-Shift-`'], context: 'grid' },
  { command: 'fmt.currency', win: ['Mod-Shift-4'], mac: ['Ctrl-Shift-4'], sheets: ['Mod-Shift-4'], context: 'grid' },
  { command: 'fmt.percent', win: ['Mod-Shift-5'], mac: ['Ctrl-Shift-5'], sheets: ['Mod-Shift-5'], context: 'grid' },
  { command: 'fmt.scientific', win: ['Mod-Shift-6'], mac: ['Ctrl-Shift-6'], sheets: ['Mod-Shift-6'], context: 'grid' },
  { command: 'fmt.date', win: ['Mod-Shift-3'], mac: ['Ctrl-Shift-3'], sheets: ['Mod-Shift-3'], context: 'grid' },
  { command: 'fmt.time', win: ['Mod-Shift-2'], mac: ['Ctrl-Shift-2'], sheets: ['Mod-Shift-2'], context: 'grid' },
  { command: 'fmt.number', win: ['Mod-Shift-1'], mac: ['Ctrl-Shift-1'], sheets: ['Mod-Shift-1'], context: 'grid' },
  { command: 'fmt.outlineBorder', win: ['Mod-Shift-7'], mac: ['Mod-Alt-0'], sheets: ['Alt-Shift-7'], context: 'grid' },
  { command: 'fmt.noBorder', win: ['Mod-Shift-Minus'], mac: ['Mod-Alt-Minus'], sheets: ['Alt-Shift-6'], context: 'grid' },
  { command: 'fmt.alignLeft', win: [], mac: [], sheets: ['Mod-Shift-l'], context: 'grid' },
  { command: 'fmt.alignCenter', win: [], mac: [], sheets: ['Mod-Shift-e'], context: 'grid' },
  { command: 'fmt.alignRight', win: [], mac: [], sheets: ['Mod-Shift-r'], context: 'grid' },
  { command: 'fmt.hideRows', win: ['Mod-9'], mac: ['Ctrl-9'], sheets: ['Mod-Alt-9'], context: 'grid' },
  { command: 'fmt.unhideRows', win: ['Mod-Shift-9'], mac: ['Ctrl-Shift-9'], context: 'grid' },
  { command: 'fmt.hideCols', win: ['Mod-0'], mac: ['Ctrl-0'], sheets: ['Mod-Alt-0'], context: 'grid' },
  { command: 'fmt.unhideCols', win: ['Mod-Shift-0'], mac: ['Ctrl-Shift-0'], context: 'grid' },
  // Formulas
  { command: 'formulas.autosum', win: ['Alt-='], mac: ['Mod-Shift-t'], sheets: ['Alt-Shift-='], context: 'grid' },
  { command: 'formulas.insertFunction', win: ['Shift-F3'], mac: ['Shift-F3'], context: 'global' },
  { command: 'formulas.names', win: ['Mod-F3'], mac: ['Mod-F3'], context: 'grid' },
  { command: 'formulas.calculate', win: ['F9'], mac: ['F9', 'Mod-='], context: 'grid' },
  { command: 'formulas.calculateAll', win: ['Mod-Alt-F9'], mac: ['Mod-Alt-F9'], context: 'grid' },
  { command: 'formulas.showFormulas', win: ['Mod-`'], mac: ['Ctrl-`'], sheets: ['Mod-`'], context: 'grid' },
  // Sheets and view
  { command: 'sheet.insert', win: ['Shift-F11', 'Alt-Shift-F1'], mac: ['Shift-F11'], sheets: ['Shift-F11'], context: 'grid' },
  { command: 'sheet.next', win: ['Mod-PageDown'], mac: ['Ctrl-PageDown', 'Alt-ArrowRight'], sheets: ['Mod-Shift-PageDown', 'Alt-ArrowDown'], context: 'grid' },
  { command: 'sheet.prev', win: ['Mod-PageUp'], mac: ['Ctrl-PageUp', 'Alt-ArrowLeft'], sheets: ['Mod-Shift-PageUp', 'Alt-ArrowUp'], context: 'grid' },
  { command: 'view.ribbon', win: ['Mod-F1'], mac: ['Mod-Alt-r'], context: 'global' },
  { command: 'view.zoomIn', win: ['Mod-Alt-='], mac: [], sheets: [], context: 'grid' },
  { command: 'view.zoomOut', win: ['Mod-Alt-Minus'], mac: [], sheets: [], context: 'grid' },
  { command: 'tools.palette', win: ['Alt-q'], mac: ['Ctrl-Alt-q'], sheets: ['Alt-/'], context: 'global' },
  { command: 'tools.shortcuts', win: ['Mod-/'], mac: ['Mod-/'], sheets: ['Mod-/'], context: 'global' },
  { command: 'tools.settings', win: ['Mod-,'], mac: ['Mod-,'], context: 'global' },
  { command: 'data.sort', win: [], context: 'grid' }
]

/** Handled by the grid and the in-cell editor (src/grid/keys.ts). */
export const BUILTIN: BuiltinShortcut[] = [
  { keys: 'Arrow keys', action: 'Move one cell', context: 'grid', status: 'done' },
  { keys: 'Shift-Arrow', action: 'Extend the selection by one cell', context: 'grid', status: 'done' },
  { keys: 'Mod-Arrow', mac: 'Mod-Arrow', action: 'Move to the edge of the data region', context: 'grid', status: 'done' },
  { keys: 'Mod-Shift-Arrow', action: 'Extend the selection to the edge of the data region', context: 'grid', status: 'done' },
  { keys: 'Home', mac: 'Fn-ArrowLeft', action: 'Move to the start of the row', context: 'grid', status: 'done' },
  { keys: 'Mod-Home', mac: 'Fn-Ctrl-ArrowLeft', action: 'Move to A1 (or the top-left of the unfrozen area)', context: 'grid', status: 'done' },
  { keys: 'Mod-End', mac: 'Fn-Ctrl-ArrowRight', action: 'Move to the last used cell', context: 'grid', status: 'done' },
  { keys: 'Mod-Shift-End', action: 'Extend the selection to the last used cell', context: 'grid', status: 'done' },
  { keys: 'PageDown / PageUp', action: 'Move one screen down / up', context: 'grid', status: 'done' },
  { keys: 'Alt-PageDown / Alt-PageUp', mac: 'Alt-PageDown', action: 'Move one screen right / left', context: 'grid', status: 'done' },
  { keys: 'Enter / Shift-Enter', action: 'Move down / up (within a selection: cycle through it)', context: 'grid', status: 'done' },
  { keys: 'Tab / Shift-Tab', action: 'Move right / left (within a selection: cycle through it)', context: 'grid', status: 'done' },
  { keys: 'Shift-Space', action: 'Select the entire row', context: 'grid', status: 'done' },
  { keys: 'Mod-Space', mac: 'Ctrl-Space', action: 'Select the entire column', context: 'grid', status: 'done' },
  { keys: 'Delete', action: 'Clear the contents of the selection', context: 'grid', status: 'done' },
  { keys: 'Backspace', action: 'Clear the active cell and start editing it', context: 'grid', status: 'done' },
  { keys: 'Escape', action: 'Cancel editing, or remove the copy marquee', context: 'grid', status: 'done' },
  { keys: 'F2', action: 'Edit the active cell; while editing, switch between Enter and Edit mode', context: 'editing', status: 'done' },
  { keys: 'Alt-Enter', mac: 'Ctrl-Alt-Enter', action: 'New line in the same cell', context: 'editing', status: 'done' },
  { keys: 'Mod-Enter', action: 'Fill the selected cells with the current entry', context: 'editing', status: 'done' },
  { keys: 'Mod-Shift-Enter', action: 'Enter an array formula (legacy CSE)', context: 'editing', status: 'done' },
  { keys: 'F4', mac: 'Mod-t', action: 'Cycle the reference at the cursor through absolute/relative', context: 'editing', status: 'done' },
  { keys: 'Arrow keys (in a formula)', action: 'Point mode: insert and move a cell reference', context: 'editing', status: 'done' },
  { keys: 'Shift-Arrow (in a formula)', action: 'Point mode: extend the inserted reference to a range', context: 'editing', status: 'done' },
  { keys: 'Mod-Shift-l', action: 'Toggle AutoFilter', context: 'grid', status: 'planned' },
  { keys: 'Mod-t', action: 'Create table', context: 'grid', status: 'planned' },
  { keys: 'Mod-e', action: 'Flash Fill', context: 'grid', status: 'planned' },
  { keys: 'Mod-k', action: 'Insert hyperlink', context: 'grid', status: 'planned' },
  { keys: 'Shift-F2', action: 'Insert or edit a note', context: 'grid', status: 'planned' },
  { keys: 'Mod-Alt-v', action: 'Paste Special dialog', context: 'grid', status: 'planned' },
  { keys: 'Mod-[', action: 'Select direct precedents', context: 'grid', status: 'planned' },
  { keys: 'Mod-]', action: 'Select direct dependents', context: 'grid', status: 'planned' },
  { keys: 'Alt-F1', action: 'Insert a chart of the current data', context: 'grid', status: 'planned' },
  { keys: 'Alt (ribbon KeyTips)', action: 'Show KeyTips for the ribbon', context: 'global', status: 'planned' },
  { keys: 'Mod-Shift-u', action: 'Expand or collapse the formula bar', context: 'grid', status: 'planned' }
]

/** Shortcuts the operating system keeps for itself, so Simple Cell can't receive them. */
export const RESERVED: { keys: string; os: string; note: string }[] = [
  { keys: 'Mod-Alt-Delete', os: 'Windows', note: 'Security screen' },
  { keys: 'Alt-F4', os: 'Windows/Linux', note: 'Closes the window (Simple Cell asks to save first)' },
  { keys: 'Mod-h', os: 'macOS', note: 'Hides the app: use Ctrl-H for Replace on macOS, as in Excel for Mac' },
  { keys: 'Mod-Space', os: 'macOS', note: 'Spotlight: select the column with Ctrl-Space' },
  { keys: 'Ctrl-Alt-Arrow', os: 'Linux (GNOME)', note: 'Switches workspaces' },
  { keys: 'Ctrl-Alt-F1…F12', os: 'Linux', note: 'Switches virtual terminals; use the Formulas tab for Calculate All' }
]
