// Central command registry. Ribbon buttons, menus, the command palette and keyboard shortcuts all dispatch through
// here. Shortcuts follow Microsoft Excel (docs/SHORTCUT_REFERENCE.md); users can rebind any of them.
import { areasOf } from '../state/view'
import { openCreateTable, toggleFilter } from './Tables'
import { get, set, toast, updateSettings, HOME, activeTab } from '../state/store'
import * as docs from '../state/documents'
import * as A from './actions'
import { getView, setView } from '../state/view'
import { MAX_COLS, MAX_ROWS } from '@/core/engine/values'
import { SHORTCUTS, type ShortcutDef } from './shortcuts'

export interface Command {
  id: string
  label: string
  /** Category shown in the command palette and shortcut list. */
  group: string
  /** Runs even while focus is in a text field or dialog. */
  global?: boolean
  /** Needs an open workbook. */
  needsBook?: boolean
  /** Not while a cell is being edited (most grid commands commit first instead). */
  run: (c: A.Ctx | null) => unknown
  keywords?: string
}

export const isMac = typeof navigator !== 'undefined' && /Mac/i.test(navigator.platform)

const commands = new Map<string, Command>()
const def = (c: Omit<Command, 'group'> & { group?: string }) => commands.set(c.id, { needsBook: true, group: c.id.split('.')[0], ...c })

export const getCommand = (id: string) => commands.get(id)
export const allCommands = () => [...commands.values()]

export async function runCommand(id: string) {
  const cmd = commands.get(id)
  if (!cmd) return
  let c = A.ctx()
  if (cmd.needsBook && !c) return
  // Commands act on committed content: finish an edit in progress first (except clipboard in the editor).
  if (c?.view.editing && !cmd.global) {
    const ok = await A.commitEdit(c, 'none')
    if (!ok) return
    c = A.ctx()
  }
  try {
    const r = cmd.run(c)
    if (r instanceof Promise) await r.catch((e: Error) => toast(e.message, 'error'))
  } catch (e) {
    toast((e as Error).message, 'error')
  }
}

// ---------- shortcuts ----------

/** The key bound to a command, with the user's overrides. */
export function shortcutFor(id: string): string | undefined {
  const o = get().settings.shortcuts[id]
  if (o !== undefined) return o || undefined
  return presetFor(id)?.[0]
}

function presetFor(id: string): string[] | undefined {
  const preset = get().settings.shortcutPreset
  const d = SHORTCUTS.find((s) => s.command === id)
  if (!d) return undefined
  const keys = preset === 'sheets' && d.sheets ? d.sheets : isMac ? (d.mac ?? d.win) : d.win
  return keys
}

/** All keys that trigger a command (a command can have several, e.g. Ctrl+Y and F4 for Redo). */
export function shortcutsFor(id: string): string[] {
  const o = get().settings.shortcuts[id]
  if (o !== undefined) return o ? [o] : []
  return presetFor(id) ?? []
}

export function formatShortcut(s: string | undefined): string {
  if (!s) return ''
  return s
    .split(' ')
    .map((chord) =>
      chord
        .split('-')
        .map((k) => {
          if (k === 'Mod') return isMac ? '⌘' : 'Ctrl'
          if (k === 'Shift') return isMac ? '⇧' : 'Shift'
          if (k === 'Alt') return isMac ? '⌥' : 'Alt'
          if (k === 'Ctrl') return isMac ? '⌃' : 'Ctrl'
          if (k === 'Enter') return isMac ? '↩' : 'Enter'
          if (k === 'Space') return 'Space'
          if (k === 'Minus') return '-'
          if (k === 'Plus') return '+'
          return k.length === 1 ? k.toUpperCase() : k
        })
        .join(isMac ? '' : '+')
    )
    .join(', ')
}

/** Normalize a keyboard event to the registry's key format (Mod-Shift-l, Alt-=, F9, Mod-;). */
export function eventToShortcut(e: KeyboardEvent): string {
  const parts: string[] = []
  const mod = isMac ? e.metaKey : e.ctrlKey
  if (mod) parts.push('Mod')
  if (isMac && e.ctrlKey) parts.push('Ctrl')
  if (e.altKey) parts.push('Alt')
  let key = e.key
  // Physical keys for letters/digits/punctuation with modifiers so Shift/Alt/layouts don't change them.
  const PUNCT: Record<string, string> = { Semicolon: ';', Quote: "'", Backquote: '`', Equal: '=', Minus: 'Minus', Comma: ',', Period: '.', Slash: '/', Backslash: '\\', BracketLeft: '[', BracketRight: ']' }
  if ((mod || e.altKey) && /^(Key[A-Z]|Digit\d)$/.test(e.code)) key = e.code.replace(/^(Key|Digit)/, '').toLowerCase()
  else if ((mod || e.altKey) && PUNCT[e.code]) key = PUNCT[e.code]
  else if (key === ' ') key = 'Space'
  else if (key === '+') key = 'Plus'
  else if (key === '-') key = 'Minus'
  else if (key.length === 1) key = key.toLowerCase()
  if (e.shiftKey) parts.push('Shift')
  parts.push(key)
  return parts.join('-')
}

/** Keydown handler installed at window level (capture), so registry shortcuts win. */
export function handleKeydown(e: KeyboardEvent) {
  if (e.isComposing) return
  const combo = eventToShortcut(e)
  const target = e.target as HTMLElement | null
  const inGridInput = !!target?.hasAttribute?.('data-grid-focus')
  const inField = !inGridInput && !!target?.closest('input, textarea, select, .cm-editor, [role=dialog], [role=menu]')
  const editing = !!A.ctx()?.view.editing
  for (const c of commands.values()) {
    if (!shortcutsFor(c.id).includes(combo)) continue
    if (inField && !c.global) return
    if (c.needsBook && !A.ctx()) return
    // While a cell is edited, only commands that make sense in the editor run (the rest are text keys).
    if (editing && !EDITOR_OK.has(c.id)) return
    e.preventDefault()
    e.stopPropagation()
    void runCommand(c.id)
    return
  }
}

const EDITOR_OK = new Set(['file.save', 'file.saveAs', 'file.open', 'file.new', 'file.print', 'tools.palette', 'formulas.insertFunction', 'edit.insertDate', 'edit.insertTime', 'fmt.bold', 'fmt.italic', 'fmt.underline'])

// ---------- helpers ----------

function needsSelection(c: A.Ctx | null): A.Ctx {
  if (!c) throw new Error('Open a workbook first')
  return A.ctx() ?? c
}

const lastRange = (c: A.Ctx) => c.view.sel.ranges[c.view.sel.ranges.length - 1]

function zoomBy(c: A.Ctx, f: number) {
  setView(c.tab, (v) => ({ zoom: Math.min(4, Math.max(0.1, Math.round(v.zoom * f * 100) / 100)) }))
}

function cycleTab(dir: 1 | -1) {
  const ids = [HOME, ...get().tabs.map((t) => t.id)]
  const i = ids.indexOf(get().active)
  const next = ids[(i + dir + ids.length) % ids.length]
  if (next === HOME) set({ active: HOME })
  else docs.activate(next)
}

function cycleSheet(c: A.Ctx, dir: 1 | -1) {
  const sheets = c.book.meta.sheets.filter((s) => !s.hidden)
  const i = sheets.findIndex((s) => s.id === c.view.sheet)
  const next = sheets[(i + dir + sheets.length) % sheets.length]
  if (next) docs.switchSheet(c.tab, next.id)
}

async function insertNow(c: A.Ctx, kind: 'date' | 'time') {
  const d = new Date()
  const text = kind === 'date' ? `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}` : `${d.getHours() % 12 || 12}:${String(d.getMinutes()).padStart(2, '0')} ${d.getHours() < 12 ? 'AM' : 'PM'}`
  const v = getView(c.tab)
  if (v.editing) {
    const e = v.editing
    setView(c.tab, { editing: { ...e, text: e.text.slice(0, e.caret) + text + e.text.slice(e.caret), caret: e.caret + text.length } })
  } else setView(c.tab, { editing: { text, where: 'cell', enterMode: true, caret: text.length } })
}

/** AutoSum (Alt+=): SUM of the numbers above (or to the left of) the active cell. */
async function autoSum(c: A.Ctx, fn = 'SUM') {
  const r = lastRange(c)
  const { r: ar, c: ac } = c.view.sel.active
  const single = r.r0 === r.r1 && r.c0 === r.c1
  if (!single) {
    // Selected numbers: put the totals below each column.
    const sheet = c.view.sheet
    const rows = Math.min(r.r1, MAX_ROWS - 2)
    await A.edit(c, async (b) => {
      for (let col = r.c0; col <= r.c1; col++) {
        const res = await b.api.input(sheet, rows + 1, col, `=${fn}(${A.areaName({ r0: r.r0, c0: col, r1: rows, c1: col })})`)
        if (!res.ok) return res
      }
      return { ok: true }
    })
    return
  }
  // Find the run of numbers above, else to the left.
  const book = c.book
  const isNum = (rr: number, cc: number) => {
    const v = book.cell(c.view.sheet, rr, cc)
    return v?.k === 'n' || v?.k === 'd'
  }
  let top = ar - 1
  while (top >= 0 && !isNum(top, ac) && !book.cell(c.view.sheet, top, ac)?.t) top--
  let start = top
  while (start - 1 >= 0 && isNum(start - 1, ac)) start--
  let range = ''
  if (top >= 0 && isNum(top, ac)) range = A.areaName({ r0: start, c0: ac, r1: top, c1: ac })
  else {
    let left = ac - 1
    let s2 = left
    while (s2 - 1 >= 0 && isNum(ar, s2 - 1)) s2--
    if (left >= 0 && isNum(ar, left)) range = A.areaName({ r0: ar, c0: s2, r1: ar, c1: left })
    void left
  }
  const text = `=${fn}(${range})`
  setView(c.tab, { editing: { text, where: 'cell', enterMode: false, caret: text.length - 1 } })
}

// ---------- definitions ----------

// File
def({ id: 'file.new', label: 'New workbook', global: true, needsBook: false, run: () => docs.newWorkbook() })
def({ id: 'file.open', label: 'Open…', global: true, needsBook: false, run: () => docs.openDialog() })
def({ id: 'file.save', label: 'Save', global: true, run: () => docs.save() })
def({ id: 'file.saveAs', label: 'Save as…', global: true, run: () => docs.save(undefined, { as: true }) })
def({ id: 'file.close', label: 'Close workbook', global: true, needsBook: false, run: () => (get().active !== HOME ? docs.closeTab(get().active) : undefined) })
def({ id: 'file.home', label: 'Home', needsBook: false, global: true, run: () => set({ active: HOME, backstage: false }) })
def({ id: 'file.backstage', label: 'File menu', run: () => set((s) => ({ backstage: !s.backstage })) })
def({ id: 'file.print', label: 'Print…', global: true, run: () => toast('Printing arrives with the page layout engine (see IMPLEMENTATION_STATUS.md).', 'info') })
def({ id: 'tab.next', label: 'Next workbook', group: 'view', global: true, needsBook: false, run: () => cycleTab(1) })
def({ id: 'tab.prev', label: 'Previous workbook', group: 'view', global: true, needsBook: false, run: () => cycleTab(-1) })

// Edit
def({ id: 'edit.undo', label: 'Undo', run: (c) => A.undo(needsSelection(c)) })
def({ id: 'edit.redo', label: 'Redo', run: (c) => A.redo(needsSelection(c)) })
def({ id: 'edit.copy', label: 'Copy', run: (c) => A.copy(needsSelection(c)) })
def({ id: 'edit.cut', label: 'Cut', run: (c) => A.copy(needsSelection(c), true) })
def({ id: 'edit.paste', label: 'Paste', run: (c) => A.paste(needsSelection(c)) })
def({ id: 'edit.pasteValues', label: 'Paste values', run: (c) => A.paste(needsSelection(c), 'values') })
def({ id: 'edit.pasteFormats', label: 'Paste formatting', run: (c) => A.paste(needsSelection(c), 'formats') })
def({ id: 'edit.pasteFormulas', label: 'Paste formulas', run: (c) => A.paste(needsSelection(c), 'formulas') })
def({ id: 'edit.clear', label: 'Clear contents', run: (c) => A.clearContents(needsSelection(c)) })
def({ id: 'edit.clearFormats', label: 'Clear formats', run: (c) => A.edit(needsSelection(c), (b) => b.api.clear(A.areasOf(c!.view), 'formats')) })
def({ id: 'edit.clearAll', label: 'Clear all', run: (c) => A.edit(needsSelection(c), (b) => b.api.clear(A.areasOf(c!.view), 'all')) })
def({ id: 'edit.fillDown', label: 'Fill down', run: (c) => A.edit(needsSelection(c), (b) => b.api.fill({ sheet: c!.view.sheet, ...lastRange(c!) }, 'down')) })
def({ id: 'edit.fillRight', label: 'Fill right', run: (c) => A.edit(needsSelection(c), (b) => b.api.fill({ sheet: c!.view.sheet, ...lastRange(c!) }, 'right')) })
def({ id: 'edit.fillUp', label: 'Fill up', run: (c) => A.edit(needsSelection(c), (b) => b.api.fill({ sheet: c!.view.sheet, ...lastRange(c!) }, 'up')) })
def({ id: 'edit.fillLeft', label: 'Fill left', run: (c) => A.edit(needsSelection(c), (b) => b.api.fill({ sheet: c!.view.sheet, ...lastRange(c!) }, 'left')) })
def({ id: 'edit.find', label: 'Find…', run: () => set({ dialog: { kind: 'find', replace: false } }) })
def({ id: 'edit.replace', label: 'Replace…', run: () => set({ dialog: { kind: 'find', replace: true } }) })
def({ id: 'edit.goto', label: 'Go to…', run: () => set({ dialog: { kind: 'goto' } }) })
def({ id: 'edit.selectAll', label: 'Select all', run: (c) => A.selectAll(needsSelection(c)) })
def({ id: 'edit.selectRow', label: 'Select row', run: (c) => A.selectRows(needsSelection(c)) })
def({ id: 'edit.selectColumn', label: 'Select column', run: (c) => A.selectCols(needsSelection(c)) })
def({ id: 'edit.insertDate', label: 'Insert current date', global: true, run: (c) => insertNow(needsSelection(c), 'date') })
def({ id: 'edit.insertTime', label: 'Insert current time', global: true, run: (c) => insertNow(needsSelection(c), 'time') })
def({ id: 'edit.editCell', label: 'Edit active cell', run: (c) => A.startEdit(needsSelection(c)) })
def({ id: 'edit.insertCells', label: 'Insert cells…', run: (c) => insertForSelection(needsSelection(c)) })
def({ id: 'edit.deleteCells', label: 'Delete cells…', run: (c) => deleteForSelection(needsSelection(c)) })
def({ id: 'edit.insertRows', label: 'Insert sheet rows', run: (c) => A.insertAt(needsSelection(c), 'row') })
def({ id: 'edit.insertCols', label: 'Insert sheet columns', run: (c) => A.insertAt(needsSelection(c), 'col') })
def({ id: 'edit.deleteRows', label: 'Delete sheet rows', run: (c) => A.deleteAt(needsSelection(c), 'row') })
def({ id: 'edit.deleteCols', label: 'Delete sheet columns', run: (c) => A.deleteAt(needsSelection(c), 'col') })

async function insertForSelection(c: A.Ctx) {
  const r = lastRange(c)
  if (r.c0 === 0 && r.c1 === MAX_COLS - 1) return A.insertAt(c, 'row')
  if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) return A.insertAt(c, 'col')
  set({ dialog: { kind: 'shift-cells', mode: 'insert' } })
}
async function deleteForSelection(c: A.Ctx) {
  const r = lastRange(c)
  if (r.c0 === 0 && r.c1 === MAX_COLS - 1) return A.deleteAt(c, 'row')
  if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) return A.deleteAt(c, 'col')
  set({ dialog: { kind: 'shift-cells', mode: 'delete' } })
}

// Format
def({ id: 'fmt.bold', label: 'Bold', global: false, run: (c) => A.toggleSel(needsSelection(c), 'bold') })
def({ id: 'fmt.italic', label: 'Italic', run: (c) => A.toggleSel(needsSelection(c), 'italic') })
def({ id: 'fmt.underline', label: 'Underline', run: async (c) => {
  const cc = needsSelection(c)
  const s = await A.activeStyle(cc)
  return A.formatSel(cc, { underline: s.underline ? undefined : 'single' })
} })
def({ id: 'fmt.strike', label: 'Strikethrough', run: (c) => A.toggleSel(needsSelection(c), 'strike') })
def({ id: 'fmt.wrap', label: 'Wrap text', run: (c) => A.toggleSel(needsSelection(c), 'wrap') })
def({ id: 'fmt.alignLeft', label: 'Align left', run: (c) => A.formatSel(needsSelection(c), { hAlign: 'left' }) })
def({ id: 'fmt.alignCenter', label: 'Center', run: (c) => A.formatSel(needsSelection(c), { hAlign: 'center' }) })
def({ id: 'fmt.alignRight', label: 'Align right', run: (c) => A.formatSel(needsSelection(c), { hAlign: 'right' }) })
def({ id: 'fmt.alignTop', label: 'Top align', run: (c) => A.formatSel(needsSelection(c), { vAlign: 'top' }) })
def({ id: 'fmt.alignMiddle', label: 'Middle align', run: (c) => A.formatSel(needsSelection(c), { vAlign: 'center' }) })
def({ id: 'fmt.alignBottom', label: 'Bottom align', run: (c) => A.formatSel(needsSelection(c), { vAlign: undefined }) })
def({ id: 'fmt.indent', label: 'Increase indent', run: async (c) => {
  const cc = needsSelection(c)
  const s = await A.activeStyle(cc)
  return A.formatSel(cc, { indent: Math.min(15, (s.indent ?? 0) + 1), hAlign: s.hAlign === 'right' ? 'right' : 'left' })
} })
def({ id: 'fmt.outdent', label: 'Decrease indent', run: async (c) => {
  const cc = needsSelection(c)
  const s = await A.activeStyle(cc)
  return A.formatSel(cc, { indent: Math.max(0, (s.indent ?? 0) - 1) || undefined })
} })
def({ id: 'fmt.mergeCenter', label: 'Merge & center', run: async (c) => {
  const cc = needsSelection(c)
  const r = lastRange(cc)
  const merged = cc.book.sheetMeta(cc.view.sheet)!.merges.some((m) => m.r0 === r.r0 && m.c0 === r.c0 && m.r1 === r.r1 && m.c1 === r.c1)
  return A.edit(cc, (b) => (merged ? b.api.unmerge({ sheet: cc.view.sheet, ...r }) : b.api.merge({ sheet: cc.view.sheet, ...r }, 'center')))
} })
def({ id: 'fmt.merge', label: 'Merge cells', run: (c) => A.edit(needsSelection(c), (b) => b.api.merge({ sheet: c!.view.sheet, ...lastRange(c!) }, 'all')) })
def({ id: 'fmt.mergeAcross', label: 'Merge across', run: (c) => A.edit(needsSelection(c), (b) => b.api.merge({ sheet: c!.view.sheet, ...lastRange(c!) }, 'across')) })
def({ id: 'fmt.unmerge', label: 'Unmerge cells', run: (c) => A.edit(needsSelection(c), (b) => b.api.unmerge({ sheet: c!.view.sheet, ...lastRange(c!) })) })
def({ id: 'fmt.general', label: 'General number format', run: (c) => A.formatSel(needsSelection(c), { numFmt: undefined }) })
def({ id: 'fmt.number', label: 'Number format (2 decimals, thousands separator)', run: (c) => A.formatSel(needsSelection(c), { numFmt: '#,##0.00' }) })
def({ id: 'fmt.currency', label: 'Currency format', run: (c) => A.formatSel(needsSelection(c), { numFmt: '"$"#,##0.00_);\\("$"#,##0.00\\)' }) })
def({ id: 'fmt.accounting', label: 'Accounting format', run: (c) => A.formatSel(needsSelection(c), { numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' }) })
def({ id: 'fmt.percent', label: 'Percent format', run: (c) => A.formatSel(needsSelection(c), { numFmt: '0%' }) })
def({ id: 'fmt.scientific', label: 'Scientific format', run: (c) => A.formatSel(needsSelection(c), { numFmt: '0.00E+00' }) })
def({ id: 'fmt.date', label: 'Date format', run: (c) => A.formatSel(needsSelection(c), { numFmt: 'd-mmm-yy' }) })
def({ id: 'fmt.shortDate', label: 'Short date format', run: (c) => A.formatSel(needsSelection(c), { numFmt: 'm/d/yyyy' }) })
def({ id: 'fmt.longDate', label: 'Long date format', run: (c) => A.formatSel(needsSelection(c), { numFmt: '[$-x-sysdate]dddd, mmmm d, yyyy' }) })
def({ id: 'fmt.time', label: 'Time format', run: (c) => A.formatSel(needsSelection(c), { numFmt: 'h:mm AM/PM' }) })
def({ id: 'fmt.text', label: 'Text format', run: (c) => A.formatSel(needsSelection(c), { numFmt: '@' }) })
def({ id: 'fmt.thousands', label: 'Comma style', run: (c) => A.formatSel(needsSelection(c), { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }) })
def({ id: 'fmt.moreDecimals', label: 'Increase decimal', run: (c) => changeDecimals(needsSelection(c), 1) })
def({ id: 'fmt.lessDecimals', label: 'Decrease decimal', run: (c) => changeDecimals(needsSelection(c), -1) })
def({ id: 'fmt.outlineBorder', label: 'Outline border', run: (c) => outline(needsSelection(c)) })
def({ id: 'fmt.noBorder', label: 'Remove borders', run: (c) => A.formatSel(needsSelection(c), { borderTop: undefined, borderBottom: undefined, borderLeft: undefined, borderRight: undefined, borderDiagDown: undefined, borderDiagUp: undefined }) })
def({ id: 'fmt.cells', label: 'Format cells…', run: () => set({ dialog: { kind: 'format-cells' } }) })
def({ id: 'fmt.font', label: 'Format font…', run: () => set({ dialog: { kind: 'format-cells', tab: 'font' } }) })
def({ id: 'fmt.rowHeight', label: 'Row height…', run: () => set({ dialog: { kind: 'row-height' } }) })
def({ id: 'fmt.colWidth', label: 'Column width…', run: () => set({ dialog: { kind: 'col-width' } }) })
def({ id: 'fmt.hideRows', label: 'Hide rows', run: (c) => hideSel(needsSelection(c), 'row', true) })
def({ id: 'fmt.unhideRows', label: 'Unhide rows', run: (c) => hideSel(needsSelection(c), 'row', false) })
def({ id: 'fmt.hideCols', label: 'Hide columns', run: (c) => hideSel(needsSelection(c), 'col', true) })
def({ id: 'fmt.unhideCols', label: 'Unhide columns', run: (c) => hideSel(needsSelection(c), 'col', false) })

async function changeDecimals(c: A.Ctx, d: number) {
  const s = await A.activeStyle(c)
  let code = s.numFmt ?? 'General'
  if (code === 'General') {
    const v = c.book.cell(c.view.sheet, c.view.sel.active.r, c.view.sel.active.c)
    const decs = /\.(\d+)/.exec(v?.t ?? '')?.[1].length ?? 0
    code = decs ? '0.' + '0'.repeat(decs) : '0'
  }
  const next = code
    .split(';')
    .map((sec) => {
      if (/"/.test(sec) && !/0/.test(sec)) return sec
      if (d > 0) return /\.0+/.test(sec) ? sec.replace(/(\.0+)/, '$10') : sec.replace(/(0)(?!.*0)/, '$1.0')
      return sec.replace(/\.(0+)/, (_m, z: string) => (z.length > 1 ? '.' + z.slice(1) : ''))
    })
    .join(';')
  return A.formatSel(c, { numFmt: next })
}

function outline(c: A.Ctx) {
  const r = lastRange(c)
  const side = { style: 'thin' as const }
  const areas = (o: Partial<typeof r>) => [{ sheet: c.view.sheet, ...r, ...o }]
  return A.edit(c, async (b) => {
    await b.api.formatCells(areas({ r1: r.r0 }), { borderTop: side })
    await b.api.formatCells(areas({ r0: r.r1 }), { borderBottom: side })
    await b.api.formatCells(areas({ c1: r.c0 }), { borderLeft: side })
    return b.api.formatCells(areas({ c0: r.c1 }), { borderRight: side })
  })
}

function hideSel(c: A.Ctx, axis: 'row' | 'col', hidden: boolean) {
  const idx: number[] = []
  for (const r of c.view.sel.ranges) {
    if (axis === 'row') for (let i = r.r0; i <= Math.min(r.r1, r.r0 + 100_000); i++) idx.push(i)
    else for (let i = r.c0; i <= r.c1; i++) idx.push(i)
  }
  return A.edit(c, (b) => b.api.hide(c.view.sheet, axis, idx, hidden))
}

// Formulas
def({ id: 'formulas.autosum', label: 'AutoSum', run: (c) => autoSum(needsSelection(c)) })
def({ id: 'formulas.average', label: 'AutoSum: Average', run: (c) => autoSum(needsSelection(c), 'AVERAGE') })
def({ id: 'formulas.count', label: 'AutoSum: Count numbers', run: (c) => autoSum(needsSelection(c), 'COUNT') })
def({ id: 'formulas.max', label: 'AutoSum: Max', run: (c) => autoSum(needsSelection(c), 'MAX') })
def({ id: 'formulas.min', label: 'AutoSum: Min', run: (c) => autoSum(needsSelection(c), 'MIN') })
def({ id: 'formulas.insertFunction', label: 'Insert function…', global: true, run: () => set({ dialog: { kind: 'insert-function' } }) })
def({ id: 'formulas.names', label: 'Name manager…', run: () => set({ dialog: { kind: 'names' } }) })
def({ id: 'formulas.calculate', label: 'Calculate now', run: (c) => A.edit(needsSelection(c), (b) => b.api.calculate(false)) })
def({ id: 'formulas.calculateAll', label: 'Calculate all (full)', run: (c) => A.edit(needsSelection(c), (b) => b.api.calculate(true)) })
def({ id: 'formulas.calcAuto', label: 'Calculation options: Automatic', run: (c) => needsSelection(c).book.api.setCalc('auto').then(() => c!.book.refreshMeta()) })
def({ id: 'formulas.calcManual', label: 'Calculation options: Manual', run: (c) => needsSelection(c).book.api.setCalc('manual').then(() => c!.book.refreshMeta()) })
def({ id: 'formulas.showFormulas', label: 'Show formulas', run: () => set((s) => ({ showFormulas: !s.showFormulas })) })

// View
def({ id: 'view.freezePanes', label: 'Freeze panes', run: (c) => {
  const cc = needsSelection(c)
  const m = cc.book.sheetMeta(cc.view.sheet)!
  const { r, c: col } = cc.view.sel.active
  return A.edit(cc, (b) => (m.freeze.rows || m.freeze.cols ? b.api.freeze(cc.view.sheet, 0, 0) : b.api.freeze(cc.view.sheet, r, col)))
} })
def({ id: 'view.freezeTopRow', label: 'Freeze top row', run: (c) => A.edit(needsSelection(c), (b) => b.api.freeze(c!.view.sheet, 1, 0)) })
def({ id: 'view.freezeFirstColumn', label: 'Freeze first column', run: (c) => A.edit(needsSelection(c), (b) => b.api.freeze(c!.view.sheet, 0, 1)) })
def({ id: 'view.unfreeze', label: 'Unfreeze panes', run: (c) => A.edit(needsSelection(c), (b) => b.api.freeze(c!.view.sheet, 0, 0)) })
def({ id: 'view.gridlines', label: 'Show gridlines', run: (c) => {
  const cc = needsSelection(c)
  const m = cc.book.sheetMeta(cc.view.sheet)!
  return A.edit(cc, (b) => b.api.sheetProps(cc.view.sheet, { showGridlines: !m.showGridlines }))
} })
def({ id: 'view.headings', label: 'Show headings', run: () => set((s) => ({ showHeadings: !s.showHeadings })) })
def({ id: 'view.formulaBar', label: 'Show formula bar', run: () => set((s) => ({ showFormulaBar: !s.showFormulaBar })) })
def({ id: 'view.zoomIn', label: 'Zoom in', run: (c) => zoomBy(needsSelection(c), 1.1) })
def({ id: 'view.zoomOut', label: 'Zoom out', run: (c) => zoomBy(needsSelection(c), 1 / 1.1) })
def({ id: 'view.zoom100', label: 'Zoom to 100%', run: (c) => setView(needsSelection(c).tab, { zoom: 1 }) })
def({ id: 'view.theme', label: 'Toggle dark mode', needsBook: false, global: true, run: () => updateSettings({ theme: get().dark ? 'light' : 'dark' }) })
def({ id: 'view.ribbon', label: 'Collapse or expand the ribbon', needsBook: false, global: true, run: () => updateSettings({ ribbonCollapsed: !get().settings.ribbonCollapsed }) })
def({ id: 'view.ribbonMode', label: 'Switch between the full ribbon and the compact toolbar', needsBook: false, run: () => updateSettings({ ribbonMode: get().settings.ribbonMode === 'compact' ? 'expanded' : 'compact' }) })

// Sheets
def({ id: 'sheet.insert', label: 'Insert worksheet', run: async (c) => {
  const cc = needsSelection(c)
  const r = await cc.book.api.addSheet(undefined, cc.book.meta.sheets.findIndex((s) => s.id === cc.view.sheet) + 1)
  if (!r.ok) throw new Error(r.error)
  docs.markDirty(cc.tab)
  docs.switchSheet(cc.tab, r.id)
} })
def({ id: 'sheet.next', label: 'Next sheet', run: (c) => cycleSheet(needsSelection(c), 1) })
def({ id: 'sheet.prev', label: 'Previous sheet', run: (c) => cycleSheet(needsSelection(c), -1) })

// Data
def({ id: 'data.sortAsc', label: 'Sort A to Z', run: (c) => quickSort(needsSelection(c), 1) })
def({ id: 'data.sortDesc', label: 'Sort Z to A', run: (c) => quickSort(needsSelection(c), -1) })
def({ id: 'data.sort', label: 'Custom sort…', run: () => set({ dialog: { kind: 'sort' } }) })
def({ id: 'data.table', label: 'Create table…', keywords: 'format as table ctrl+t list', run: () => void openCreateTable() })
def({ id: 'data.filter', label: 'Filter (turn on or off)', keywords: 'autofilter', run: () => void toggleFilter() })
def({
  id: 'data.clearFilter',
  label: 'Clear filters',
  run: (c) => {
    const x = needsSelection(c)
    for (const f of x.book.sheetMeta(x.view.sheet)?.filters ?? []) void A.edit(x, (b) => b.api.clearFilters(x.view.sheet, f.key))
  }
})
def({
  id: 'data.reapply',
  label: 'Reapply filters',
  run: (c) => {
    const x = needsSelection(c)
    for (const f of x.book.sheetMeta(x.view.sheet)?.filters ?? []) void A.edit(x, (b) => b.api.reapplyFilter(x.view.sheet, f.key))
  }
})
def({
  id: 'table.convert',
  label: 'Convert table to range',
  run: (c) => {
    const x = needsSelection(c)
    const { r, c: col } = x.view.sel.active
    const t = x.book.meta.tables.find((t) => t.sheet === x.view.sheet && r >= t.area.r0 && r <= t.area.r1 && col >= t.area.c0 && col <= t.area.c1)
    if (!t) return toast('Select a cell in a table first.', 'info')
    void A.edit(x, (b) => b.api.convertTableToRange(t.name))
  }
})
def({ id: 'insert.note', label: 'New or edit note', keywords: 'comment annotation', run: (c) => (needsSelection(c), set({ dialog: { kind: 'note' } })) })
def({
  id: 'note.delete',
  label: 'Delete note',
  run: (c) => {
    const x = needsSelection(c)
    void A.edit(x, (b) => b.api.setNote(x.view.sheet, x.view.sel.active.r, x.view.sel.active.c, null))
  }
})
def({ id: 'insert.link', label: 'Insert link…', keywords: 'hyperlink url web address', run: (c) => (needsSelection(c), set({ dialog: { kind: 'link' } })) })
def({
  id: 'link.remove',
  label: 'Remove link',
  run: (c) => {
    const x = needsSelection(c)
    void A.edit(x, (b) => b.api.setLink(areasOf(x.view), null))
  }
})
def({ id: 'edit.pasteSpecial', label: 'Paste special…', keywords: 'transpose values operation skip blanks', run: (c) => (needsSelection(c), set({ dialog: { kind: 'paste-special' } })) })
def({ id: 'fmt.painter', label: 'Format painter', keywords: 'copy formatting', run: (c) => A.startPainter(needsSelection(c), false) })
def({ id: 'edit.gotoSpecial', label: 'Go to special…', keywords: 'select blanks constants formulas visible cells', run: () => set({ dialog: { kind: 'goto-special' } }) })
def({ id: 'select.visible', label: 'Select visible cells only', run: (c) => A.goToSpecial(needsSelection(c), 'visible') })
def({ id: 'select.formulas', label: 'Select formulas', run: (c) => A.goToSpecial(needsSelection(c), 'formulas') })
def({ id: 'select.notes', label: 'Select notes', run: (c) => A.goToSpecial(needsSelection(c), 'notes') })
def({ id: 'select.condFormats', label: 'Select conditional formatting', run: (c) => A.goToSpecial(needsSelection(c), 'condFormats') })
def({ id: 'select.validation', label: 'Select data validation', run: (c) => A.goToSpecial(needsSelection(c), 'validation') })
def({ id: 'select.constants', label: 'Select constants', run: (c) => A.goToSpecial(needsSelection(c), 'constants') })
def({ id: 'data.flashFill', label: 'Flash Fill', keywords: 'fill by example pattern split names', run: async (c) => {
    const x = needsSelection(c)
    const { r, c: col } = x.view.sel.active
    const res = await x.book.api.flashFill(x.view.sheet, r, col)
    if (!res.ok) return toast(res.error, 'info')
    A.markDirtyTab(x.tab)
    toast(`Flash Fill filled ${res.filled} ${res.filled === 1 ? 'cell' : 'cells'}.`, 'success')
  }
})
def({ id: 'data.validation', label: 'Data validation…', keywords: 'dropdown list restrict allowed values', run: () => set({ dialog: { kind: 'data-validation' } }) })
def({
  id: 'data.circleInvalid',
  label: 'Circle invalid data',
  run: async (c) => {
    const n = await A.circleInvalid(needsSelection(c))
    if (!n) toast('Every cell with a validation rule has a valid value.', 'info')
  }
})
def({ id: 'data.clearCircles', label: 'Clear validation circles', run: (c) => A.clearCircles(needsSelection(c)) })

async function quickSort(c: A.Ctx, order: 1 | -1) {
  let r = lastRange(c)
  // A single cell sorts its whole data region (current region), with a header row if the first row is text.
  if (r.r0 === r.r1 && r.c0 === r.c1) r = await currentRegion(c)
  const u = c.book.sheetMeta(c.view.sheet)!.used
  r = { ...r, r1: Math.min(r.r1, u.r), c1: Math.min(r.c1, u.c) }
  const header = guessHeader(c, r)
  return A.edit(c, (b) => b.api.sort({ sheet: c.view.sheet, ...r }, [{ index: c.view.sel.active.c, order }], header))
}

export function guessHeader(c: A.Ctx, r: { r0: number; c0: number; r1: number; c1: number }) {
  let textTop = 0
  let numBelow = 0
  for (let col = r.c0; col <= r.c1; col++) {
    if (c.book.cell(c.view.sheet, r.r0, col)?.k === 's') textTop++
    const below = c.book.cell(c.view.sheet, r.r0 + 1, col)?.k
    if (below === 'n' || below === 'd') numBelow++
  }
  return textTop > 0 && textTop === r.c1 - r.c0 + 1 && (numBelow > 0 || r.r1 > r.r0)
}

/** The block of non-empty cells around the active cell (Ctrl+A's first step, sorting a single cell). */
export async function currentRegion(c: A.Ctx) {
  const { r, c: col } = c.view.sel.active
  let r0 = r
  let r1 = r
  let c0 = col
  let c1 = col
  const filled = (rr: number, cc: number) => !!c.book.cell(c.view.sheet, rr, cc)?.t
  for (let grow = true; grow; ) {
    grow = false
    if (r0 > 0 && rangeAny(c0 - 1, c1 + 1, (cc) => filled(r0 - 1, cc))) {
      r0--
      grow = true
    }
    if (rangeAny(c0 - 1, c1 + 1, (cc) => filled(r1 + 1, cc))) {
      r1++
      grow = true
    }
    if (c0 > 0 && rangeAny(r0 - 1, r1 + 1, (rr) => filled(rr, c0 - 1))) {
      c0--
      grow = true
    }
    if (rangeAny(r0 - 1, r1 + 1, (rr) => filled(rr, c1 + 1))) {
      c1++
      grow = true
    }
    if (r1 - r0 > 200_000) break
  }
  return { r0, c0, r1, c1 }
}

function rangeAny(a: number, b: number, f: (i: number) => boolean) {
  for (let i = Math.max(0, a); i <= b; i++) if (f(i)) return true
  return false
}

// Tools
def({ id: 'tools.palette', label: 'Command palette', global: true, needsBook: false, run: () => set({ dialog: { kind: 'palette' } }) })
def({ id: 'tools.settings', label: 'Settings', global: true, needsBook: false, run: () => set({ dialog: { kind: 'settings' } }) })
def({ id: 'tools.shortcuts', label: 'Keyboard shortcuts', global: true, needsBook: false, run: () => set({ dialog: { kind: 'shortcuts' } }) })
def({ id: 'tools.compat', label: 'Compatibility report', run: () => {
  const t = activeTab()
  if (t) set({ dialog: { kind: 'compat', tabId: t.id } })
} })
def({ id: 'tools.about', label: 'About Simple Cell', needsBook: false, run: () => set({ dialog: { kind: 'about' } }) })
def({ id: 'tools.checkUpdates', label: 'Check for updates', needsBook: false, run: () => window.simpleCell.checkForUpdates() })

export type { ShortcutDef }
