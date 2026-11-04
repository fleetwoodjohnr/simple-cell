// Actions on the active workbook view: moving the selection, editing cells, clipboard, structure.
// The grid, the ribbon, menus and keyboard shortcuts all call these.
import { MAX_COLS, MAX_ROWS, type Area } from '@/core/engine/values'
import type { Style } from '@/core/model/styles'
import { areaName, cellName, colName, parseArea } from '@/core/refs'
import { activeBook, books, get, set, toast } from '../state/store'
import { areasOf, cellSel, clampCell, getView, rectOf, setView, type Sel, type TabView } from '../state/view'
import { markDirty } from '../state/documents'
import { axesFor } from '@/grid/layout'
import type { Book } from '../state/client'
import type { DvAlert } from '@/core/validation'
import { showValidationAlert } from './Validation'

export interface Ctx {
  tab: string
  book: Book
  view: TabView
}

export function ctx(): Ctx | null {
  const tab = get().active
  const book = books.get(tab)
  const view = getView(tab)
  return book && view ? { tab, book, view } : null
}

/** Run a core edit; report errors; mark the workbook dirty. */
export async function edit(c: Ctx, fn: (b: Book) => Promise<{ ok: boolean; error?: string } | unknown>): Promise<boolean> {
  const r = (await fn(c.book)) as { ok?: boolean; error?: string } | undefined
  if (r && r.ok === false) {
    toast(r.error ?? 'That didn’t work', 'error')
    return false
  }
  markDirty(c.tab)
  return true
}

// ---------- selection ----------

export function setSel(c: Ctx, sel: Sel) {
  setView(c.tab, { sel })
}

function axes(c: Ctx) {
  const m = c.book.sheetMeta(c.view.sheet)!
  return axesFor(m)
}

/** Arrow keys: move the active cell (Shift extends the selection). Hidden rows/columns are skipped. */
export function move(c: Ctx, dr: number, dc: number, extend = false) {
  const { rows, cols } = axes(c)
  const s = c.view.sel
  const from = extend ? s.focus : s.active
  let r = from.r
  let col = from.c
  // Moving out of a merged cell jumps past the merge.
  const merges = c.book.sheetMeta(c.view.sheet)!.merges
  const m = merges.find((x) => r >= x.r0 && r <= x.r1 && col >= x.c0 && col <= x.c1)
  if (m && !extend) {
    if (dr > 0) r = m.r1
    if (dc > 0) col = m.c1
    if (dr < 0) r = m.r0
    if (dc < 0) col = m.c0
  }
  if (dr) r = rows.visible(r + dr, dr > 0 ? 1 : -1)
  if (dc) col = cols.visible(col + dc, dc > 0 ? 1 : -1)
  const p = clampCell(r, col)
  if (extend) setSel(c, { ...s, ranges: [...s.ranges.slice(0, -1), expandMerges(rectOf(s.active, p), merges)], focus: p })
  else setSel(c, cellSel(p.r, p.c))
}

/** Grow a rectangle until it contains every merged cell it touches (as Excel does). */
export function expandMerges(r: { r0: number; c0: number; r1: number; c1: number }, merges: Area[]) {
  let out = { ...r }
  for (let changed = true; changed; ) {
    changed = false
    for (const m of merges) {
      if (m.r1 < out.r0 || m.r0 > out.r1 || m.c1 < out.c0 || m.c0 > out.c1) continue
      const next = { r0: Math.min(out.r0, m.r0), c0: Math.min(out.c0, m.c0), r1: Math.max(out.r1, m.r1), c1: Math.max(out.c1, m.c1) }
      if (next.r0 !== out.r0 || next.c0 !== out.c0 || next.r1 !== out.r1 || next.c1 !== out.c1) {
        out = next
        changed = true
      }
    }
  }
  return out
}

/** Ctrl+Arrow: jump to the edge of the data region. */
export async function jump(c: Ctx, dir: 'up' | 'down' | 'left' | 'right', extend = false) {
  const s = c.view.sel
  const from = extend ? s.focus : s.active
  const p = await c.book.api.navigate(c.view.sheet, from.r, from.c, dir)
  const v = getView(c.tab)
  if (extend) setView(c.tab, { sel: { ...v.sel, ranges: [...v.sel.ranges.slice(0, -1), rectOf(v.sel.active, p)], focus: p } })
  else setView(c.tab, { sel: cellSel(p.r, p.c) })
}

export async function jumpEnd(c: Ctx, which: 'home' | 'end', extend = false) {
  const p = await c.book.api.navigate(c.view.sheet, 0, 0, which)
  if (extend) setSel(c, { ...c.view.sel, ranges: [...c.view.sel.ranges.slice(0, -1), rectOf(c.view.sel.active, p)], focus: p })
  else setSel(c, cellSel(p.r, p.c))
}

export function selectRows(c: Ctx) {
  const s = c.view.sel
  const r = s.ranges[s.ranges.length - 1]
  setSel(c, { ...s, ranges: [...s.ranges.slice(0, -1), { r0: r.r0, r1: r.r1, c0: 0, c1: MAX_COLS - 1 }] })
}
export function selectCols(c: Ctx) {
  const s = c.view.sel
  const r = s.ranges[s.ranges.length - 1]
  setSel(c, { ...s, ranges: [...s.ranges.slice(0, -1), { r0: 0, r1: MAX_ROWS - 1, c0: r.c0, c1: r.c1 }] })
}
export function selectAll(c: Ctx) {
  setSel(c, { ranges: [{ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }], active: c.view.sel.active, focus: { r: MAX_ROWS - 1, c: MAX_COLS - 1 } })
}

/** Go to a reference or defined name typed in the Name Box or Go To dialog. Returns false if not understood. */
export async function goTo(c: Ctx, text: string): Promise<boolean> {
  const t = text.trim()
  if (!t) return false
  let sheet = c.view.sheet
  let ref = t
  const bang = t.lastIndexOf('!')
  if (bang > 0) {
    const name = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
    const s = c.book.meta.sheets.find((x) => x.name.toLowerCase() === name.toLowerCase())
    if (!s) return false
    sheet = s.id
    ref = t.slice(bang + 1)
  }
  let a = parseArea(ref, sheet)
  if (!a) {
    const n = c.book.meta.names.find((x) => x.name.toLowerCase() === t.toLowerCase())
    const tb = c.book.meta.tables.find((x) => x.name.toLowerCase() === t.toLowerCase())
    if (tb) a = tb.area
    else if (n) {
      const m = /^(?:'?(.+?)'?!)?\$?([A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(n.text)
      if (!m) return false
      const s = m[1] ? c.book.meta.sheets.find((x) => x.name.toLowerCase() === m[1].toLowerCase()) : null
      a = parseArea(m[2], s?.id ?? sheet)
    }
    if (!a) return false
  }
  if (a.sheet !== c.view.sheet) setView(c.tab, { sheet: a.sheet })
  setView(c.tab, { sel: { ranges: [{ r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }], active: { r: a.r0, c: a.c0 }, focus: { r: a.r1, c: a.c1 } } })
  return true
}

/** Text for the Name Box: a defined name if the selection is exactly one, else the active cell or range size. */
export function nameBoxText(c: Ctx): string {
  const s = c.view.sel
  const r = s.ranges[s.ranges.length - 1]
  for (const n of c.book.meta.names) {
    const m = /^(?:'?(.+?)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(n.text)
    const a = m && parseArea(m[2])
    const sh = m?.[1] ? c.book.meta.sheets.find((x) => x.name === m[1])?.id : c.view.sheet
    if (a && sh === c.view.sheet && a.r0 === r.r0 && a.c0 === r.c0 && a.r1 === r.r1 && a.c1 === r.c1) return n.name
  }
  return cellName(s.active.r, s.active.c)
}

export function selectionLabel(r: { r0: number; c0: number; r1: number; c1: number }) {
  if (r.r0 === 0 && r.r1 === MAX_ROWS - 1 && r.c0 === 0 && r.c1 === MAX_COLS - 1) return 'All cells'
  if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) return r.c0 === r.c1 ? `Column ${colName(r.c0)}` : `${r.c1 - r.c0 + 1}C`
  if (r.c0 === 0 && r.c1 === MAX_COLS - 1) return r.r0 === r.r1 ? `Row ${r.r0 + 1}` : `${r.r1 - r.r0 + 1}R`
  return `${r.r1 - r.r0 + 1}R × ${r.c1 - r.c0 + 1}C`
}

// ---------- editing ----------

export async function startEdit(c: Ctx, initial?: string, where: 'cell' | 'bar' = 'cell') {
  if (c.view.editing) return
  const { r, c: col } = c.view.sel.active
  const text = initial ?? (await c.book.api.editText(c.view.sheet, r, col))
  setView(c.tab, { editing: { text, where, enterMode: initial !== undefined, caret: text.length }, copy: getView(c.tab).copy })
}

export function updateEdit(c: Ctx, text: string, caret: number) {
  const e = c.view.editing
  if (!e) return
  setView(c.tab, { editing: { ...e, text, caret, point: undefined } })
}

export function cancelEdit(c: Ctx) {
  setView(c.tab, { editing: null })
}

/**
 * Commit what is being typed. Enter moves down, Tab right (Shift reverses). Ctrl+Enter fills the selection,
 * Ctrl+Shift+Enter enters a legacy array formula over it.
 */
export async function commitEdit(c: Ctx, dir: 'down' | 'up' | 'right' | 'left' | 'none' = 'down', mode: 'one' | 'all' | 'array' = 'one'): Promise<boolean> {
  const e = c.view.editing
  if (!e) return true
  const { r, c: col } = c.view.sel.active
  let text = e.text
  // Close unbalanced parentheses like Excel does.
  if (text.startsWith('=')) {
    let depth = 0
    let inStr = false
    for (const ch of text) {
      if (ch === '"') inStr = !inStr
      else if (!inStr && ch === '(') depth++
      else if (!inStr && ch === ')') depth--
    }
    if (depth > 0 && !inStr) text += ')'.repeat(depth)
  }
  const areas = areasOf(c.view)
  let res: { ok: boolean; error?: string; pos?: number; alert?: DvAlert }
  if (mode === 'all') res = await c.book.api.inputAll(areas, text)
  else if (mode === 'array') {
    const a = areas[areas.length - 1]
    res = await c.book.api.input(c.view.sheet, a.r0, a.c0, text, a)
  } else res = await c.book.api.input(c.view.sheet, r, col, text)
  // Data validation: a stop rule already undid the entry; warning and information rules ask whether to keep it.
  if (res.alert) {
    const choice = await showValidationAlert(res.alert)
    if (res.alert.style === 'stop') {
      if (choice === 'cancel') setView(c.tab, { editing: null })
      return false
    }
    if (choice === 'no' || choice === 'cancel') {
      await c.book.api.undo()
      if (choice === 'cancel') setView(c.tab, { editing: null })
      return false
    }
  }
  if (!res.ok) {
    // Formula errors carry the position of the problem; other refusals (a data table's cells) don't.
    if (res.pos === undefined && res.error) {
      // The entry is discarded, as in Excel (there is nothing to fix in it).
      setView(c.tab, { editing: null })
      set({ dialog: { kind: 'message', title: 'Can’t change that', body: res.error } })
    }
    else set({ dialog: { kind: 'message', title: 'There’s a problem with this formula', body: `${res.error ?? 'The formula could not be read.'}\n\nTo type text that starts with “=”, begin it with an apostrophe: '=` } })
    return false
  }
  markDirty(c.tab)
  setView(c.tab, { editing: null })
  if (getView(c.tab).circles?.length) void circleInvalid(c)
  const v = getView(c.tab)
  if (dir !== 'none') advance({ ...c, view: v }, dir)
  return true
}

/** Enter/Tab movement: inside a multi-cell selection the active cell cycles through it. */
export function advance(c: Ctx, dir: 'down' | 'up' | 'right' | 'left') {
  const s = c.view.sel
  const last = s.ranges[s.ranges.length - 1]
  const multi = s.ranges.length > 1 || last.r0 !== last.r1 || last.c0 !== last.c1
  if (!multi) {
    const d = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] }[dir]
    return move(c, d[0], d[1])
  }
  let { r, c: col } = s.active
  if (dir === 'down' || dir === 'up') {
    r += dir === 'down' ? 1 : -1
    if (r > last.r1) {
      r = last.r0
      col = col + 1 > last.c1 ? last.c0 : col + 1
    } else if (r < last.r0) {
      r = last.r1
      col = col - 1 < last.c0 ? last.c1 : col - 1
    }
  } else {
    col += dir === 'right' ? 1 : -1
    if (col > last.c1) {
      col = last.c0
      r = r + 1 > last.r1 ? last.r0 : r + 1
    } else if (col < last.c0) {
      col = last.c1
      r = r - 1 < last.r0 ? last.r1 : r - 1
    }
  }
  setSel(c, { ...s, active: { r, c: col } })
}

/** Delete key: clear contents of the selection. */
export function clearContents(c: Ctx) {
  return edit(c, (b) => b.api.clear(areasOf(c.view), 'contents'))
}

// ---------- formatting ----------

export function formatSel(c: Ctx, patch: Partial<Style>) {
  return edit(c, (b) => b.api.formatCells(areasOf(c.view), patch))
}

export function toggleSel(c: Ctx, key: 'bold' | 'italic' | 'strike' | 'wrap') {
  return edit(c, (b) => b.api.toggle(areasOf(c.view), key, { sheet: c.view.sheet, ...c.view.sel.active }))
}

export async function activeStyle(c: Ctx): Promise<Style> {
  return c.book.api.styleAt(c.view.sheet, c.view.sel.active.r, c.view.sel.active.c)
}

// ---------- clipboard ----------

const CLIP_MIME = 'web application/x-simple-cell'

export async function copy(c: Ctx, cut = false) {
  const s = c.view.sel
  if (s.ranges.length > 1) return toast('That command can’t be used on multiple selections.', 'error')
  const r = s.ranges[0]
  const area: Area = { sheet: c.view.sheet, ...clipRect(c, r) }
  const { tsv, html } = await c.book.api.copyText(area)
  const items: Record<string, Blob> = {
    'text/plain': new Blob([tsv], { type: 'text/plain' }),
    'text/html': new Blob([html], { type: 'text/html' })
  }
  try {
    items[CLIP_MIME] = new Blob([JSON.stringify({ book: c.book.id, area, cut, tsv })], { type: CLIP_MIME })
    await navigator.clipboard.write([new ClipboardItem(items)])
  } catch {
    await navigator.clipboard.writeText(tsv)
  }
  internalClip = { book: c.book.id, area, cut, tsv }
  setView(c.tab, { copy: { ...area, cut } })
}

let internalClip: { book: string; area: Area; cut: boolean; tsv: string } | null = null

/** Whole rows/columns are clipped to the used range for copying. */
function clipRect(c: Ctx, r: { r0: number; c0: number; r1: number; c1: number }) {
  const u = c.book.sheetMeta(c.view.sheet)!.used
  return { r0: r.r0, c0: r.c0, r1: r.r1 === MAX_ROWS - 1 ? Math.max(r.r0, u.r) : r.r1, c1: r.c1 === MAX_COLS - 1 ? Math.max(r.c0, u.c) : r.c1 }
}

export async function paste(c: Ctx, what: 'all' | 'values' | 'formats' | 'formulas' = 'all') {
  const s = c.view.sel
  const dst = s.ranges[s.ranges.length - 1]
  let text = ''
  try {
    text = await navigator.clipboard.readText()
  } catch {
    /* no permission */
  }
  // Copied inside Simple Cell (same text still on the clipboard): paste cells with formulas and formats.
  if (internalClip && internalClip.tsv === text) {
    const src = internalClip.area
    const srcBook = books.get(get().tabs.find((t) => books.get(t.id)?.id === internalClip!.book)?.id ?? '') ?? null
    if (srcBook === c.book) {
      const h = src.r1 - src.r0 + 1
      const w = src.c1 - src.c0 + 1
      // Paste into a single cell: the size of the copied range. Larger selections that are multiples tile it.
      const single = dst.r0 === dst.r1 && dst.c0 === dst.c1
      const target: Area = { sheet: c.view.sheet, r0: dst.r0, c0: dst.c0, r1: single ? dst.r0 + h - 1 : dst.r1, c1: single ? dst.c0 + w - 1 : dst.c1 }
      const ok = await edit(c, (b) => b.api.pasteCells(src, target, what, internalClip!.cut))
      if (ok) {
        setSel(c, { ranges: [{ r0: target.r0, c0: target.c0, r1: Math.max(target.r1, target.r0 + h - 1), c1: Math.max(target.c1, target.c0 + w - 1) }], active: { r: target.r0, c: target.c0 }, focus: { r: target.r1, c: target.c1 } })
        if (internalClip.cut) {
          internalClip = null
          setView(c.tab, { copy: null })
        }
      }
      return
    }
  }
  if (!text) return
  const rows = parseClipboardText(text)
  const ok = await edit(c, (b) => b.api.pasteText(c.view.sheet, dst.r0, dst.c0, rows))
  if (ok) {
    const h = rows.length
    const w = Math.max(...rows.map((x) => x.length))
    setSel(c, { ranges: [{ r0: dst.r0, c0: dst.c0, r1: dst.r0 + h - 1, c1: dst.c0 + w - 1 }], active: { r: dst.r0, c: dst.c0 }, focus: { r: dst.r0 + h - 1, c: dst.c0 + w - 1 } })
  }
}

/** Tab-separated text from the clipboard (quoted fields may contain tabs and line breaks). */
export function parseClipboardText(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  const t = text.replace(/\r\n/g, '\n').replace(/\n$/, '')
  for (let i = 0; i < t.length; i++) {
    const ch = t[i]
    if (quoted) {
      if (ch === '"' && t[i + 1] === '"') {
        field += '"'
        i++
      } else if (ch === '"') quoted = false
      else field += ch
    } else if (ch === '"' && field === '') quoted = true
    else if (ch === '\t') {
      row.push(field)
      field = ''
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else field += ch
  }
  row.push(field)
  rows.push(row)
  return rows
}

export function clearCopy(c: Ctx) {
  if (c.view.copy) setView(c.tab, { copy: null })
}

// ---------- structure ----------

/** Insert rows/columns for the selection (whole rows/columns selected: that many; otherwise asks via the menu). */
export function insertAt(c: Ctx, axis: 'row' | 'col') {
  const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  return axis === 'row' ? edit(c, (b) => b.api.insert(c.view.sheet, 'row', r.r0, r.r1 - r.r0 + 1)) : edit(c, (b) => b.api.insert(c.view.sheet, 'col', r.c0, r.c1 - r.c0 + 1))
}

export function deleteAt(c: Ctx, axis: 'row' | 'col') {
  const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  return axis === 'row' ? edit(c, (b) => b.api.remove(c.view.sheet, 'row', r.r0, r.r1 - r.r0 + 1)) : edit(c, (b) => b.api.remove(c.view.sheet, 'col', r.c0, r.c1 - r.c0 + 1))
}

export async function undo(c: Ctx) {
  if (!c.book.meta.canUndo) return
  await c.book.api.undo()
  markDirty(c.tab)
}
export async function redo(c: Ctx) {
  if (!c.book.meta.canRedo) return
  await c.book.api.redo()
  markDirty(c.tab)
}

export function activeBookOrNull() {
  return activeBook()
}

export { areaName, areasOf }

// ---------- data validation ----------

/** Circle Invalid Data: mark cells whose values break their rules (refreshed after each entry). */
export async function circleInvalid(c: Ctx) {
  const cells = await c.book.api.invalidCells(c.view.sheet)
  setView(c.tab, { circles: cells })
  return cells.length
}

export function clearCircles(c: Ctx) {
  setView(c.tab, { circles: [] })
}

// ---------- paste special, format painter ----------

/** The cells copied inside Simple Cell for this workbook (Paste Special needs them). */
export function copiedArea(c: Ctx): Area | null {
  return internalClip && internalClip.book === c.book.id && !internalClip.cut ? internalClip.area : null
}

export async function pasteSpecial(c: Ctx, o: import('@/core/doc').PasteSpecialOptions) {
  const src = copiedArea(c)
  if (!src) return toast('Copy cells in Simple Cell first; Paste Special works with copied cells.', 'info')
  const s = c.view.sel
  const dst = s.ranges[s.ranges.length - 1]
  const h = o.transpose ? src.c1 - src.c0 + 1 : src.r1 - src.r0 + 1
  const w = o.transpose ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1
  const single = dst.r0 === dst.r1 && dst.c0 === dst.c1
  const target: Area = { sheet: c.view.sheet, r0: dst.r0, c0: dst.c0, r1: single ? dst.r0 + h - 1 : dst.r1, c1: single ? dst.c0 + w - 1 : dst.c1 }
  if (await edit(c, (b) => b.api.pasteSpecial(src, target, o))) setView(c.tab, (v) => ({ sel: { ...v.sel, ranges: [target], focus: { r: target.r1, c: target.c1 } } }))
}

/** Format Painter: remember the selection's formats; the next selection gets them (sticky: until Esc). */
export function startPainter(c: Ctx, sticky: boolean) {
  const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  setView(c.tab, { painter: { area: { sheet: c.view.sheet, ...r }, sticky } })
}

export async function applyPainter(c: Ctx) {
  const p = c.view.painter
  if (!p) return
  const dst = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  const h = p.area.r1 - p.area.r0 + 1
  const w = p.area.c1 - p.area.c0 + 1
  const single = dst.r0 === dst.r1 && dst.c0 === dst.c1
  const target: Area = { sheet: c.view.sheet, r0: dst.r0, c0: dst.c0, r1: single ? dst.r0 + h - 1 : dst.r1, c1: single ? dst.c0 + w - 1 : dst.c1 }
  if (!p.sticky) setView(c.tab, { painter: undefined })
  await edit(c, (b) => b.api.pasteSpecial(p.area, target, { what: 'formats' }))
}

/** Go To Special: select the matching cells (within the selection, or the used range from a single cell). */
export async function goToSpecial(c: Ctx, kind: import('@/core/doc').GoToKind, types = 'ntle') {
  const r = await c.book.api.goToSpecial(c.view.sheet, areasOf(c.view), kind, types)
  if (!r.ok) return toast(r.error, 'info')
  const ranges = r.areas.map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 }))
  const first = ranges[0]
  setView(c.tab, { sel: { ranges, active: { r: first.r0, c: first.c0 }, focus: { r: first.r1, c: first.c1 } } })
}

export function markDirtyTab(tab: string) {
  markDirty(tab)
}
