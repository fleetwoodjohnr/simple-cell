// Keyboard behaviour of the grid and the in-cell editor (Excel's): navigation, entering, point mode, F4.
import { MAX_COLS, MAX_ROWS } from '@/core/engine/values'
import { cellName, colIndex, colName, parseCell } from '@/core/refs'
import * as A from '@/renderer/app/actions'
import { getView, rectOf, setView } from '@/renderer/state/view'
import { visibleRows, type Layout } from './layout'
import type { Book } from '@/renderer/state/client'
import { isMac } from '@/renderer/app/commands'

const REF_COLORS = ['#2f6fde', '#d1414b', '#7a46c9', '#1f8a5b', '#c9741a', '#c23f8c', '#2697a6']

/** Characters after which arrow keys and clicks insert a reference into a formula. */
const POINT_AFTER = /[=(,+\-*/^&<>:;{\s%]$/

const REF_RE = /(?:('(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(])/g

/** References in a formula being typed, with the colours Excel uses for them. */
export function refHighlights(text: string, sheet: number, book: Book) {
  if (!text.startsWith('=')) return []
  const out: { area: { r0: number; c0: number; r1: number; c1: number }; color: string }[] = []
  // Skip text inside quotes
  const stripped = text.replace(/"(?:[^"]|"")*"/g, (m) => ' '.repeat(m.length))
  let i = 0
  for (const m of stripped.matchAll(REF_RE)) {
    const sheetName = m[1]?.replace(/^'|'$/g, '').replace(/''/g, "'")
    if (sheetName) {
      const s = book.meta.sheets.find((x) => x.name.toLowerCase() === sheetName.toLowerCase())
      if (!s || s.id !== sheet) {
        i++
        continue
      }
    }
    const [a, b] = m[2].replace(/\$/g, '').split(':')
    let area: { r0: number; c0: number; r1: number; c1: number } | null = null
    if (/^[A-Za-z]+$/.test(a)) area = { r0: 0, r1: MAX_ROWS - 1, c0: Math.min(colIndex(a), colIndex(b)), c1: Math.max(colIndex(a), colIndex(b)) }
    else if (/^\d+$/.test(a)) area = { c0: 0, c1: MAX_COLS - 1, r0: Math.min(+a, +b) - 1, r1: Math.max(+a, +b) - 1 }
    else {
      const p = parseCell(a)
      const q = b ? parseCell(b) : p
      if (p && q) area = rectOf(p, q)
    }
    if (area) out.push({ area, color: REF_COLORS[i % REF_COLORS.length] })
    i++
  }
  return out
}

function refText(a: { r: number; c: number }, b: { r: number; c: number }) {
  const r = rectOf(a, b)
  return r.r0 === r.r1 && r.c0 === r.c1 ? cellName(r.r0, r.c0) : `${cellName(r.r0, r.c0)}:${cellName(r.r1, r.c1)}`
}

/**
 * Point mode: insert (or replace) a reference at the caret while a formula is being typed.
 * `focus` null = start a new reference at `anchor`. Returns false when the caret is not at a point where a reference
 * can go (then the click commits the edit instead).
 */
export function insertPointRef(c: A.Ctx, anchor: { r: number; c: number }, focus: { r: number; c: number } | null): boolean {
  const e = c.view.editing
  if (!e || !e.text.startsWith('=')) return false
  if (e.point && focus) {
    const ref = refText(e.point.anchor, focus)
    const text = e.text.slice(0, e.point.start) + ref + e.text.slice(e.point.end)
    setView(c.tab, { editing: { ...e, text, caret: e.point.start + ref.length, point: { ...e.point, end: e.point.start + ref.length, focus } } })
    return true
  }
  const before = e.text.slice(0, e.caret)
  if (!POINT_AFTER.test(before)) return false
  const ref = refText(anchor, focus ?? anchor)
  const text = before + ref + e.text.slice(e.caret)
  setView(c.tab, { editing: { ...e, text, caret: e.caret + ref.length, point: { start: e.caret, end: e.caret + ref.length, anchor, focus: focus ?? anchor } } })
  return true
}

/** F4: cycle the reference at the caret through A1 → $A$1 → A$1 → $A1. */
export function cycleRef(text: string, caret: number): { text: string; caret: number } | null {
  const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g
  for (const m of text.matchAll(re)) {
    const start = m.index!
    const end = start + m[0].length
    if (caret < start || caret > end) continue
    if (/[\w"]/.test(text[start - 1] ?? '') || /[\w(]/.test(text[end] ?? '')) continue
    const [, ca, col, ra, row] = m
    const state = (ca ? 2 : 0) + (ra ? 1 : 0) // 0 A1, 3 $A$1, 1 A$1, 2 $A1
    const next = { 0: 3, 3: 1, 1: 2, 2: 0 }[state] as number
    const rep = `${next & 2 ? '$' : ''}${col}${next & 1 ? '$' : ''}${row}`
    return { text: text.slice(0, start) + rep + text.slice(end), caret: start + rep.length }
  }
  return null
}

const ARROWS: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }

export function handleGridKey(e: React.KeyboardEvent<HTMLTextAreaElement>, c: A.Ctx, L: Layout | null) {
  if (e.nativeEvent.isComposing) return
  const mod = isMac ? e.metaKey : e.ctrlKey
  const editing = c.view.editing
  const key = e.key
  if (editing) {
    if (key === 'Escape') {
      e.preventDefault()
      A.cancelEdit(c)
      return
    }
    if (key === 'Enter' && (e.altKey || (isMac && e.ctrlKey && !e.metaKey))) {
      // New line inside the cell
      e.preventDefault()
      const t = e.currentTarget
      const pos = t.selectionStart
      const text = editing.text.slice(0, pos) + '\n' + editing.text.slice(t.selectionEnd)
      // The caret follows the state (useCaret in Grid.tsx).
      setView(c.tab, { editing: { ...editing, text, caret: pos + 1 } })
      return
    }
    if (key === 'Enter') {
      e.preventDefault()
      const mode = mod && e.shiftKey ? 'array' : mod ? 'all' : 'one'
      void A.commitEdit(c, mode === 'one' ? (e.shiftKey ? 'up' : 'down') : 'none', mode)
      return
    }
    if (key === 'Tab') {
      e.preventDefault()
      void A.commitEdit(c, e.shiftKey ? 'left' : 'right')
      return
    }
    if (key === 'F2') {
      e.preventDefault()
      setView(c.tab, { editing: { ...editing, enterMode: !editing.enterMode, point: undefined } })
      return
    }
    if (key === 'F4' && editing.text.startsWith('=')) {
      e.preventDefault()
      const r = cycleRef(editing.text, e.currentTarget.selectionStart)
      if (r) {
        setView(c.tab, { editing: { ...editing, text: r.text, caret: r.caret } })
      }
      return
    }
    const arrow = ARROWS[key]
    if (arrow && editing.enterMode && !e.altKey && !mod) {
      // Point mode in formulas; otherwise the arrow commits and moves (Excel's "Enter" mode).
      if (editing.text.startsWith('=') && (editing.point || POINT_AFTER.test(editing.text.slice(0, editing.caret)))) {
        e.preventDefault()
        const base = editing.point ? editing.point.focus : c.view.sel.active
        const next = { r: clamp(base.r + arrow[0], MAX_ROWS), c: clamp(base.c + arrow[1], MAX_COLS) }
        if (editing.point && e.shiftKey) insertPointRef(c, editing.point.anchor, next)
        else if (editing.point) {
          const ref = cellName(next.r, next.c)
          const text = editing.text.slice(0, editing.point.start) + ref + editing.text.slice(editing.point.end)
          setView(c.tab, { editing: { ...editing, text, caret: editing.point.start + ref.length, point: { start: editing.point.start, end: editing.point.start + ref.length, anchor: next, focus: next } } })
        } else insertPointRef(c, next, null)
        return
      }
      e.preventDefault()
      void A.commitEdit(c, key === 'ArrowUp' ? 'up' : key === 'ArrowDown' ? 'down' : key === 'ArrowLeft' ? 'left' : 'right')
      return
    }
    if (editing.point && !['Shift', 'Control', 'Meta', 'Alt'].includes(key)) setView(c.tab, { editing: { ...editing, point: undefined } })
    return // typing goes into the textarea
  }

  // ---------- not editing ----------
  const arrow = ARROWS[key]
  if (arrow) {
    e.preventDefault()
    // Alt+Down opens the cell's validation dropdown.
    if (e.altKey && key === 'ArrowDown' && !mod) return setView(c.tab, { listOpen: true })
    if (e.altKey) return
    if (mod) void A.jump(c, key === 'ArrowUp' ? 'up' : key === 'ArrowDown' ? 'down' : key === 'ArrowLeft' ? 'left' : 'right', e.shiftKey)
    else A.move(c, arrow[0], arrow[1], e.shiftKey)
    return
  }
  switch (key) {
    case 'Enter':
      e.preventDefault()
      if (mod) return
      A.advance(c, e.shiftKey ? 'up' : 'down')
      return
    case 'Tab':
      e.preventDefault()
      A.advance(c, e.shiftKey ? 'left' : 'right')
      return
    case 'F2':
      e.preventDefault()
      void A.startEdit(c)
      return
    case 'Delete':
      e.preventDefault()
      void A.clearContents(c)
      return
    case 'Backspace':
      e.preventDefault()
      if (isMac && !mod) void A.clearContents(c)
      else {
        void A.clearContents(c)
        setView(c.tab, { editing: { text: '', where: 'cell', enterMode: true, caret: 0 } })
      }
      return
    case 'Escape':
      A.clearCopy(c)
      if (c.view.painter) setView(c.tab, { painter: undefined })
      return
    case 'Home':
      e.preventDefault()
      if (mod) void A.jumpEnd(c, 'home', e.shiftKey)
      else if (e.shiftKey) A.setSel(c, { ...c.view.sel, ranges: [...c.view.sel.ranges.slice(0, -1), rectOf(c.view.sel.active, { r: c.view.sel.focus.r, c: 0 })], focus: { r: c.view.sel.focus.r, c: 0 } })
      else A.setSel(c, { ranges: [{ r0: c.view.sel.active.r, c0: 0, r1: c.view.sel.active.r, c1: 0 }], active: { r: c.view.sel.active.r, c: 0 }, focus: { r: c.view.sel.active.r, c: 0 } })
      return
    case 'End':
      if (mod) {
        e.preventDefault()
        void A.jumpEnd(c, 'end', e.shiftKey)
      }
      return
    case 'PageDown':
    case 'PageUp': {
      e.preventDefault()
      const page = L ? Math.max(1, visibleRows(L)[1] - visibleRows(L)[0]) : 20
      const d = key === 'PageDown' ? page : -page
      if (e.altKey) A.move(c, 0, d > 0 ? 10 : -10, e.shiftKey)
      else if (mod) return
      else A.move(c, d, 0, e.shiftKey)
      return
    }
    case ' ':
      if (e.shiftKey && !mod) {
        e.preventDefault()
        A.selectRows(c)
        return
      }
      if (mod && !e.shiftKey) {
        e.preventDefault()
        A.selectCols(c)
        return
      }
      if (mod && e.shiftKey) {
        e.preventDefault()
        A.selectAll(c)
        return
      }
  }
}

const clamp = (v: number, max: number) => Math.max(0, Math.min(max - 1, v))

export { colName }
