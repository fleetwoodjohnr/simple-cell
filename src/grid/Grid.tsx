// The spreadsheet grid: a canvas painted from the cell cache, native scrollbars, and one textarea that receives the
// keyboard (invisible over the active cell until you type, then it becomes the in-cell editor).
import { ValidationOverlay } from '@/renderer/app/Validation'
import { FilterOverlay } from '@/renderer/app/Tables'
import { useCallback, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react'
import { useShallow } from 'zustand/react/shallow'
import { MAX_COLS, MAX_ROWS } from '@/core/engine/values'
import { books, get, set, useApp } from '@/renderer/state/store'
import { getView, rectOf, setView, useViews, cellSel, type Sel } from '@/renderer/state/view'
import * as A from '@/renderer/app/actions'
import { handleGridKey, refHighlights, insertPointRef } from './keys'
import { colX, hit, makeLayout, rowY, scrollToShow, visibleCols, visibleRows, axesFor, type Layout } from './layout'
import { clearMeasureCache, fontOf, paint, type Theme } from './paint'
import { GridMenu } from './GridMenu'

function readTheme(dark: boolean): Theme {
  const cs = getComputedStyle(document.documentElement)
  const v = (n: string) => cs.getPropertyValue(n).trim()
  return { bg: v('--grid-bg') || '#fff', line: v('--grid-line'), header: v('--grid-header'), headerText: v('--grid-header-text'), headerActive: v('--grid-header-active'), accent: v('--accent'), selFill: v('--grid-sel-fill'), text: dark ? '#e7e7ea' : '#000000', dark }
}

type Drag =
  | { kind: 'select'; add: boolean; start: { r: number; c: number } }
  | { kind: 'cols' | 'rows'; start: number }
  | { kind: 'resize-col' | 'resize-row'; index: number; from: number; size: number; indexes: number[] }
  | { kind: 'fill'; src: { r0: number; c0: number; r1: number; c1: number } }
  | { kind: 'point'; anchor: { r: number; c: number } }

export function Grid({ tab, hidden }: { tab: string; hidden: boolean }) {
  const book = books.get(tab)!
  const view = useViews((s) => s.views[tab])
  const { dark, showFormulas, showHeadings, defaultFont, defaultFontSize } = useApp(useShallow((s) => ({ dark: s.dark, showFormulas: s.showFormulas, showHeadings: s.showHeadings, defaultFont: s.settings.defaultFont, defaultFontSize: s.settings.defaultFontSize })))
  const version = useSyncExternalStore(
    useCallback((fn) => book.subscribe(fn), [book]),
    () => book.version
  )
  const hostRef = useRef<HTMLDivElement>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const [size, setSize] = useState({ w: 0, h: 0 })
  const [dash, setDash] = useState(0)
  const [guide, setGuide] = useState<{ x?: number; y?: number } | null>(null)
  const [fillPreview, setFillPreview] = useState<{ r0: number; c0: number; r1: number; c1: number } | null>(null)
  const drag = useRef<Drag | null>(null)
  const formulas = useRef(new Map<string, string>())

  const meta = book.sheetMeta(view?.sheet ?? -1)
  const zoom = (view?.zoom ?? 1) * (meta?.zoom ?? 1)
  const scroll = view?.scroll[view.sheet] ?? { x: 0, y: 0 }
  const L: Layout | null = meta && size.w ? makeLayout(meta, size.w, size.h, zoom, scroll.x, scroll.y, showHeadings) : null

  // Size of the scroll area: the used range plus room to keep going ("infinite" scrolling up to the sheet limits).
  let spacer = { w: 0, h: 0 }
  if (L && meta && view) {
    const [, r1] = visibleRows(L)
    const [, c1] = visibleCols(L)
    const rowsExtent = Math.min(MAX_ROWS, Math.max(meta.used.r + 1, view.sel.active.r + 1, r1 + 1) + 60)
    const colsExtent = Math.min(MAX_COLS, Math.max(meta.used.c + 1, view.sel.active.c + 1, c1 + 1) + 12)
    spacer = { w: L.headerW + L.cols.offset(colsExtent) * zoom, h: L.headerH + L.rows.offset(rowsExtent) * zoom }
  }

  // Canvas text doesn't trigger webfont loading: load every face the workbook uses, then repaint.
  const [fontsLoaded, setFontsLoaded] = useState(0)
  useEffect(() => {
    const faces = new Set<string>()
    for (const st of book.meta.styles) for (const b of [false, true]) faces.add(fontOf({ ...st, bold: st.bold || b }, 1, defaultFont, defaultFontSize))
    let cancelled = false
    void Promise.all([...faces].map((f) => document.fonts.load(f).catch(() => []))).then(() => {
      if (cancelled) return
      clearMeasureCache()
      setFontsLoaded((n) => n + 1)
    })
    return () => {
      cancelled = true
    }
  }, [book.meta.styles.length, defaultFont, defaultFontSize])
  void fontsLoaded

  // Track the element size.
  useLayoutEffect(() => {
    const el = scrollRef.current
    if (!el) return
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }))
    ro.observe(el)
    setSize({ w: el.clientWidth, h: el.clientHeight })
    return () => ro.disconnect()
  }, [])

  // Fetch the visible cells (scrolling pane and frozen panes).
  useEffect(() => {
    if (!L || !view || hidden) return
    const [r0, r1] = visibleRows(L)
    const [c0, c1] = visibleCols(L)
    const wins = [{ r0, r1, c0, c1 }]
    if (L.fr) wins.push({ r0: 0, r1: L.fr - 1, c0, c1 })
    if (L.fc) wins.push({ r0, r1, c0: 0, c1: L.fc - 1 })
    if (L.fr && L.fc) wins.push({ r0: 0, r1: L.fr - 1, c0: 0, c1: L.fc - 1 })
    void book.ensure(view.sheet, wins)
  })

  // Formula texts for "Show formulas" (fetched lazily for visible formula cells).
  useEffect(() => {
    if (!showFormulas || !L || !view) return
    const [r0, r1] = visibleRows(L)
    const [c0, c1] = visibleCols(L)
    let missing = false
    for (let r = r0; r <= r1; r++)
      for (let c = c0; c <= c1; c++) {
        const v = book.cell(view.sheet, r, c)
        const k = `${view.sheet}:${r}:${c}`
        if (v?.f && !formulas.current.has(k)) {
          missing = true
          void book.api.editText(view.sheet, r, c).then((t) => {
            formulas.current.set(k, t)
            setDash((d) => d + 0.0001)
          })
        }
      }
    void missing
  }, [showFormulas, version, L?.sx, L?.sy, view?.sheet])
  useEffect(() => formulas.current.clear(), [version])

  // Marching ants
  useEffect(() => {
    if (!view?.copy) return
    const t = setInterval(() => setDash((d) => (Math.floor(d) + 1) % 9), 90)
    return () => clearInterval(t)
  }, [!!view?.copy])

  // Paint
  useLayoutEffect(() => {
    const cv = canvasRef.current
    if (!cv || !L || !meta || !view || hidden) return
    const dpr = window.devicePixelRatio || 1
    if (cv.width !== Math.round(L.width * dpr) || cv.height !== Math.round(L.height * dpr)) {
      cv.width = Math.round(L.width * dpr)
      cv.height = Math.round(L.height * dpr)
      cv.style.width = `${L.width}px`
      cv.style.height = `${L.height}px`
    }
    const ctx = cv.getContext('2d')!
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    const sheet = view.sheet
    paint(ctx, {
      L,
      meta,
      styles: book.meta.styles,
      cell: (r, c) => book.cell(sheet, r, c),
      sel: view.sel,
      editingCell: view.editing ? view.sel.active : null,
      copy: view.copy,
      dashOffset: Math.floor(dash),
      theme: readTheme(dark),
      defaultFont,
      defaultSize: defaultFontSize,
      showFormulas,
      formulaText: (r, c) => formulas.current.get(`${sheet}:${r}:${c}`),
      refHighlights: view.editing ? refHighlights(view.editing.text, sheet, book) : [],
      fillPreview
    })
  })

  // Keep the scroll position in step with the view state (e.g. after Go To).
  useLayoutEffect(() => {
    const el = scrollRef.current
    if (!el || !L) return
    const x = Math.round(scroll.x * zoom)
    const y = Math.round(scroll.y * zoom)
    if (Math.abs(el.scrollLeft - x) > 1) el.scrollLeft = x
    if (Math.abs(el.scrollTop - y) > 1) el.scrollTop = y
  }, [scroll.x, scroll.y, zoom, spacer.w, spacer.h])

  // Scroll the active/focus cell into view when the selection moves.
  const lastSel = useRef<Sel | null>(null)
  useEffect(() => {
    if (!L || !view || lastSel.current === view.sel) return
    lastSel.current = view.sel
    const p = view.sel.focus
    const next = scrollToShow(L, p.r, p.c)
    if (next) setView(tab, (v) => ({ scroll: { ...v.scroll, [v.sheet]: { x: next.sx, y: next.sy } } }))
  })

  // Focus the keyboard sink when this tab becomes visible.
  useEffect(() => {
    if (!hidden && !view?.editing) inputRef.current?.focus({ preventScroll: true })
  }, [hidden, view?.sheet])

  if (!view || !meta) return null

  const onScroll = () => {
    const el = scrollRef.current!
    const x = el.scrollLeft / zoom
    const y = el.scrollTop / zoom
    if (x !== scroll.x || y !== scroll.y) setView(tab, (v) => ({ scroll: { ...v.scroll, [v.sheet]: { x, y } } }))
  }

  const pos = (e: { clientX: number; clientY: number }) => {
    const rect = scrollRef.current!.getBoundingClientRect()
    return { x: e.clientX - rect.left, y: e.clientY - rect.top }
  }

  const onPointerDown = (e: React.PointerEvent) => {
    if (!L) return
    const { x, y } = pos(e)
    if (x > L.width || y > L.height) return // scrollbars
    const c = A.ctx()
    if (!c || c.tab !== tab) return
    const h = hit(L, x, y)
    const v = getView(tab)
    if (e.button === 2) {
      // Right-click keeps a selection that contains the cell.
      const inside = v.sel.ranges.some((r) => h.r >= r.r0 && h.r <= r.r1 && h.c >= r.c0 && h.c <= r.c1)
      if (!inside && h.zone === 'cell') setView(tab, { sel: cellSel(h.r, h.c) })
      if (!inside && h.zone === 'colhdr') setView(tab, { sel: { ranges: [{ r0: 0, r1: MAX_ROWS - 1, c0: h.c, c1: h.c }], active: { r: 0, c: h.c }, focus: { r: MAX_ROWS - 1, c: h.c } } })
      if (!inside && h.zone === 'rowhdr') setView(tab, { sel: { ranges: [{ r0: h.r, r1: h.r, c0: 0, c1: MAX_COLS - 1 }], active: { r: h.r, c: 0 }, focus: { r: h.r, c: MAX_COLS - 1 } } })
      return
    }
    if (e.button !== 0) return
    e.preventDefault()
    scrollRef.current!.setPointerCapture(e.pointerId)
    // Formula point mode: clicking a cell while typing a formula inserts its reference.
    if (v.editing && h.zone === 'cell' && insertPointRef({ ...c, view: v }, { r: h.r, c: h.c }, null)) {
      drag.current = { kind: 'point', anchor: { r: h.r, c: h.c } }
      return
    }
    if (v.editing) {
      void A.commitEdit({ ...c, view: v }, 'none').then((ok) => ok && inputRef.current?.focus())
    }
    // Column/row resize handles
    if (h.zone === 'colhdr') {
      const edge = nearEdge(L, x, 'col')
      if (edge !== null) {
        const sel = v.sel.ranges.filter((r) => r.r0 === 0 && r.r1 === MAX_ROWS - 1 && edge >= r.c0 && edge <= r.c1)
        const indexes = sel.length ? sel.flatMap((r) => range(r.c0, r.c1)) : [edge]
        drag.current = { kind: 'resize-col', index: edge, from: x, size: L.cols.size(edge), indexes }
        return
      }
    }
    if (h.zone === 'rowhdr') {
      const edge = nearEdge(L, y, 'row')
      if (edge !== null) {
        const sel = v.sel.ranges.filter((r) => r.c0 === 0 && r.c1 === MAX_COLS - 1 && edge >= r.r0 && edge <= r.r1)
        const indexes = sel.length ? sel.flatMap((r) => range(r.r0, r.r1)) : [edge]
        drag.current = { kind: 'resize-row', index: edge, from: y, size: L.rows.size(edge), indexes }
        return
      }
    }
    if (h.zone === 'corner') {
      A.selectAll({ ...c, view: v })
      return
    }
    // Fill handle
    const last = v.sel.ranges[v.sel.ranges.length - 1]
    if (h.zone === 'cell' && v.sel.ranges.length === 1) {
      const hx = colX(L, last.c1 + 1)
      const hy = rowY(L, last.r1 + 1)
      if (Math.abs(x - hx) <= 5 && Math.abs(y - hy) <= 5) {
        drag.current = { kind: 'fill', src: last }
        return
      }
    }
    const add = e.ctrlKey || e.metaKey
    if (h.zone === 'colhdr') {
      const start = e.shiftKey ? v.sel.active.c : h.c
      const r = { r0: 0, r1: MAX_ROWS - 1, c0: Math.min(start, h.c), c1: Math.max(start, h.c) }
      setView(tab, { sel: { ranges: add ? [...v.sel.ranges, r] : [r], active: { r: L.rows.offset(0) >= 0 ? Math.max(0, visibleRows(L)[0] < L.fr ? 0 : visibleRows(L)[0]) : 0, c: start }, focus: { r: MAX_ROWS - 1, c: h.c } } })
      drag.current = { kind: 'cols', start }
      return
    }
    if (h.zone === 'rowhdr') {
      const start = e.shiftKey ? v.sel.active.r : h.r
      const r = { r0: Math.min(start, h.r), r1: Math.max(start, h.r), c0: 0, c1: MAX_COLS - 1 }
      setView(tab, { sel: { ranges: add ? [...v.sel.ranges, r] : [r], active: { r: start, c: visibleCols(L)[0] < L.fc ? 0 : visibleCols(L)[0] }, focus: { r: h.r, c: MAX_COLS - 1 } } })
      drag.current = { kind: 'rows', start }
      return
    }
    const p = { r: h.r, c: h.c }
    if (e.shiftKey) {
      const s = v.sel
      setView(tab, { sel: { ...s, ranges: [...s.ranges.slice(0, -1), A.expandMerges(rectOf(s.active, p), meta.merges)], focus: p } })
      drag.current = { kind: 'select', add: false, start: s.active }
    } else {
      const m = meta.merges.find((mm) => p.r >= mm.r0 && p.r <= mm.r1 && p.c >= mm.c0 && p.c <= mm.c1)
      const range0 = m ? { r0: m.r0, c0: m.c0, r1: m.r1, c1: m.c1 } : { r0: p.r, c0: p.c, r1: p.r, c1: p.c }
      const active = m ? { r: m.r0, c: m.c0 } : p
      setView(tab, { sel: { ranges: add ? [...v.sel.ranges, range0] : [range0], active, focus: p } })
      drag.current = { kind: 'select', add, start: active }
    }
    inputRef.current?.focus({ preventScroll: true })
  }

  const autoScroll = useRef<number | null>(null)
  const onPointerMove = (e: React.PointerEvent) => {
    if (!L) return
    const { x, y } = pos(e)
    const d = drag.current
    const el = scrollRef.current!
    if (!d) {
      // Cursor hints
      const h = hit(L, x, y)
      const cur = h.zone === 'colhdr' && nearEdge(L, x, 'col') !== null ? 'col-resize' : h.zone === 'rowhdr' && nearEdge(L, y, 'row') !== null ? 'row-resize' : h.zone === 'colhdr' ? 's-resize' : h.zone === 'rowhdr' ? 'e-resize' : 'cell'
      el.style.cursor = cur === 'cell' ? (nearFillHandle(L, x, y) ? 'crosshair' : 'cell') : cur
      return
    }
    // Auto-scroll when dragging past the edge.
    const edgeX = x > L.width - 10 ? 1 : x < L.headerW + 4 ? -1 : 0
    const edgeY = y > L.height - 10 ? 1 : y < L.headerH + 4 ? -1 : 0
    if ((edgeX || edgeY) && autoScroll.current === null && d.kind !== 'resize-col' && d.kind !== 'resize-row') {
      autoScroll.current = window.setInterval(() => {
        el.scrollLeft += edgeX * 30
        el.scrollTop += edgeY * 30
      }, 40)
    } else if (!edgeX && !edgeY && autoScroll.current !== null) {
      clearInterval(autoScroll.current)
      autoScroll.current = null
    }
    const h = hit(L, Math.max(L.headerW + 1, Math.min(x, L.width - 1)), Math.max(L.headerH + 1, Math.min(y, L.height - 1)))
    const v = getView(tab)
    if (d.kind === 'select') {
      const p = { r: h.r, c: h.c }
      const r = A.expandMerges(rectOf(d.start, p), meta.merges)
      setView(tab, { sel: { ...v.sel, ranges: [...v.sel.ranges.slice(0, -1), r], focus: p } })
    } else if (d.kind === 'cols') {
      setView(tab, { sel: { ...v.sel, ranges: [...v.sel.ranges.slice(0, -1), { r0: 0, r1: MAX_ROWS - 1, c0: Math.min(d.start, h.c), c1: Math.max(d.start, h.c) }], focus: { r: MAX_ROWS - 1, c: h.c } } })
    } else if (d.kind === 'rows') {
      setView(tab, { sel: { ...v.sel, ranges: [...v.sel.ranges.slice(0, -1), { r0: Math.min(d.start, h.r), r1: Math.max(d.start, h.r), c0: 0, c1: MAX_COLS - 1 }], focus: { r: h.r, c: MAX_COLS - 1 } } })
    } else if (d.kind === 'resize-col') {
      setGuide({ x: Math.max(colX(L, d.index) + 2, x) })
    } else if (d.kind === 'resize-row') {
      setGuide({ y: Math.max(rowY(L, d.index) + 2, y) })
    } else if (d.kind === 'fill') {
      const s = d.src
      // Extend in the direction of the larger distance outside the source.
      const dr = h.r > s.r1 ? h.r - s.r1 : h.r < s.r0 ? h.r - s.r0 : 0
      const dc = h.c > s.c1 ? h.c - s.c1 : h.c < s.c0 ? h.c - s.c0 : 0
      if (!dr && !dc) setFillPreview(null)
      else if (Math.abs(dr) >= Math.abs(dc)) setFillPreview(dr > 0 ? { ...s, r1: h.r } : { ...s, r0: h.r })
      else setFillPreview(dc > 0 ? { ...s, c1: h.c } : { ...s, c0: h.c })
    } else if (d.kind === 'point') {
      const c = A.ctx()
      if (c) insertPointRef({ ...c, view: v }, d.anchor, { r: h.r, c: h.c })
    }
  }

  const onPointerUp = async (e: React.PointerEvent) => {
    if (autoScroll.current !== null) {
      clearInterval(autoScroll.current)
      autoScroll.current = null
    }
    const d = drag.current
    drag.current = null
    if (!d || !L) return
    const c = A.ctx()
    if (!c) return
    const { x, y } = pos(e)
    if (d.kind === 'resize-col') {
      setGuide(null)
      const w = Math.max(0, Math.round(d.size + (x - d.from) / zoom))
      if (Math.abs(x - d.from) > 1) await A.edit(c, (b) => b.api.setColWidths(c.view.sheet, d.indexes, w))
    } else if (d.kind === 'resize-row') {
      setGuide(null)
      const hgt = Math.max(0, Math.round(d.size + (y - d.from) / zoom))
      if (Math.abs(y - d.from) > 1) await A.edit(c, (b) => b.api.setRowHeights(c.view.sheet, d.indexes, hgt))
    } else if (d.kind === 'fill') {
      const dst = fillPreview
      setFillPreview(null)
      if (dst) {
        const ok = await A.edit(c, (b) => b.api.autofill({ sheet: c.view.sheet, ...d.src }, { sheet: c.view.sheet, ...dst }))
        if (ok) setView(tab, (v) => ({ sel: { ...v.sel, ranges: [dst], focus: { r: dst.r1, c: dst.c1 } } }))
      }
    }
    if (!getView(tab).editing) inputRef.current?.focus({ preventScroll: true })
  }

  const onDoubleClick = async (e: React.MouseEvent) => {
    if (!L) return
    const { x, y } = pos(e)
    const h = hit(L, x, y)
    const c = A.ctx()
    if (!c) return
    if (h.zone === 'colhdr') {
      const edge = nearEdge(L, x, 'col')
      if (edge !== null) await autofit(c, edge)
      return
    }
    if (h.zone === 'rowhdr') {
      const edge = nearEdge(L, y, 'row')
      if (edge !== null) await A.edit(c, (b) => b.api.setRowHeights(c.view.sheet, [edge], undefined))
      return
    }
    if (h.zone === 'cell') await A.startEdit(c)
  }

  const onWheel = (e: React.WheelEvent) => {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault()
      setView(tab, (v) => ({ zoom: Math.min(4, Math.max(0.1, Math.round((v.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)) * 100) / 100)) }))
    }
  }

  // Editor placement over the active cell
  const editing = view.editing
  let editorStyle: React.CSSProperties = { left: -10000, top: 0, width: 1, height: 1, opacity: 0 }
  if (L) {
    const a = view.sel.active
    const m = meta.merges.find((mm) => a.r >= mm.r0 && a.r <= mm.r1 && a.c >= mm.c0 && a.c <= mm.c1)
    const x0 = colX(L, m ? m.c0 : a.c)
    const y0 = rowY(L, m ? m.r0 : a.r)
    const w = m ? colX(L, m.c1 + 1) - x0 : L.cols.size(a.c) * zoom
    const h = m ? rowY(L, m.r1 + 1) - y0 : L.rows.size(a.r) * zoom
    if (editing && editing.where === 'cell') {
      const st = book.meta.styles[book.cell(view.sheet, a.r, a.c)?.s ?? 0] ?? {}
      const font = fontOf(st, zoom, defaultFont, defaultFontSize)
      const lines = editing.text.split('\n')
      const longest = Math.max(...lines.map((l) => measureText(font, l)))
      editorStyle = {
        left: x0 - 1,
        top: y0 - 1,
        minWidth: w + 2,
        width: Math.min(Math.max(w + 2, longest + 16), L.width - x0),
        minHeight: h + 2,
        height: Math.max(h + 2, lines.length * (st.size ?? defaultFontSize) * 1.333 * zoom * 1.25 + 6),
        font,
        color: st.color ?? (dark ? '#e7e7ea' : '#000'),
        background: st.fill ?? (dark ? '#141416' : '#fff'),
        textAlign: 'left',
        zIndex: 5
      }
    } else editorStyle = { left: x0, top: y0, width: Math.max(4, w), height: Math.max(4, h), opacity: 0, zIndex: 1 }
  }

  return (
    <div ref={hostRef} className="relative h-full w-full overflow-hidden bg-[var(--grid-bg)]" hidden={hidden} data-testid="grid">
      <canvas ref={canvasRef} className="pointer-events-none absolute left-0 top-0" aria-hidden="true" />
      <GridMenu tab={tab}>
        <div
          ref={scrollRef}
          className="absolute inset-0 overflow-auto"
          onScroll={onScroll}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onDoubleClick={onDoubleClick}
          onWheel={onWheel}
          data-testid="grid-scroll"
        >
          <div style={{ width: spacer.w, height: spacer.h }} />
        </div>
      </GridMenu>
      <textarea
        ref={inputRef}
        data-grid-focus=""
        hidden={hidden}
        aria-label={`Cell ${A.areaName({ r0: view.sel.active.r, c0: view.sel.active.c, r1: view.sel.active.r, c1: view.sel.active.c })}`}
        aria-describedby={`grid-live-${tab}`}
        spellCheck={false}
        className={editing && editing.where === 'cell' ? 'cell-editor absolute' : 'absolute resize-none overflow-hidden border-0 bg-transparent p-0 text-transparent caret-transparent outline-none'}
        style={editorStyle}
        value={editing ? editing.text : ''}
        onChange={(e) => {
          const c = A.ctx()
          if (!c) return
          const t = e.target.value
          if (!c.view.editing) {
            // Typing starts editing (and replaces the cell's content).
            setView(tab, { editing: { text: t, where: 'cell', enterMode: true, caret: t.length } })
          } else A.updateEdit(c, t, e.target.selectionStart ?? t.length)
        }}
        onSelect={(e) => {
          const c = A.ctx()
          const ed = c?.view.editing
          if (c && ed) {
            const pos = (e.target as HTMLTextAreaElement).selectionStart ?? 0
            if (pos !== ed.caret) setView(tab, { editing: { ...ed, caret: pos } })
          }
        }}
        onKeyDown={(e) => {
          const c = A.ctx()
          if (c) handleGridKey(e, { ...c, view: getView(tab) }, L)
        }}
        onCopy={(e) => {
          if (getView(tab).editing) return
          e.preventDefault()
          const c = A.ctx()
          if (c) void A.copy(c)
        }}
        onCut={(e) => {
          if (getView(tab).editing) return
          e.preventDefault()
          const c = A.ctx()
          if (c) void A.copy(c, true)
        }}
        onPaste={(e) => {
          if (getView(tab).editing) return
          e.preventDefault()
          const c = A.ctx()
          if (c) void A.paste(c)
        }}
      />
      {L && <FilterOverlay tab={tab} book={book} L={L} />}
      {L && <ValidationOverlay tab={tab} book={book} L={L} />}
      {guide && <div className="pointer-events-none absolute z-10 bg-accent" style={guide.x !== undefined ? { left: guide.x, top: 0, width: 1, height: '100%' } : { top: guide.y, left: 0, height: 1, width: '100%' }} />}
      <LiveRegion tab={tab} />
    </div>
  )
}

/** Screen-reader announcement of the active cell (the canvas itself is not readable). */
function LiveRegion({ tab }: { tab: string }) {
  const view = useViews((s) => s.views[tab])
  const [text, setText] = useState('')
  useEffect(() => {
    const book = books.get(tab)
    if (!book || !view) return
    const { r, c } = view.sel.active
    let cancelled = false
    void book.api.cellInfo(view.sheet, r, c).then((info) => {
      if (cancelled) return
      const shown = book.cell(view.sheet, r, c)?.t ?? ''
      const name = A.areaName({ r0: r, c0: c, r1: r, c1: c })
      setText(`${name}, ${shown || 'blank'}${info.formula ? `, formula ${info.editText}` : ''}`)
    })
    return () => {
      cancelled = true
    }
  }, [view?.sel.active.r, view?.sel.active.c, view?.sheet])
  return (
    <div id={`grid-live-${tab}`} className="sr-only" aria-live="polite" role="status">
      {text}
    </div>
  )
}

let mctx: CanvasRenderingContext2D | null = null
function measureText(font: string, text: string) {
  mctx ??= document.createElement('canvas').getContext('2d')
  mctx!.font = font
  return mctx!.measureText(text).width
}

function nearEdge(L: Layout, px: number, axis: 'col' | 'row'): number | null {
  if (axis === 'col') {
    const h = hit(L, px, L.headerH + 1)
    for (const c of [h.c - 1, h.c]) {
      if (c < 0) continue
      const edge = colX(L, c + 1)
      if (Math.abs(px - edge) <= 4 && px >= L.headerW) return c
    }
    return null
  }
  const h = hit(L, L.headerW + 1, px)
  for (const r of [h.r - 1, h.r]) {
    if (r < 0) continue
    const edge = rowY(L, r + 1)
    if (Math.abs(px - edge) <= 3 && px >= L.headerH) return r
  }
  return null
}

function nearFillHandle(L: Layout, x: number, y: number) {
  const c = A.ctx()
  if (!c || c.view.sel.ranges.length !== 1) return false
  const r = c.view.sel.ranges[0]
  return Math.abs(x - colX(L, r.c1 + 1)) <= 5 && Math.abs(y - rowY(L, r.r1 + 1)) <= 5
}

/** Double-click a column border: fit the widest displayed text (measured with the real fonts). */
export async function autofit(c: A.Ctx, col: number) {
  const s = getView(c.tab)
  const sel = s.sel.ranges.filter((r) => r.r0 === 0 && r.r1 === MAX_ROWS - 1 && col >= r.c0 && col <= r.c1)
  const cols = sel.length ? sel.flatMap((r) => range(r.c0, r.c1)) : [col]
  const { defaultFont, defaultFontSize } = get().settings
  for (const cc of cols) {
    const texts = await c.book.api.columnTexts(c.view.sheet, cc)
    let w = 0
    for (const t of texts) {
      const st = c.book.meta.styles[t.s] ?? {}
      w = Math.max(w, measureText(fontOf(st, 1, defaultFont, defaultFontSize), t.t) + 10)
    }
    const m = c.book.sheetMeta(c.view.sheet)!
    await A.edit(c, (b) => b.api.setColWidths(c.view.sheet, [cc], texts.length ? Math.ceil(Math.min(Math.max(w, 16), 1200)) : m.defaultColWidth))
  }
}

function range(a: number, b: number) {
  return Array.from({ length: b - a + 1 }, (_, i) => a + i)
}

export { axesFor, set }
