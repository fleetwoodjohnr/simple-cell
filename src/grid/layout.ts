// Screen geometry of a sheet view: headers, frozen panes, scrolling and zoom.
import { MAX_COLS, MAX_ROWS } from '@/core/engine/values'
import type { SheetMeta } from '@/core/host'
import { Axis } from './axis'

export interface Layout {
  zoom: number
  width: number
  height: number
  headerW: number
  headerH: number
  /** Outline gutters (group bars and +/− buttons) at the left of the row headers and above the column headers. */
  gutterW: number
  gutterH: number
  rows: Axis
  cols: Axis
  fr: number
  fc: number
  /** Unzoomed scroll offsets of the scrolling pane. */
  sx: number
  sy: number
  /** Unzoomed size of the frozen block. */
  frozenW: number
  frozenH: number
  showHeaders: boolean
}

const axisCache = new WeakMap<SheetMeta, { rows: Axis; cols: Axis }>()

export function axesFor(m: SheetMeta) {
  let a = axisCache.get(m)
  if (!a) {
    a = { rows: new Axis(m.defaultRowHeight, m.rows, MAX_ROWS), cols: new Axis(m.defaultColWidth, m.cols, MAX_COLS) }
    axisCache.set(m, a)
  }
  return a
}

export function makeLayout(m: SheetMeta, width: number, height: number, zoom: number, sx: number, sy: number, showHeaders: boolean): Layout {
  const { rows, cols } = axesFor(m)
  const fr = m.freeze.rows
  const fc = m.freeze.cols
  const frozenW = cols.offset(fc)
  const frozenH = rows.offset(fr)
  // Row header width grows with the number of digits of the last visible row.
  const lastRow = rows.indexAt(frozenH + sy + height / zoom)
  const digits = String(lastRow + 1).length
  const levels = (list: [number, number][]) => list.reduce((m, [, l]) => Math.max(m, l), 0)
  const rl = levels(m.outlineRows ?? [])
  const cl = levels(m.outlineCols ?? [])
  const gutterW = rl ? (rl + 1) * GUTTER_STEP + 4 : 0
  const gutterH = cl ? (cl + 1) * GUTTER_STEP + 4 : 0
  const headerW = gutterW + (showHeaders ? Math.max(32, Math.round((digits * 7.5 + 14) * Math.min(zoom, 1.5))) : 0)
  const headerH = gutterH + (showHeaders ? Math.round(22 * Math.min(Math.max(zoom, 0.75), 1.5)) : 0)
  return { zoom, width, height, headerW, headerH, gutterW, gutterH, rows, cols, fr, fc, sx, sy, frozenW, frozenH, showHeaders }
}

/** Width of one outline level in the gutter. */
export const GUTTER_STEP = 14

/** Screen x of the left edge of column c (may be off screen). */
export function colX(L: Layout, c: number): number {
  return L.headerW + L.zoom * (c < L.fc ? L.cols.offset(c) : L.cols.offset(c) - L.sx)
}
export function rowY(L: Layout, r: number): number {
  return L.headerH + L.zoom * (r < L.fr ? L.rows.offset(r) : L.rows.offset(r) - L.sy)
}

/** Range of scrolling-pane columns/rows on screen. */
export function visibleCols(L: Layout): [number, number] {
  const start = L.cols.indexAt(L.frozenW + L.sx)
  const end = L.cols.indexAt(L.sx + (L.width - L.headerW) / L.zoom)
  return [Math.max(start, L.fc), Math.min(MAX_COLS - 1, end)]
}
export function visibleRows(L: Layout): [number, number] {
  const start = L.rows.indexAt(L.frozenH + L.sy)
  const end = L.rows.indexAt(L.sy + (L.height - L.headerH) / L.zoom)
  return [Math.max(start, L.fr), Math.min(MAX_ROWS - 1, end)]
}

/** Cell under a screen point; null over headers. */
export function hit(L: Layout, x: number, y: number): { r: number; c: number; zone: 'cell' | 'colhdr' | 'rowhdr' | 'corner' | 'gutter' } {
  if (x < L.gutterW || y < L.gutterH) return { r: 0, c: 0, zone: 'gutter' }
  const inColHdr = y < L.headerH
  const inRowHdr = x < L.headerW
  const cx = (x - L.headerW) / L.zoom
  const cy = (y - L.headerH) / L.zoom
  const c = cx < L.frozenW ? L.cols.indexAt(Math.max(0, cx)) : L.cols.indexAt(cx + L.sx)
  const r = cy < L.frozenH ? L.rows.indexAt(Math.max(0, cy)) : L.rows.indexAt(cy + L.sy)
  const zone = inColHdr && inRowHdr ? 'corner' : inColHdr ? 'colhdr' : inRowHdr ? 'rowhdr' : 'cell'
  return { r, c, zone }
}

/** Scroll needed (unzoomed) so that a cell is fully visible; null if it already is. */
export function scrollToShow(L: Layout, r: number, c: number): { sx: number; sy: number } | null {
  let { sx, sy } = L
  const viewW = (L.width - L.headerW) / L.zoom - L.frozenW
  const viewH = (L.height - L.headerH) / L.zoom - L.frozenH
  if (c >= L.fc) {
    const left = L.cols.offset(c) - L.frozenW
    const right = L.cols.offset(c + 1) - L.frozenW
    if (left < sx) sx = left
    else if (right > sx + viewW) sx = Math.min(left, right - viewW)
  }
  if (r >= L.fr) {
    const top = L.rows.offset(r) - L.frozenH
    const bottom = L.rows.offset(r + 1) - L.frozenH
    if (top < sy) sy = top
    else if (bottom > sy + viewH) sy = Math.min(top, bottom - viewH)
  }
  sx = Math.max(0, sx)
  sy = Math.max(0, sy)
  return sx === L.sx && sy === L.sy ? null : { sx, sy }
}
