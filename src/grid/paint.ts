// Drawing a sheet onto a canvas: fills, gridlines, text (with overflow), borders, merges, selection and headers.
import type { ViewCell } from '@/core/doc'
import type { SheetMeta } from '@/core/host'
import type { BorderSide, Style } from '@/core/model/styles'
import { colName } from '@/core/refs'
import { MAX_COLS, MAX_ROWS, type Area } from '@/core/engine/values'
import { colX, rowY, visibleCols, visibleRows, type Layout } from './layout'
import type { Sel } from '@/renderer/state/view'

export interface Theme {
  bg: string
  line: string
  header: string
  headerText: string
  headerActive: string
  accent: string
  selFill: string
  text: string
  dark: boolean
}

export interface PaintInput {
  L: Layout
  meta: SheetMeta
  styles: Style[]
  cell: (r: number, c: number) => ViewCell | undefined
  sel: Sel | null
  /** Hide the active cell's text while the in-cell editor covers it. */
  editingCell: { r: number; c: number } | null
  copy: (Area & { cut: boolean }) | null
  dashOffset: number
  theme: Theme
  defaultFont: string
  defaultSize: number
  showFormulas: boolean
  formulaText?: (r: number, c: number) => string | undefined
  /** Coloured reference outlines while a formula is edited. */
  refHighlights: { area: { r0: number; c0: number; r1: number; c1: number }; color: string }[]
  /** Autofill preview area while dragging the fill handle. */
  fillPreview: { r0: number; c0: number; r1: number; c1: number } | null
}

const PT = 96 / 72
const widthCache = new Map<string, number>()

export function measure(ctx: CanvasRenderingContext2D, font: string, text: string): number {
  const key = font + '\u0000' + text
  let w = widthCache.get(key)
  if (w === undefined) {
    ctx.font = font
    w = ctx.measureText(text).width
    if (widthCache.size > 20000) widthCache.clear()
    widthCache.set(key, w)
  }
  return w
}

export function fontOf(s: Style, zoom: number, defFont: string, defSize: number): string {
  const size = (s.size ?? defSize) * PT * zoom
  const family = s.font ?? defFont
  return `${s.italic ? 'italic ' : ''}${s.bold ? '700 ' : ''}${size.toFixed(2)}px "${family}", Calibri, Carlito, sans-serif`
}

const BORDER_W: Record<string, number> = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 }
const BORDER_DASH: Record<string, number[]> = { dotted: [1, 2], hair: [1, 1], dashed: [3, 2], mediumDashed: [6, 3], dashDot: [6, 2, 2, 2], mediumDashDot: [8, 3, 3, 3], dashDotDot: [6, 2, 2, 2, 2, 2], mediumDashDotDot: [8, 3, 3, 3, 3, 3], slantDashDot: [8, 3, 3, 3] }

function strokeBorder(ctx: CanvasRenderingContext2D, b: BorderSide, x0: number, y0: number, x1: number, y1: number, theme: Theme) {
  const w = BORDER_W[b.style] ?? 1
  ctx.strokeStyle = b.color ?? (theme.dark ? '#d4d4d8' : '#000000')
  ctx.setLineDash(BORDER_DASH[b.style] ?? [])
  if (b.style === 'double') {
    ctx.lineWidth = 1
    const horiz = y0 === y1
    for (const d of [-1, 1]) {
      ctx.beginPath()
      if (horiz) {
        ctx.moveTo(x0, y0 + d + 0.5)
        ctx.lineTo(x1, y1 + d + 0.5)
      } else {
        ctx.moveTo(x0 + d + 0.5, y0)
        ctx.lineTo(x1 + d + 0.5, y1)
      }
      ctx.stroke()
    }
    ctx.setLineDash([])
    return
  }
  ctx.lineWidth = w
  const off = w % 2 ? 0.5 : 0
  ctx.beginPath()
  if (y0 === y1) {
    ctx.moveTo(x0, Math.round(y0) + off)
    ctx.lineTo(x1, Math.round(y1) + off)
  } else {
    ctx.moveTo(Math.round(x0) + off, y0)
    ctx.lineTo(Math.round(x1) + off, y1)
  }
  ctx.stroke()
  ctx.setLineDash([])
}

interface Pane {
  r0: number
  r1: number
  c0: number
  c1: number
  x: number
  y: number
  w: number
  h: number
}

function panes(L: Layout): Pane[] {
  const [sr0, sr1] = visibleRows(L)
  const [sc0, sc1] = visibleCols(L)
  const fx = L.headerW + L.frozenW * L.zoom
  const fy = L.headerH + L.frozenH * L.zoom
  const out: Pane[] = []
  const add = (r0: number, r1: number, c0: number, c1: number, x: number, y: number, w: number, h: number) => {
    if (r1 >= r0 && c1 >= c0 && w > 0 && h > 0) out.push({ r0, r1, c0, c1, x, y, w, h })
  }
  add(sr0, sr1, sc0, sc1, fx, fy, L.width - fx, L.height - fy)
  if (L.fc) add(sr0, sr1, 0, L.fc - 1, L.headerW, fy, fx - L.headerW, L.height - fy)
  if (L.fr) add(0, L.fr - 1, sc0, sc1, fx, L.headerH, L.width - fx, fy - L.headerH)
  if (L.fr && L.fc) add(0, L.fr - 1, 0, L.fc - 1, L.headerW, L.headerH, fx - L.headerW, fy - L.headerH)
  return out
}

export function paint(ctx: CanvasRenderingContext2D, P: PaintInput) {
  const { L, theme } = P
  ctx.fillStyle = theme.bg
  ctx.fillRect(0, 0, L.width, L.height)
  for (const pane of panes(L)) {
    ctx.save()
    ctx.beginPath()
    ctx.rect(pane.x, pane.y, pane.w, pane.h)
    ctx.clip()
    paintCells(ctx, P, pane)
    ctx.restore()
  }
  // Selection over all panes (clipped to the cell area).
  ctx.save()
  ctx.beginPath()
  ctx.rect(L.headerW, L.headerH, L.width - L.headerW, L.height - L.headerH)
  ctx.clip()
  paintOverlays(ctx, P)
  ctx.restore()
  if (L.showHeaders) paintHeaders(ctx, P)
  // Frozen pane dividers
  ctx.strokeStyle = theme.dark ? '#55555f' : '#b5b5bd'
  ctx.lineWidth = 1
  if (L.fc) {
    const x = Math.round(L.headerW + L.frozenW * L.zoom) + 0.5
    ctx.beginPath()
    ctx.moveTo(x, 0)
    ctx.lineTo(x, L.height)
    ctx.stroke()
  }
  if (L.fr) {
    const y = Math.round(L.headerH + L.frozenH * L.zoom) + 0.5
    ctx.beginPath()
    ctx.moveTo(0, y)
    ctx.lineTo(L.width, y)
    ctx.stroke()
  }
}

function cellRect(L: Layout, r: number, c: number) {
  const x = colX(L, c)
  const y = rowY(L, r)
  return { x, y, w: L.cols.size(c) * L.zoom, h: L.rows.size(r) * L.zoom }
}

function areaRect(L: Layout, a: { r0: number; c0: number; r1: number; c1: number }) {
  const x0 = colX(L, a.c0)
  const y0 = rowY(L, a.r0)
  const x1 = colX(L, Math.min(a.c1, MAX_COLS - 1)) + L.cols.size(Math.min(a.c1, MAX_COLS - 1)) * L.zoom
  const y1 = rowY(L, Math.min(a.r1, MAX_ROWS - 1)) + L.rows.size(Math.min(a.r1, MAX_ROWS - 1)) * L.zoom
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

function paintCells(ctx: CanvasRenderingContext2D, P: PaintInput, pane: Pane) {
  const { L, theme, meta, styles } = P
  const z = L.zoom
  const merges = meta.merges.filter((m) => m.r1 >= pane.r0 && m.r0 <= pane.r1 && m.c1 >= pane.c0 && m.c0 <= pane.c1)
  const mergeAt = (r: number, c: number) => merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1)
  // 1. Fills
  for (let r = pane.r0; r <= pane.r1; r++) {
    const h = L.rows.size(r)
    if (!h) continue
    for (let c = pane.c0; c <= pane.c1; c++) {
      const v = P.cell(r, c)
      if (!v?.s) continue
      const st = styles[v.s]
      if (!st?.fill && !st?.pattern) continue
      const m = mergeAt(r, c)
      if (m && (m.r0 !== r || m.c0 !== c)) continue
      const rect = m ? areaRect(L, m) : cellRect(L, r, c)
      ctx.fillStyle = st.fill ?? st.pattern?.fg ?? '#fff'
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h)
    }
  }
  // 2. Gridlines (not under fills or inside merges)
  if (meta.showGridlines) {
    ctx.strokeStyle = theme.line
    ctx.lineWidth = 1
    ctx.beginPath()
    for (let c = pane.c0; c <= pane.c1 + 1; c++) {
      if (c <= MAX_COLS && L.cols.size(Math.min(c, MAX_COLS - 1)) === 0 && c <= pane.c1) continue
      const x = Math.round(colX(L, c)) - 0.5
      ctx.moveTo(x, pane.y)
      ctx.lineTo(x, pane.y + pane.h)
    }
    for (let r = pane.r0; r <= pane.r1 + 1; r++) {
      if (r <= pane.r1 && L.rows.size(r) === 0) continue
      const y = Math.round(rowY(L, r)) - 0.5
      ctx.moveTo(pane.x, y)
      ctx.lineTo(pane.x + pane.w, y)
    }
    ctx.stroke()
    // Cover gridlines under filled cells and inside merges.
    for (let r = pane.r0; r <= pane.r1; r++)
      for (let c = pane.c0; c <= pane.c1; c++) {
        const v = P.cell(r, c)
        const st = v?.s ? styles[v.s] : undefined
        if (st?.fill) {
          const rect = cellRect(L, r, c)
          ctx.fillStyle = st.fill
          ctx.fillRect(rect.x - 1, rect.y - 1, rect.w + 1, rect.h + 1)
        }
      }
  }
  for (const m of merges) {
    const rect = areaRect(L, m)
    const st = styles[P.cell(m.r0, m.c0)?.s ?? 0]
    ctx.fillStyle = st?.fill ?? theme.bg
    ctx.fillRect(rect.x, rect.y, rect.w - 1, rect.h - 1)
  }
  // 3. Text
  const padX = 3 * z
  for (let r = pane.r0; r <= pane.r1; r++) {
    const rh = L.rows.size(r) * z
    if (!rh) continue
    const y = rowY(L, r)
    // Text overflow needs to know which neighbours are empty.
    for (let c = pane.c0; c <= pane.c1; c++) {
      const v = P.cell(r, c)
      if (!v) continue
      if (P.editingCell && P.editingCell.r === r && P.editingCell.c === c) continue
      const m = mergeAt(r, c)
      if (m && (m.r0 !== r || m.c0 !== c)) continue
      let text = v.t
      if (P.showFormulas && v.f && P.formulaText) text = P.formulaText(r, c) ?? text
      if (!text) continue
      const st = styles[v.s] ?? {}
      const font = fontOf(st, z, P.defaultFont, P.defaultSize)
      const rect = m ? areaRect(L, m) : { x: colX(L, c), y, w: L.cols.size(c) * z, h: rh }
      if (rect.w <= 0) continue
      let align = st.hAlign && st.hAlign !== 'general' ? st.hAlign : v.k === 'n' || v.k === 'd' ? 'right' : v.k === 'b' || v.k === 'e' ? 'center' : 'left'
      if (P.showFormulas && v.f) align = 'left'
      const indent = (st.indent ?? 0) * 9 * z
      ctx.font = font
      ctx.fillStyle = v.fc ?? st.color ?? theme.text
      const tw = measure(ctx, font, text)
      // Numbers that don't fit show ####.
      if ((v.k === 'n' || v.k === 'd') && !P.showFormulas && tw > rect.w - 2 * padX && !st.shrink) {
        const hw = measure(ctx, font, '#')
        text = '#'.repeat(Math.max(1, Math.floor((rect.w - 2 * padX) / hw)))
      }
      const wrap = st.wrap && !P.showFormulas
      const lines = wrap ? wrapText(ctx, font, text, rect.w - 2 * padX) : text.split('\n').slice(0, 1).length && text.includes('\n') ? [text.split('\n')[0]] : [text]
      const lineH = (st.size ?? P.defaultSize) * PT * z * 1.2
      const blockH = lines.length * lineH
      const va = st.vAlign ?? 'bottom'
      let ty = va === 'top' ? rect.y + 2 * z : va === 'center' ? rect.y + (rect.h - blockH) / 2 : rect.y + rect.h - blockH - 2 * z
      // Clip: own cell, or overflow into empty neighbours for unwrapped text.
      let clipX0 = rect.x
      let clipX1 = rect.x + rect.w
      if (!wrap && !m && v.k === 's' && tw > rect.w - 2 * padX) {
        if (align === 'left' || align === 'center') for (let cc = c + 1; cc <= Math.min(pane.c1 + 20, MAX_COLS - 1) && clipX1 < rect.x + tw + 2 * padX + 1; cc++) {
          if (P.cell(r, cc)?.t || mergeAt(r, cc)) break
          clipX1 += L.cols.size(cc) * z
        }
        if (align === 'right' || align === 'center') for (let cc = c - 1; cc >= Math.max(0, pane.c0 - 20) && clipX0 > rect.x + rect.w - tw - 2 * padX - 1; cc--) {
          if (P.cell(r, cc)?.t || mergeAt(r, cc)) break
          clipX0 -= L.cols.size(cc) * z
        }
      }
      ctx.save()
      ctx.beginPath()
      ctx.rect(clipX0, rect.y, clipX1 - clipX0, rect.h)
      ctx.clip()
      ctx.textBaseline = 'alphabetic'
      for (const line of lines) {
        const lw = measure(ctx, font, line)
        const tx = align === 'right' ? rect.x + rect.w - padX - lw - indent : align === 'center' || align === 'centerContinuous' ? rect.x + (rect.w - lw) / 2 : rect.x + padX + indent
        const base = ty + lineH * 0.8
        ctx.fillText(line, tx, base)
        if (st.underline) {
          ctx.fillRect(tx, base + 1.5 * z, lw, Math.max(1, z))
          if (st.underline.startsWith('double')) ctx.fillRect(tx, base + 3.5 * z, lw, Math.max(1, z))
        }
        if (st.strike) ctx.fillRect(tx, base - lineH * 0.28, lw, Math.max(1, z))
        ty += lineH
      }
      ctx.restore()
    }
  }
  // 4. Borders
  for (let r = pane.r0; r <= pane.r1; r++)
    for (let c = pane.c0; c <= pane.c1; c++) {
      const v = P.cell(r, c)
      const st = v?.s ? styles[v.s] : undefined
      if (!st || !(st.borderTop || st.borderBottom || st.borderLeft || st.borderRight || st.borderDiagDown || st.borderDiagUp)) continue
      const m = mergeAt(r, c)
      const rect = m && m.r0 === r && m.c0 === c ? areaRect(L, m) : cellRect(L, r, c)
      const x0 = rect.x - 1
      const y0 = rect.y - 1
      const x1 = rect.x + rect.w - 1
      const y1 = rect.y + rect.h - 1
      if (st.borderTop) strokeBorder(ctx, st.borderTop, x0, y0, x1 + 1, y0, theme)
      if (st.borderBottom) strokeBorder(ctx, st.borderBottom, x0, y1, x1 + 1, y1, theme)
      if (st.borderLeft) strokeBorder(ctx, st.borderLeft, x0, y0, x0, y1 + 1, theme)
      if (st.borderRight) strokeBorder(ctx, st.borderRight, x1, y0, x1, y1 + 1, theme)
      if (st.borderDiagDown) {
        ctx.strokeStyle = st.borderDiagDown.color ?? theme.text
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0, y0)
        ctx.lineTo(x1, y1)
        ctx.stroke()
      }
      if (st.borderDiagUp) {
        ctx.strokeStyle = st.borderDiagUp.color ?? theme.text
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0, y1)
        ctx.lineTo(x1, y0)
        ctx.stroke()
      }
    }
}

function wrapText(ctx: CanvasRenderingContext2D, font: string, text: string, width: number): string[] {
  const out: string[] = []
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/)
    let line = ''
    for (const w of words) {
      const next = line + w
      if (line && measure(ctx, font, next.trimEnd()) > width) {
        out.push(line.trimEnd())
        line = w.trimStart()
      } else line = next
    }
    out.push(line)
  }
  return out
}

function paintOverlays(ctx: CanvasRenderingContext2D, P: PaintInput) {
  const { L, theme, sel } = P
  // Spilled ranges around the active cell
  if (sel) {
    const a = sel.active
    const sp = P.meta.spills.find((s) => a.r >= s.r0 && a.r <= s.r1 && a.c >= s.c0 && a.c <= s.c1)
    if (sp) {
      const rect = areaRect(L, sp)
      ctx.strokeStyle = '#2f6fde'
      ctx.setLineDash([4, 3])
      ctx.lineWidth = 1
      ctx.strokeRect(rect.x - 0.5, rect.y - 0.5, rect.w, rect.h)
      ctx.setLineDash([])
    }
  }
  for (const h of P.refHighlights) {
    const rect = areaRect(L, h.area)
    ctx.fillStyle = h.color + '1f'
    ctx.fillRect(rect.x, rect.y, rect.w, rect.h)
    ctx.strokeStyle = h.color
    ctx.lineWidth = 2
    ctx.strokeRect(rect.x, rect.y, rect.w - 1, rect.h - 1)
  }
  if (sel) {
    const ar = sel.ranges[sel.ranges.length - 1]
    const am = P.meta.merges.find((m) => sel.active.r >= m.r0 && sel.active.r <= m.r1 && sel.active.c >= m.c0 && sel.active.c <= m.c1)
    const ac = am ? areaRect(L, am) : cellRect(L, sel.active.r, sel.active.c)
    const single = sel.ranges.length === 1 && ((ar.r0 === ar.r1 && ar.c0 === ar.c1) || (am && am.r0 === ar.r0 && am.c0 === ar.c0 && am.r1 === ar.r1 && am.c1 === ar.c1))
    if (!single) {
      // Shade the selection, leaving the active cell clear.
      ctx.save()
      ctx.beginPath()
      ctx.rect(0, 0, L.width, L.height)
      ctx.rect(ac.x, ac.y, ac.w, ac.h)
      ctx.clip('evenodd')
      ctx.fillStyle = theme.selFill
      for (const r of sel.ranges) {
        const rect = areaRect(L, r)
        ctx.fillRect(rect.x, rect.y, rect.w, rect.h)
      }
      ctx.restore()
    }
    // Border around the active range
    const rect = areaRect(L, ar)
    ctx.strokeStyle = theme.accent
    ctx.lineWidth = 2
    ctx.strokeRect(rect.x - 1, rect.y - 1, rect.w + 1, rect.h + 1)
    if (sel.ranges.length > 1) {
      ctx.lineWidth = 1
      ctx.strokeRect(ac.x - 0.5, ac.y - 0.5, ac.w, ac.h)
    }
    // Fill handle
    if (sel.ranges.length === 1) {
      ctx.fillStyle = theme.accent
      ctx.fillRect(rect.x + rect.w - 4, rect.y + rect.h - 4, 7, 7)
      ctx.fillStyle = theme.bg
      ctx.fillRect(rect.x + rect.w - 4, rect.y + rect.h - 4, 7, 1)
      ctx.fillRect(rect.x + rect.w - 4, rect.y + rect.h - 4, 1, 7)
    }
  }
  if (P.fillPreview) {
    const rect = areaRect(L, P.fillPreview)
    ctx.strokeStyle = theme.dark ? '#a1a1aa' : '#52525b'
    ctx.setLineDash([3, 2])
    ctx.lineWidth = 1
    ctx.strokeRect(rect.x - 0.5, rect.y - 0.5, rect.w, rect.h)
    ctx.setLineDash([])
  }
  if (P.copy && P.copy.sheet === P.meta.id) {
    const rect = areaRect(L, P.copy)
    ctx.strokeStyle = theme.accent
    ctx.lineWidth = 2
    ctx.setLineDash([5, 4])
    ctx.lineDashOffset = -P.dashOffset
    ctx.strokeRect(rect.x - 1, rect.y - 1, rect.w + 1, rect.h + 1)
    ctx.setLineDash([])
    ctx.lineDashOffset = 0
  }
}

function paintHeaders(ctx: CanvasRenderingContext2D, P: PaintInput) {
  const { L, theme, sel } = P
  const z = L.zoom
  const selCols = new Set<number>()
  const selRows = new Set<number>()
  const fullCols = new Set<number>()
  const fullRows = new Set<number>()
  const [sc0, sc1] = visibleCols(L)
  const [sr0, sr1] = visibleRows(L)
  const colsShown = [...range(0, L.fc - 1), ...range(sc0, sc1)]
  const rowsShown = [...range(0, L.fr - 1), ...range(sr0, sr1)]
  if (sel)
    for (const r of sel.ranges) {
      for (const c of colsShown) if (c >= r.c0 && c <= r.c1) selCols.add(c)
      for (const rr of rowsShown) if (rr >= r.r0 && rr <= r.r1) selRows.add(rr)
      if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) for (const c of colsShown) if (c >= r.c0 && c <= r.c1) fullCols.add(c)
      if (r.c0 === 0 && r.c1 === MAX_COLS - 1) for (const rr of rowsShown) if (rr >= r.r0 && rr <= r.r1) fullRows.add(rr)
    }
  ctx.font = `${(11 * Math.min(Math.max(z, 0.8), 1.4)).toFixed(1)}px "Inter Variable", system-ui, sans-serif`
  ctx.textBaseline = 'middle'
  ctx.textAlign = 'center'
  // Column header band
  ctx.fillStyle = theme.header
  ctx.fillRect(0, 0, L.width, L.headerH)
  ctx.fillRect(0, 0, L.headerW, L.height)
  for (const c of colsShown) {
    const w = L.cols.size(c) * z
    if (!w) continue
    const x = colX(L, c)
    if (x > L.width || x + w < L.headerW) continue
    if (c >= L.fc && x < L.headerW + L.frozenW * z - 1) continue
    if (selCols.has(c)) {
      ctx.fillStyle = fullCols.has(c) ? theme.accent : theme.headerActive
      ctx.fillRect(x, 0, w, L.headerH)
    }
    ctx.fillStyle = fullCols.has(c) ? '#fff' : selCols.has(c) ? theme.accent : theme.headerText
    if (w > 10) ctx.fillText(colName(c), x + w / 2, L.headerH / 2 + 0.5)
    ctx.fillStyle = theme.line
    ctx.fillRect(Math.round(x + w) - 1, 0, 1, L.headerH)
  }
  for (const r of rowsShown) {
    const h = L.rows.size(r) * z
    if (!h) continue
    const y = rowY(L, r)
    if (y > L.height || y + h < L.headerH) continue
    if (r >= L.fr && y < L.headerH + L.frozenH * z - 1) continue
    if (selRows.has(r)) {
      ctx.fillStyle = fullRows.has(r) ? theme.accent : theme.headerActive
      ctx.fillRect(0, y, L.headerW, h)
    }
    ctx.fillStyle = fullRows.has(r) ? '#fff' : selRows.has(r) ? theme.accent : theme.headerText
    if (h > 8) ctx.fillText(String(r + 1), L.headerW / 2, y + h / 2 + 0.5)
    ctx.fillStyle = theme.line
    ctx.fillRect(0, Math.round(y + h) - 1, L.headerW, 1)
  }
  // Bottom/right edges of the header bands and the corner
  ctx.fillStyle = theme.line
  ctx.fillRect(0, L.headerH - 1, L.width, 1)
  ctx.fillRect(L.headerW - 1, 0, 1, L.height)
  ctx.fillStyle = theme.header
  ctx.fillRect(0, 0, L.headerW - 1, L.headerH - 1)
  ctx.fillStyle = theme.headerText
  ctx.beginPath()
  ctx.moveTo(L.headerW - 5, L.headerH - 5)
  ctx.lineTo(L.headerW - 5, L.headerH - 13)
  ctx.lineTo(L.headerW - 13, L.headerH - 5)
  ctx.closePath()
  ctx.globalAlpha = 0.45
  ctx.fill()
  ctx.globalAlpha = 1
  ctx.textAlign = 'start'
}

function range(a: number, b: number): number[] {
  const out: number[] = []
  for (let i = a; i <= b; i++) out.push(i)
  return out
}
