// Table styles: Excel's built-in families (Light 1–21, Medium 1–28, Dark 1–11) generated from the theme colours.
// ponytail: an approximation of Excel's 60 built-in definitions (fills, borders, bold) — close in look, not
// pixel-identical; custom table styles from files fall back to Medium 2.
import type { Table } from './model/workbook'
import type { BorderSide, Style } from './model/styles'

/** Office theme: lt1, dk1, lt2, dk2, accent1–6, hyperlink, followed hyperlink (hex without #). */
export const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']

interface Spec {
  header: Partial<Style>
  /** Odd data rows (row stripes), odd columns (column stripes). */
  band?: Partial<Style>
  body?: Partial<Style>
  total: Partial<Style>
  outline?: BorderSide
  hLine?: BorderSide
  vLine?: BorderSide
  headerLine?: BorderSide
}

function tint(hex: string, t: number): string {
  const n = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16))
  const f = (v: number) => Math.round(t >= 0 ? v + (255 - v) * t : v * (1 + t))
  return '#' + n.map((v) => f(v).toString(16).padStart(2, '0')).join('').toUpperCase()
}

function spec(name: string, theme: string[]): Spec {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name) ?? ['', 'Medium', '2']
  const family = m[1]
  const n = Number(m[2])
  // Seven colours per group: dark 1 (black) then accents 1–6.
  const colorOf = (i: number) => (i % 7 === 0 ? theme[1] : theme[3 + (i % 7)])
  const k = (n - 1) % 7
  const base = colorOf(k)
  const c = '#' + base
  const thin = (color: string): BorderSide => ({ style: 'thin', color })
  const white = '#FFFFFF'
  if (family === 'Light') {
    const group = Math.floor((n - 1) / 7)
    if (group === 0)
      return { header: { bold: true, color: k ? tint(base, -0.25) : undefined }, band: { fill: tint(base, 0.8) }, body: { color: k ? tint(base, -0.25) : undefined }, total: { bold: true }, outline: thin(c), headerLine: thin(c) }
    if (group === 1) return { header: { bold: true, fill: c, color: white }, total: { bold: true }, outline: thin(c), hLine: thin(c) }
    return { header: { bold: true }, band: { fill: tint(base, 0.8) }, total: { bold: true }, outline: thin(c), hLine: thin(c), vLine: thin(c), headerLine: thin(c) }
  }
  if (family === 'Medium') {
    const group = Math.floor((n - 1) / 7)
    if (group === 0) return { header: { bold: true, fill: c, color: white }, band: { fill: tint(base, 0.8) }, total: { bold: true }, hLine: thin(tint(base, 0.4)), outline: thin(tint(base, 0.4)) }
    if (group === 1) return { header: { bold: true, fill: c, color: white }, body: { fill: tint(base, 0.8) }, band: { fill: tint(base, 0.6) }, total: { bold: true, fill: c, color: white }, hLine: thin(white), vLine: thin(white) }
    if (group === 2) return { header: { bold: true, fill: k ? tint(base, -0.25) : '#000000', color: white }, band: { fill: tint(base, 0.6) }, body: { fill: tint(base, 0.8) }, total: { bold: true }, outline: thin('#000000') }
    return { header: { bold: true }, body: { fill: tint(base, 0.8) }, band: { fill: tint(base, 0.6) }, total: { bold: true }, hLine: thin(tint(base, 0.4)), outline: thin(tint(base, 0.4)) }
  }
  // Dark
  if (n <= 7) return { header: { bold: true, fill: '#000000', color: white }, body: { fill: tint(base, k ? -0.25 : 0.25), color: white }, band: { fill: tint(base, k ? -0.5 : 0) }, total: { bold: true, fill: tint(base, -0.5), color: white } }
  const pair = [
    [0, 0],
    [1, 2],
    [3, 4],
    [5, 6]
  ][n - 8] ?? [1, 2]
  return { header: { bold: true, fill: '#' + colorOf(pair[1] || 0), color: white }, body: { fill: tint(colorOf(pair[0]), 0.8) }, band: { fill: tint(colorOf(pair[0]), 0.6) }, total: { bold: true } }
}

const cache = new Map<string, Spec>()

/** The table style's formatting for one cell of the table (cell formatting goes on top of it). */
export function tableCellStyle(t: Table, r: number, c: number, theme: string[] = OFFICE_THEME): Style | null {
  if (t.style === undefined || t.style === '' || t.style === 'None') return null
  const key = `${t.style}|${theme.join()}`
  let s = cache.get(key)
  if (!s) cache.set(key, (s = spec(t.style, theme)))
  const { r0, r1, c0, c1 } = t.area
  const header = t.headerRow && r === r0
  const total = t.totalsRow && r === r1
  const first = t.headerRow ? r0 + 1 : r0
  const out: Style = { ...(header ? s.header : total ? s.total : s.body) }
  if (!header && !total) {
    if (t.showRowStripes !== false && s.band && (r - first) % 2 === 0) Object.assign(out, s.band)
    if (t.showColumnStripes && s.band && (c - c0) % 2 === 0) Object.assign(out, s.band)
    if ((t.showFirstColumn && c === c0) || (t.showLastColumn && c === c1)) out.bold = true
  }
  if (s.outline) {
    if (r === r0) out.borderTop = s.outline
    if (r === r1) out.borderBottom = s.outline
    if (c === c0) out.borderLeft = s.outline
    if (c === c1) out.borderRight = s.outline
  }
  if (s.hLine && r < r1 && !header) out.borderBottom = s.hLine
  if (s.vLine && c < c1) out.borderRight = s.vLine
  if (header && s.headerLine) out.borderBottom = s.headerLine
  if (total) out.borderTop = { style: 'double', color: (s.header.fill as string) ?? '#000000' }
  for (const k of Object.keys(out) as (keyof Style)[]) if (out[k] === undefined) delete out[k]
  return out
}

/** Built-in style names, for the gallery. */
export const TABLE_STYLES = [...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`), ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`), ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`)]
