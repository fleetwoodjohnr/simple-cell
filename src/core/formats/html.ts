// HTML tables: import every <table> as a sheet; export the sheet as a formatted table.
import { Workbook } from '../model/workbook'
import { areaName } from '../refs'
import { parseInput } from '../input'
import { formatValue } from '../numfmt/format'
import { esc } from './opc'
import { Warnings, type ExportResult, type ImportResult } from './types'
import { decodeText } from './csv'

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', euro: '€', pound: '£', copy: '©', reg: '®' }
const decode = (s: string) => s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e: string) => (e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : (ENT[e.toLowerCase()] ?? m)))
const cellText = (html: string) =>
  decode(
    html
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|li)>/gi, '\n')
      .replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')
      .replace(/<[^>]+>/g, '')
  )
    .replace(/[ \t\r]+/g, ' ')
    .replace(/ ?\n ?/g, '\n')
    .trim()

export function importHtml(bytes: Uint8Array): ImportResult {
  const w = new Warnings()
  const src = decodeText(bytes).replace(/<!--[\s\S]*?-->/g, '')
  const wb = new Workbook()
  const tables = [...src.matchAll(/<table\b[^>]*>([\s\S]*?)<\/table>/gi)]
  if (!tables.length) throw new Error('This page has no table to open.')
  for (const [i, t] of tables.entries()) {
    if (/<table\b/i.test(t[1])) w.add('info', 'Tables inside tables were opened as text.')
    const sh = wb.addSheet(tables.length > 1 ? `Table${i + 1}` : 'Sheet1')
    const taken = new Set<string>()
    let r = 0
    for (const tr of t[1].matchAll(/<tr\b[^>]*>([\s\S]*?)(?=<tr\b|<\/table>|$)/gi)) {
      let c = 0
      for (const td of tr[1].matchAll(/<(td|th)\b([^>]*)>([\s\S]*?)(?=<t[dh]\b|<\/tr>|$)/gi)) {
        while (taken.has(`${r},${c}`)) c++
        const attrs = td[2]
        const cs = Math.max(1, Number(/colspan\s*=\s*["']?(\d+)/i.exec(attrs)?.[1] ?? 1))
        const rs = Math.max(1, Number(/rowspan\s*=\s*["']?(\d+)/i.exec(attrs)?.[1] ?? 1))
        const text = cellText(td[3].replace(/<\/t[dh]>[\s\S]*$/i, ''))
        if (text) {
          const p = parseInput(text)
          const bold = td[1].toLowerCase() === 'th' || /<(b|strong)\b/i.test(td[3])
          let s: number | undefined
          if (p.format || bold) s = wb.styles.intern({ ...(p.format ? { numFmt: p.format } : {}), ...(bold ? { bold: true } : {}) })
          sh.set(r, c, { v: p.formula ? text : (p.value ?? text), ...(s ? { s } : {}) })
        }
        if (cs > 1 || rs > 1) sh.merges.push({ sheet: sh.id, r0: r, c0: c, r1: r + rs - 1, c1: c + cs - 1 })
        for (let dr = 0; dr < rs; dr++) for (let dc = 0; dc < cs; dc++) taken.add(`${r + dr},${c + dc}`)
        c += cs
      }
      r++
    }
  }
  return { workbook: wb, warnings: w.list }
}

export function exportHtml(wb: Workbook, opts: { sheet?: number } = {}): ExportResult {
  const w = new Warnings()
  const sheets = opts.sheet !== undefined ? [wb.sheetById(opts.sheet)!] : wb.sheets.filter((s) => !s.hidden)
  const body = sheets
    .map((sh) => {
      const u = sh.used()
      const covered = new Set<string>()
      const spans = new Map<string, string>()
      for (const m of sh.merges) {
        spans.set(`${m.r0},${m.c0}`, `${m.c1 > m.c0 ? ` colspan="${m.c1 - m.c0 + 1}"` : ''}${m.r1 > m.r0 ? ` rowspan="${m.r1 - m.r0 + 1}"` : ''}`)
        for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(`${r},${c}`)
      }
      const rows: string[] = []
      for (let r = 0; r <= u.r; r++) {
        if (sh.rowInfo.get(r)?.hidden) continue
        const tds: string[] = []
        for (let c = 0; c <= u.c; c++) {
          if (covered.has(`${r},${c}`) || sh.cols.get(c)?.hidden) continue
          const cell = sh.get(r, c)
          const st = wb.styles.get(cell?.s)
          const f = cell ? formatValue(cell.v, st.numFmt, wb.date1904, 255) : { text: '' }
          const css: string[] = []
          if (st.bold) css.push('font-weight:bold')
          if (st.italic) css.push('font-style:italic')
          if (st.underline) css.push('text-decoration:underline')
          if (st.strike) css.push('text-decoration:line-through')
          if (f.color ?? st.color) css.push(`color:${f.color ?? st.color}`)
          if (st.fill) css.push(`background:${st.fill}`)
          if (st.font) css.push(`font-family:'${st.font.replace(/'/g, '')}'`)
          if (st.size) css.push(`font-size:${st.size}pt`)
          const align = st.hAlign && st.hAlign !== 'general' ? st.hAlign : typeof cell?.v === 'number' ? 'right' : null
          if (align) css.push(`text-align:${align === 'centerContinuous' ? 'center' : align}`)
          if (st.wrap) css.push('white-space:pre-wrap')
          for (const [side, b] of [['top', st.borderTop], ['bottom', st.borderBottom], ['left', st.borderLeft], ['right', st.borderRight]] as const) if (b) css.push(`border-${side}:${b.style === 'thick' ? 3 : b.style === 'medium' ? 2 : 1}px ${b.style === 'double' ? 'double' : /dash/.test(b.style) ? 'dashed' : /dot/.test(b.style) ? 'dotted' : 'solid'} ${b.color ?? '#000'}`)
          tds.push(`<td${spans.get(`${r},${c}`) ?? ''}${css.length ? ` style="${esc(css.join(';'))}"` : ''}>${esc(f.text).replace(/\n/g, '<br>')}</td>`)
        }
        rows.push(`<tr>${tds.join('')}</tr>`)
      }
      return `<h2>${esc(sh.name)}</h2>\n<table data-range="${areaName({ r0: 0, c0: 0, r1: Math.max(0, u.r), c1: Math.max(0, u.c) })}">\n${rows.join('\n')}\n</table>`
    })
    .join('\n')
  w.add('loss', 'A web page keeps what cells show: formulas are saved as their results.')
  const html = `<!doctype html>\n<html><head><meta charset="utf-8"><title>${esc(sheets[0]?.name ?? 'Sheet')}</title><style>body{font-family:Calibri,Carlito,Arial,sans-serif;font-size:11pt}table{border-collapse:collapse}td{padding:2px 6px;vertical-align:bottom}</style></head><body>\n${body}\n</body></html>\n`
  return { data: new TextEncoder().encode(html), warnings: w.list }
}
