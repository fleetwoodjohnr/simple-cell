// OpenDocument Spreadsheet (.ods) writer.
import type { Cell, Sheet, Workbook } from '../../model/workbook'
import type { BorderSide, Style } from '../../model/styles'
import { printFormula } from '../../engine/parser'
import { CellError, numberToText } from '../../engine/values'
import { serialToDate, timeParts } from '../../numfmt/dates'
import { isDateFormat } from '../../numfmt/format'
import { bytes, text, writeZip } from '../zip'
import { esc } from '../opc'
import { Warnings, type ExportResult } from '../types'
import { excelToOdf } from './formula'
import type { OdsCarry } from './read'

const NS: Record<string, string> = {
  'xmlns:office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  'xmlns:style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  'xmlns:text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  'xmlns:table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  'xmlns:draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  'xmlns:fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  'xmlns:xlink': 'http://www.w3.org/1999/xlink',
  'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
  'xmlns:meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  'xmlns:number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  'xmlns:svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  'xmlns:chart': 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  'xmlns:of': 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
  'xmlns:config': 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
  'xmlns:calcext': 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
  'xmlns:loext': 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0'
}

const px = (v: number) => `${(v / 96).toFixed(4)}in`

// ---------- number formats → data styles ----------

/** ODF data style for an Excel format code, or null if it can't be expressed (then the cell keeps General). */
function dataStyleXml(name: string, code: string): string | null {
  const sections = code.split(';')
  const one = (sec: string, nm: string): string | null => {
    let s = sec.replace(/\[(Red|Black|Blue|Green|Color\d+|\$[^\]]*)\]/gi, (m) => (m.startsWith('[$') ? m.slice(2, -1).split('-')[0] : ''))
    const red = /\[Red\]/i.test(sec)
    const textProps = red ? '<style:text-properties fo:color="#ff0000"/>' : ''
    if (s === '@') return `<number:text-style style:name="${nm}"><number:text-content/></number:text-style>`
    if (isDateFormat(s)) {
      let out = ''
      const parts = s.match(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss(?:\.0+)?|s|AM\/PM|\[h+\]|"[^"]*"|\\.|./gi) ?? []
      const timeish = /h|s/i.test(s)
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i]
        const low = p.toLowerCase()
        const prevH = parts.slice(0, i).reverse().find((x) => /^[a-z[]/i.test(x))?.toLowerCase()
        const nextS = parts.slice(i + 1).find((x) => /^[a-z[]/i.test(x))?.toLowerCase()
        const minute = (low === 'm' || low === 'mm') && (prevH?.startsWith('h') || prevH?.startsWith('[h') || nextS?.startsWith('s'))
        if (low === 'yyyy') out += '<number:year number:style="long"/>'
        else if (low === 'yy') out += '<number:year/>'
        else if (low === 'mmmm') out += '<number:month number:style="long" number:textual="true"/>'
        else if (low === 'mmm') out += '<number:month number:textual="true"/>'
        else if (minute) out += low === 'mm' ? '<number:minutes number:style="long"/>' : '<number:minutes/>'
        else if (low === 'mm') out += '<number:month number:style="long"/>'
        else if (low === 'm') out += '<number:month/>'
        else if (low === 'dddd') out += '<number:day-of-week number:style="long"/>'
        else if (low === 'ddd') out += '<number:day-of-week/>'
        else if (low === 'dd') out += '<number:day number:style="long"/>'
        else if (low === 'd') out += '<number:day/>'
        else if (low === 'hh' || /^\[h+\]$/.test(low)) out += '<number:hours number:style="long"/>'
        else if (low === 'h') out += '<number:hours/>'
        else if (low.startsWith('ss')) out += `<number:seconds number:style="long"${low.includes('.') ? ` number:decimal-places="${low.length - 3}"` : ''}/>`
        else if (low === 's') out += '<number:seconds/>'
        else if (low === 'am/pm') out += '<number:am-pm/>'
        else out += `<number:text>${esc(p.startsWith('"') ? p.slice(1, -1) : p.startsWith('\\') ? p.slice(1) : p)}</number:text>`
      }
      const elapsed = /\[h/i.test(s) ? ' number:truncate-on-overflow="false"' : ''
      const tag = timeish && !/[yd]/i.test(s.replace(/"[^"]*"/g, '')) && !/m{3,}/i.test(s) ? 'time-style' : 'date-style'
      return `<number:${tag} style:name="${nm}"${elapsed}>${textProps}${out}</number:${tag}>`
    }
    const pct = s.includes('%')
    const sci = /E[+-]/i.test(s)
    const frac = /\?\/|#\/\d|\d+\/\d/.test(s) && /\//.test(s)
    // literal prefix/suffix text
    const lit = (t: string) => t.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_./g, ' ').replace(/\*./g, '')
    const m = /^(.*?)([#0?,]*\.?[#0?]*(?:E[+-]0+)?|# \?+\/[?\d]+)(.*)$/i.exec(s)
    if (!m) return null
    const [, pre, num, post] = m
    const decPart = /\.([0#?]+)/.exec(num)?.[1] ?? ''
    const intPart = num.split('.')[0]
    const grouping = intPart.includes(',')
    const minInt = (intPart.match(/0/g) ?? []).length
    const minDec = (decPart.match(/0/g) ?? []).length
    const dec = decPart.length
    let body: string
    if (frac) {
      const fm = /# (\?+)\/([?\d]+)/.exec(num)
      const den = fm && /^\d+$/.test(fm[2]) ? ` number:denominator-value="${fm[2]}"` : ''
      body = `<number:fraction number:min-integer-digits="0" number:min-numerator-digits="${fm?.[1].length ?? 1}" number:min-denominator-digits="${fm?.[2].length ?? 1}"${den}/>`
    } else if (sci) {
      const ed = (/E[+-](0+)/i.exec(num)?.[1] ?? '00').length
      body = `<number:scientific-number number:decimal-places="${dec}" number:min-decimal-places="${minDec}" number:min-integer-digits="${Math.max(minInt, 1)}" number:min-exponent-digits="${ed}"/>`
    } else body = `<number:number number:decimal-places="${dec}" number:min-decimal-places="${minDec}" number:min-integer-digits="${minInt}"${grouping ? ' number:grouping="true"' : ''}/>`
    const preT = lit(pre).replace('%', '')
    const postT = lit(post).replace('%', '')
    const currency = /[$€£¥]/.exec(preT + postT)?.[0]
    const tag = pct ? 'percentage-style' : currency ? 'currency-style' : 'number-style'
    const pieces = [preT ? (currency && preT.includes(currency) ? `<number:currency-symbol>${esc(currency)}</number:currency-symbol>${preT.replace(currency, '') ? `<number:text>${esc(preT.replace(currency, ''))}</number:text>` : ''}` : `<number:text>${esc(preT)}</number:text>`) : '', body, postT ? `<number:text>${esc(postT)}</number:text>` : '', pct ? '<number:text>%</number:text>' : '']
    return `<number:${tag} style:name="${nm}">${textProps}${pieces.join('')}</number:${tag}>`
  }
  if (sections.length === 1) return one(sections[0], name)
  // Positive;negative[;zero]: the negative section is applied through a map.
  const pos = one(sections[0], name + 'P0')
  const neg = one(sections[1], name)
  if (!pos || !neg) return null
  return pos + neg.replace(/(<number:[\w-]+ style:name="[^"]*"[^>]*>)/, `$1<style:map style:condition="value()&gt;=0" style:apply-style-name="${name}P0"/>`).replace(/<number:text>-<\/number:text>/, '')
}

// ---------- cell styles ----------

function borderAttr(b: BorderSide | undefined): string {
  if (!b) return 'none'
  const w = b.style === 'thick' ? '2.5pt' : b.style === 'medium' ? '1.5pt' : b.style === 'hair' ? '0.25pt' : '0.75pt'
  const kind = b.style === 'double' ? 'double' : /dash/i.test(b.style) ? 'dashed' : /dot/i.test(b.style) ? 'dotted' : 'solid'
  return `${b.style === 'double' ? '1.1pt' : w} ${kind} ${b.color ?? '#000000'}`
}

function cellStyleXml(name: string, s: Style, dataStyle: string | null): string {
  const cell: string[] = []
  if (s.fill) cell.push(`fo:background-color="${s.fill}"`)
  for (const [k, side] of [['border-top', s.borderTop], ['border-bottom', s.borderBottom], ['border-left', s.borderLeft], ['border-right', s.borderRight]] as const) if (side) cell.push(`fo:${k}="${borderAttr(side)}"`)
  if (s.wrap) cell.push('fo:wrap-option="wrap"')
  if (s.vAlign === 'top' || s.vAlign === 'center') cell.push(`style:vertical-align="${s.vAlign === 'center' ? 'middle' : 'top'}"`)
  if (s.rotation) cell.push(`style:rotation-angle="${s.rotation < 0 ? 360 + s.rotation : s.rotation}"`)
  if (s.shrink) cell.push('style:shrink-to-fit="true"')
  if (s.locked === false) cell.push('style:cell-protect="none"')
  if (s.hAlign && s.hAlign !== 'general') cell.push('style:text-align-source="fix"')
  const para: string[] = []
  if (s.hAlign === 'center' || s.hAlign === 'centerContinuous') para.push('fo:text-align="center"')
  else if (s.hAlign === 'right') para.push('fo:text-align="end"')
  else if (s.hAlign === 'left') para.push('fo:text-align="start"')
  else if (s.hAlign === 'justify') para.push('fo:text-align="justify"')
  if (s.indent) para.push(`fo:margin-left="${(s.indent * 9 * 0.75).toFixed(1)}pt"`)
  const txt: string[] = []
  if (s.bold) txt.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"')
  if (s.italic) txt.push('fo:font-style="italic"')
  if (s.underline) txt.push(`style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"${s.underline.startsWith('double') ? ' style:text-underline-type="double"' : ''}`)
  if (s.strike) txt.push('style:text-line-through-style="solid"')
  if (s.color) txt.push(`fo:color="${s.color}"`)
  if (s.size) txt.push(`fo:font-size="${s.size}pt"`)
  if (s.font) txt.push(`style:font-name="${esc(s.font)}"`)
  if (s.vertAlign) txt.push(`style:text-position="${s.vertAlign === 'superscript' ? 'super 58%' : 'sub 58%'}"`)
  return `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${dataStyle ? ` style:data-style-name="${dataStyle}"` : ''}>${cell.length ? `<style:table-cell-properties ${cell.join(' ')}/>` : ''}${para.length ? `<style:paragraph-properties ${para.join(' ')}/>` : ''}${txt.length ? `<style:text-properties ${txt.join(' ')}/>` : ''}</style:style>`
}

// ---------- writing ----------

export function writeOds(wb: Workbook): ExportResult {
  const w = new Warnings()
  const carry = (wb.carry as OdsCarry | null)?.kind === 'ods' ? (wb.carry as OdsCarry) : undefined
  const files: Record<string, Uint8Array> = carry ? { ...carry.files } : {}
  const sheetName = (id: number) => wb.sheetById(id)?.name ?? null
  // Styles used anywhere
  const styleNames = new Map<number, string>()
  const dataNames = new Map<string, string>()
  const autoStyles: string[] = []
  const dataStyles: string[] = []
  const cellStyle = (id: number | undefined): string | null => {
    if (!id) return null
    let n = styleNames.get(id)
    if (n) return n
    n = `sc${styleNames.size + 1}`
    styleNames.set(id, n)
    const s = wb.styles.get(id)
    let ds: string | null = null
    if (s.numFmt) {
      ds = dataNames.get(s.numFmt) ?? null
      if (!ds) {
        const dn = `scN${dataNames.size + 1}`
        const x = dataStyleXml(dn, s.numFmt)
        if (x) {
          dataStyles.push(x)
          dataNames.set(s.numFmt, dn)
          ds = dn
        } else w.add('loss', `The number format “${s.numFmt}” can’t be written in OpenDocument; those cells use General.`)
      }
    }
    autoStyles.push(cellStyleXml(n, s, ds))
    return n
  }
  const colStyles = new Map<number, string>()
  const colStyle = (width: number) => {
    let n = colStyles.get(width)
    if (!n) {
      n = `scco${colStyles.size + 1}`
      colStyles.set(width, n)
      autoStyles.push(`<style:style style:name="${n}" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="${px(width)}"/></style:style>`)
    }
    return n
  }
  const rowStyles = new Map<number, string>()
  const rowStyle = (h: number | null) => {
    const key = h ?? -1
    let n = rowStyles.get(key)
    if (!n) {
      n = `scro${rowStyles.size + 1}`
      rowStyles.set(key, n)
      autoStyles.push(`<style:style style:name="${n}" style:family="table-row"><style:table-row-properties style:row-height="${px(h ?? 20)}" fo:break-before="auto" style:use-optimal-row-height="${h === null ? 'true' : 'false'}"/></style:style>`)
    }
    return n
  }
  const tableStyles: string[] = []
  const tables = wb.sheets.map((sh, i) => {
    const ts = `scta${i + 1}`
    tableStyles.push(`<style:style style:name="${ts}" style:family="table" style:master-page-name="Default"><style:table-properties table:display="${sh.hidden ? 'false' : 'true'}" style:writing-mode="lr-tb"/></style:style>`)
    return tableXml(sh, wb, ts, carry, { cellStyle, colStyle, rowStyle, sheetName, w })
  })
  const before = (carry?.spreadsheetElements ?? []).filter((x) => /^<table:(content-validations|label-ranges|tracked-changes)\b/.test(x))
  const after = (carry?.spreadsheetElements ?? []).filter((x) => !/^<table:(content-validations|label-ranges|tracked-changes|named-expressions|calculation-settings)\b/.test(x))
  const names = wb.names
    .filter((n) => n.scope === null && !/^_xlnm\./i.test(n.name))
    .map((n) => {
      const f = n.ast ? printFormula(n.ast, { row: 0, col: 0, sheetName }) : n.text
      const odf = excelToOdf('=' + f)
      const single = /^of:=\[(.*)\]$/.exec(odf)
      return single ? `<table:named-range table:name="${esc(n.name)}" table:base-cell-address="${esc(wb.sheets[0] ? `$'${wb.sheets[0].name.replace(/'/g, "''")}'.$A$1` : '$A$1')}" table:cell-range-address="${esc(single[1])}"/>` : `<table:named-expression table:name="${esc(n.name)}" table:base-cell-address="${esc(`$'${(wb.sheets[0]?.name ?? 'Sheet1').replace(/'/g, "''")}'.$A$1`)}" table:expression="${esc(odf)}"/>`
    })
  const calc = `<table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true">${wb.date1904 ? '<table:null-date table:date-value="1904-01-01"/>' : ''}${wb.calc.iterate ? `<table:iteration table:status="enable" table:steps="${wb.calc.maxIterations}" table:minimum-difference="${wb.calc.maxChange}"/>` : ''}</table:calculation-settings>`
  const ns = { ...NS, ...(carry?.rootNs ?? {}) }
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${Object.entries(ns).map(([k, v]) => `${k}="${esc(v)}"`).join(' ')} office:version="1.3">` +
    '<office:scripts/>' +
    `<office:automatic-styles>${tableStyles.join('')}${dataStyles.join('')}${autoStyles.join('')}${(carry?.autoStyles ?? []).join('')}</office:automatic-styles>` +
    `<office:body><office:spreadsheet>${calc}${before.join('')}${tables.join('')}${names.length ? `<table:named-expressions>${names.join('')}</table:named-expressions>` : ''}${after.join('')}</office:spreadsheet></office:body></office:document-content>`
  files['content.xml'] = bytes(content)
  files.mimetype = bytes('application/vnd.oasis.opendocument.spreadsheet')
  if (!files['styles.xml'])
    files['styles.xml'] = bytes(
      `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${Object.entries(NS).map(([k, v]) => `${k}="${v}"`).join(' ')} office:version="1.3"><office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri" style:font-family-generic="swiss"/></office:font-face-decls><office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:default-style><style:style style:name="Default" style:family="table-cell"/></office:styles><office:automatic-styles><style:page-layout style:name="pm1"/></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>`
    )
  files['settings.xml'] = bytes(settingsXml(wb))
  const now = new Date().toISOString().replace(/\.\d+Z$/, '')
  files['meta.xml'] = files['meta.xml'] ? bytes(text(files['meta.xml']).replace(/(<dc:date>)[^<]*(<\/dc:date>)/, `$1${now}$2`).replace(/<meta:generator>[^<]*<\/meta:generator>/, '<meta:generator>Simple Cell</meta:generator>')) : bytes(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta xmlns:office="${NS['xmlns:office']}" xmlns:meta="${NS['xmlns:meta']}" xmlns:dc="${NS['xmlns:dc']}" office:version="1.3"><office:meta><meta:generator>Simple Cell</meta:generator><dc:date>${now}</dc:date></office:meta></office:document-meta>`)
  files['META-INF/manifest.xml'] = bytes(manifestXml(files))
  return { data: writeZip(files), warnings: w.list }
}

interface TableWriteCtx {
  cellStyle: (id: number | undefined) => string | null
  colStyle: (w: number) => string
  rowStyle: (h: number | null) => string
  sheetName: (id: number) => string | null
  w: Warnings
}

function tableXml(sh: Sheet, wb: Workbook, ts: string, carry: OdsCarry | undefined, ctx: TableWriteCtx): string {
  const u = sh.used()
  const lastCol = Math.max(u.c, ...[...sh.cols.keys()].filter((c) => c < 1024), 0)
  const cols: string[] = []
  for (let c = 0; c <= lastCol; ) {
    const info = sh.cols.get(c)
    const width = info?.width ?? sh.defaultColWidth
    const hidden = info?.hidden
    let n = 1
    while (c + n <= lastCol && JSON.stringify(sh.cols.get(c + n) ?? {}) === JSON.stringify(info ?? {})) n++
    const def = info?.style ? ctx.cellStyle(info.style) : null
    cols.push(`<table:table-column table:style-name="${ctx.colStyle(width)}"${n > 1 ? ` table:number-columns-repeated="${n}"` : ''}${hidden ? ' table:visibility="collapse"' : ''} table:default-cell-style-name="${def ?? 'Default'}"/>`)
    c += n
  }
  const mergeAt = new Map<string, { r1: number; c1: number }>()
  const covered = new Set<string>()
  for (const m of sh.merges) {
    mergeAt.set(`${m.r0},${m.c0}`, { r1: m.r1, c1: m.c1 })
    for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(`${r},${c}`)
  }
  const rows: string[] = []
  const lastRow = Math.max(u.r, ...[...sh.rowInfo.keys()].filter((r) => r < 100000), 0)
  let pendingEmpty = 0
  const flushEmpty = () => {
    if (pendingEmpty) rows.push(`<table:table-row table:style-name="${ctx.rowStyle(null)}"${pendingEmpty > 1 ? ` table:number-rows-repeated="${pendingEmpty}"` : ''}><table:table-cell${lastCol > 0 ? ` table:number-columns-repeated="${lastCol + 1}"` : ''}/></table:table-row>`)
    pendingEmpty = 0
  }
  for (let r = 0; r <= lastRow; r++) {
    const row = sh.rows.get(r)
    const ri = sh.rowInfo.get(r)
    if (!row?.size && !ri && !sh.merges.some((m) => m.r0 <= r && m.r1 >= r)) {
      pendingEmpty++
      continue
    }
    flushEmpty()
    const cells: string[] = []
    let blank = 0
    const flushBlank = () => {
      if (blank) cells.push(`<table:table-cell${blank > 1 ? ` table:number-columns-repeated="${blank}"` : ''}/>`)
      blank = 0
    }
    const rowLast = Math.max(row ? Math.max(...row.keys()) : -1, ...sh.merges.filter((m) => m.r0 <= r && m.r1 >= r).map((m) => m.c1))
    for (let c = 0; c <= rowLast; c++) {
      const cell = row?.get(c)
      if (covered.has(`${r},${c}`)) {
        flushBlank()
        cells.push(`<table:covered-table-cell${cell?.s ? ` table:style-name="${ctx.cellStyle(cell.s)}"` : ''}/>`)
        continue
      }
      if (!cell) {
        blank++
        continue
      }
      flushBlank()
      cells.push(cellXml(cell, r, c, wb, mergeAt.get(`${r},${c}`), ctx))
    }
    flushBlank()
    const h = ri?.hidden ? null : (ri?.height ?? null)
    rows.push(`<table:table-row table:style-name="${ctx.rowStyle(h)}"${ri?.hidden ? ` table:visibility="${ri.filtered ? 'filter' : 'collapse'}"` : ''}>${cells.join('') || '<table:table-cell/>'}</table:table-row>`)
  }
  pendingEmpty = 0
  const kept = carry?.sheets[sh.id]
  const attrs = Object.entries(kept?.attrs ?? {})
    .filter(([k]) => k !== 'table:print' || true)
    .map(([k, v]) => ` ${k}="${esc(v)}"`)
    .join('')
  // Shapes, conditional formats and forms go before the columns; other kept elements after the rows.
  const head = (kept?.elements ?? []).filter((x) => /^<(table:shapes|calcext:conditional-formats|office:forms|table:title|table:desc|table:table-source|table:scenario)\b/.test(x)).join('')
  const tail = (kept?.elements ?? []).filter((x) => !/^<(table:shapes|calcext:conditional-formats|office:forms|table:title|table:desc|table:table-source|table:scenario)\b/.test(x)).join('')
  return `<table:table table:name="${esc(sh.name)}" table:style-name="${ts}"${attrs}>${head}${cols.join('') || '<table:table-column/>'}${rows.join('')}${tail}</table:table>`
}

function cellXml(cell: Cell, r: number, c: number, wb: Workbook, merge: { r1: number; c1: number } | undefined, ctx: TableWriteCtx): string {
  const a: string[] = []
  const style = ctx.cellStyle(cell.s)
  if (style) a.push(`table:style-name="${style}"`)
  if (merge) a.push(`table:number-columns-spanned="${merge.c1 - c + 1}" table:number-rows-spanned="${merge.r1 - r + 1}"`)
  if (cell.f) {
    const f = printFormula(cell.f.ast, { row: r, col: c, sheetName: ctx.sheetName })
    a.push(`table:formula="${esc(excelToOdf('=' + f))}"`)
    if (cell.f.array) a.push(`table:number-matrix-columns-spanned="${cell.f.array.c1 - cell.f.array.c0 + 1}" table:number-matrix-rows-spanned="${cell.f.array.r1 - cell.f.array.r0 + 1}"`)
  }
  const v = cell.v
  const fmt = wb.styles.get(cell.s).numFmt
  let shown = ''
  if (typeof v === 'number') {
    if (fmt && isDateFormat(fmt)) {
      const d = serialToDate(v, wb.date1904)
      const t = timeParts(v)
      if (Math.floor(v) === 0 && v % 1 && !/[yd]/i.test(fmt)) a.push(`office:value-type="time" office:time-value="PT${t.h}H${t.mi}M${t.s}${t.ms ? '.' + String(t.ms).padStart(3, '0') : ''}S"`)
      else a.push(`office:value-type="date" office:date-value="${String(d.y).padStart(4, '0')}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}${v % 1 ? `T${String(t.h).padStart(2, '0')}:${String(t.mi).padStart(2, '0')}:${String(t.s).padStart(2, '0')}` : ''}"`)
    } else if (fmt && /%/.test(fmt.split(';')[0])) a.push(`office:value-type="percentage" office:value="${v}"`)
    else a.push(`office:value-type="float" office:value="${v}"`)
    shown = numberToText(v)
  } else if (typeof v === 'boolean') {
    a.push(`office:value-type="boolean" office:boolean-value="${v}"`)
    shown = v ? 'TRUE' : 'FALSE'
  } else if (v instanceof CellError) {
    if (cell.f) a.push('office:value-type="string" office:string-value=""', 'calcext:value-type="error"')
    shown = v.code
  } else if (typeof v === 'string') {
    a.push('office:value-type="string"')
    shown = v
  }
  const extra = (cell.rich as { ods?: string } | undefined)?.ods ?? ''
  const paras = shown === '' && v === null ? '' : shown.split('\n').map((l) => `<text:p>${esc(l).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`)}</text:p>`).join('')
  return `<table:table-cell${a.length ? ' ' + a.join(' ') : ''}>${extra}${paras}</table:table-cell>`
}

function settingsXml(wb: Workbook): string {
  const tables = wb.sheets
    .map(
      (s) =>
        `<config:config-item-map-entry config:name="${esc(s.name)}"><config:config-item config:name="HorizontalSplitMode" config:type="short">${s.freeze.cols ? 2 : 0}</config:config-item><config:config-item config:name="VerticalSplitMode" config:type="short">${s.freeze.rows ? 2 : 0}</config:config-item><config:config-item config:name="HorizontalSplitPosition" config:type="int">${s.freeze.cols}</config:config-item><config:config-item config:name="VerticalSplitPosition" config:type="int">${s.freeze.rows}</config:config-item><config:config-item config:name="ActiveSplitRange" config:type="short">${s.freeze.rows ? 2 : s.freeze.cols ? 3 : 2}</config:config-item><config:config-item config:name="PositionLeft" config:type="int">0</config:config-item><config:config-item config:name="PositionRight" config:type="int">${s.freeze.cols}</config:config-item><config:config-item config:name="PositionTop" config:type="int">0</config:config-item><config:config-item config:name="PositionBottom" config:type="int">${s.freeze.rows}</config:config-item><config:config-item config:name="ShowGrid" config:type="boolean">${s.showGridlines}</config:config-item><config:config-item config:name="ZoomValue" config:type="int">${Math.round(s.zoom * 100)}</config:config-item></config:config-item-map-entry>`
    )
    .join('')
  return `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-settings xmlns:office="${NS['xmlns:office']}" xmlns:config="${NS['xmlns:config']}" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item config:name="ViewId" config:type="string">view1</config:config-item><config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named><config:config-item config:name="ActiveTable" config:type="string">${esc(wb.sheets[wb.activeSheet]?.name ?? '')}</config:config-item><config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`
}

function manifestXml(files: Record<string, Uint8Array>): string {
  const type = (p: string) =>
    p.endsWith('.xml') ? 'text/xml' : p.endsWith('.png') ? 'image/png' : p.endsWith('.jpg') || p.endsWith('.jpeg') ? 'image/jpeg' : p.endsWith('.svg') ? 'image/svg+xml' : p.endsWith('.gif') ? 'image/gif' : 'application/octet-stream'
  const entries = Object.keys(files)
    .filter((p) => p !== 'mimetype' && p !== 'META-INF/manifest.xml' && !p.endsWith('/'))
    .map((p) => {
      // Embedded objects (charts) are directories with their own media type.
      return `<manifest:file-entry manifest:full-path="${esc(p)}" manifest:media-type="${type(p)}"/>`
    })
  const dirs = new Set(Object.keys(files).filter((p) => /^Object \d+\//.test(p)).map((p) => p.split('/')[0]))
  const objects = [...dirs].map((d) => `<manifest:file-entry manifest:full-path="${esc(d)}/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>`)
  return `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>${objects.join('')}${entries.join('')}</manifest:manifest>`
}
