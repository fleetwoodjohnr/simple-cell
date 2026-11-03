// JSON: an array of objects (first row = keys) or an array of arrays, in and out.
import { Workbook, type Sheet } from '../model/workbook'
import { CellError, type Scalar } from '../engine/values'
import { formatValue, isDateFormat } from '../numfmt/format'
import { serialToDate, timeParts } from '../numfmt/dates'
import { parseInput } from '../input'
import { Warnings, type ExportResult, type ImportResult } from './types'
import { decodeText } from './csv'

function put(sh: Sheet, wb: Workbook, r: number, c: number, v: unknown, styleFor: Map<string, number>) {
  if (v === null || v === undefined) return
  if (typeof v === 'number' || typeof v === 'boolean') return sh.set(r, c, { v })
  if (typeof v === 'string') {
    // ISO dates and times become date cells; other text stays text.
    if (/^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/.test(v)) {
      const p = parseInput(v.replace('T', ' ').replace(/Z|[+-]\d{2}:?\d{2}$/, '').replace(/\.\d+$/, ''))
      if (typeof p.value === 'number') {
        const code = v.length > 10 ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd'
        let s = styleFor.get(code)
        if (s === undefined) styleFor.set(code, (s = wb.styles.intern({ numFmt: code })))
        return sh.set(r, c, { v: p.value, s })
      }
    }
    return sh.set(r, c, { v })
  }
  sh.set(r, c, { v: JSON.stringify(v) })
}

export function importJson(bytes: Uint8Array): ImportResult {
  const w = new Warnings()
  let data: unknown
  try {
    data = JSON.parse(decodeText(bytes))
  } catch (e) {
    throw new Error(`This file isn’t valid JSON: ${(e as Error).message}`)
  }
  const wb = new Workbook()
  const sh = wb.addSheet('Sheet1')
  const styleFor = new Map<string, number>()
  // Find the records: the value itself, or the first array inside an object ({"data": [...]})
  let rows: unknown = data
  if (!Array.isArray(rows) && rows && typeof rows === 'object') {
    const arr = Object.values(rows as Record<string, unknown>).find(Array.isArray)
    rows = arr ?? [rows]
  }
  if (!Array.isArray(rows)) rows = [[rows]]
  const list = rows as unknown[]
  if (list.every((x) => Array.isArray(x))) {
    list.forEach((row, r) => (row as unknown[]).forEach((v, c) => put(sh, wb, r, c, v, styleFor)))
  } else {
    const keys: string[] = []
    for (const o of list) if (o && typeof o === 'object' && !Array.isArray(o)) for (const k of Object.keys(o)) if (!keys.includes(k)) keys.push(k)
    keys.forEach((k, c) => sh.set(0, c, { v: k }))
    list.forEach((o, i) => {
      if (o && typeof o === 'object' && !Array.isArray(o)) keys.forEach((k, c) => put(sh, wb, i + 1, c, (o as Record<string, unknown>)[k], styleFor))
      else put(sh, wb, i + 1, 0, o, styleFor)
    })
    if (list.some((o) => o && typeof o === 'object' && Object.values(o as object).some((v) => v && typeof v === 'object'))) w.add('info', 'Nested objects and arrays were written into cells as JSON text.')
  }
  return { workbook: wb, warnings: w.list }
}

function jsonValue(v: Scalar, fmt: string | undefined, date1904: boolean): unknown {
  if (v instanceof CellError) return v.code
  if (typeof v === 'number' && fmt && isDateFormat(fmt)) {
    const d = serialToDate(v, date1904)
    const t = timeParts(v)
    const date = `${String(d.y).padStart(4, '0')}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`
    return v % 1 ? `${date}T${String(t.h).padStart(2, '0')}:${String(t.mi).padStart(2, '0')}:${String(t.s).padStart(2, '0')}` : date
  }
  return v
}

export function exportJson(wb: Workbook, opts: { sheet?: number } = {}): ExportResult {
  const w = new Warnings()
  const sh = opts.sheet !== undefined ? wb.sheetById(opts.sheet)! : wb.sheets[wb.activeSheet]
  if (wb.sheets.length > 1) w.add('loss', `Only the sheet “${sh.name}” is saved; JSON holds one table.`)
  const u = sh.used()
  const cell = (r: number, c: number) => {
    const x = sh.get(r, c)
    return x ? jsonValue(x.v, wb.styles.get(x.s).numFmt, wb.date1904) : null
  }
  const header: unknown[] = []
  for (let c = 0; c <= u.c; c++) header.push(cell(0, c))
  const useKeys = u.r >= 0 && header.every((h) => typeof h === 'string' && h !== '') && new Set(header).size === header.length
  const out: unknown[] = []
  for (let r = useKeys ? 1 : 0; r <= u.r; r++) {
    if (useKeys) {
      const o: Record<string, unknown> = {}
      header.forEach((h, c) => (o[h as string] = cell(r, c)))
      out.push(o)
    } else {
      const row: unknown[] = []
      for (let c = 0; c <= u.c; c++) row.push(cell(r, c))
      out.push(row)
    }
  }
  void formatValue
  w.add('loss', 'JSON keeps values only: formulas are saved as their results and formatting is not kept.')
  return { data: new TextEncoder().encode(JSON.stringify(out, null, 2) + '\n'), warnings: w.list }
}
