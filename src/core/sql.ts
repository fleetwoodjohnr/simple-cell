// SQL over workbook tables and optional selected data. Execution uses snapshots in a disposable worker;
// successful runs apply all source edits and output cells in one workbook transaction.
import type { SqlValue } from 'sql.js'
import type { Doc, Tx } from './doc'
import type { Table } from './model/workbook'
import { CellError, MAX_COLS, MAX_ROWS, type Area, type Scalar } from './engine/values'
import { cellName } from './refs'

/** Returned rows are ordinary cells on a new sheet or at a destination cell. */
export type SqlOut = { kind: 'sheet' } | { kind: 'cell'; sheet: number; r: number; c: number }

type Rows = { columns: string[]; rows: Scalar[][] }

/** A table's data rows (no header or total row). */
function dataArea(t: Table): Area {
  return { ...t.area, r0: t.headerRow ? t.area.r0 + 1 : t.area.r0, r1: t.totalsRow ? t.area.r1 - 1 : t.area.r1 }
}

function toSql(v: Scalar): SqlValue {
  if (v instanceof CellError) return v.code
  if (typeof v === 'boolean') return v ? 1 : 0
  return v === '' ? null : v
}

/** A cell keeps its content (and its formula) when the query gives back the value it already has. */
function same(old: Scalar, v: Scalar): boolean {
  if (old instanceof CellError) return typeof v === 'string' ? v === old.code : v === old
  if (typeof old === 'boolean') return v === (old ? 1 : 0) || v === old
  return old === v || (old === '' && v === null)
}

export type SqlRange = { area: Area; headers: boolean }

export function rangeSource(doc: Doc, source: SqlRange): Table {
  const a = source.area
  if (![a.r0, a.r1, a.c0, a.c1].every(Number.isInteger) || a.r0 < 0 || a.c0 < 0 || a.r1 < a.r0 || a.c1 < a.c0 || a.r1 >= MAX_ROWS || a.c1 >= MAX_COLS) throw new Error('Enter a valid source range.')
  doc.sheet(a.sheet)
  let name = 'Selection'
  for (let n = 2; doc.wb.tableByName(name); n++) name = `Selection${n}`
  const columns = uniqueNames(Array.from({ length: a.c1 - a.c0 + 1 }, (_, i) => source.headers ? String(doc.value(a.sheet, a.r0, a.c0 + i) ?? '') : `Column${i + 1}`)).map((name) => ({ name }))
  return { id: -1, name, displayName: name, sheet: a.sheet, area: { ...a }, headerRow: source.headers, totalsRow: false, columns, style: '', showRowStripes: false }
}

const pending = new WeakMap<Doc, () => void>()
export function cancelSql(doc: Doc) { pending.get(doc)?.() }

/** SQLite runs separately from the calculation worker, so cancellation never loses workbook state. */
async function exec(doc: Doc, sql: string, source?: SqlRange): Promise<{ result: Rows | null; changed: { table: Table; data: Rows }[] }> {
  if (pending.has(doc)) throw new Error('A SQL query is already running.')
  const version = doc.version
  const tables = [...doc.wb.tables, ...(source ? [rangeSource(doc, source)] : [])]
  let cells = 0
  const snapshot = tables.map((t) => {
    const a = dataArea(t), rows: SqlValue[][] = []
    cells += Math.max(0, a.r1 - a.r0 + 1) * t.columns.length
    if (cells > 1_000_000) throw new Error('SQL source data exceeds one million cells. Select a smaller range.')
    for (let r = a.r0; r <= a.r1; r++) rows.push(t.columns.map((_, i) => toSql(doc.value(t.sheet, r, a.c0 + i))))
    if (t.id !== -1 && rows.length === 1 && rows[0].every((v) => v === null)) rows.length = 0
    return { name: t.name, columns: t.columns.map((c) => c.name), rows }
  })
  const input = { sql, tables: snapshot }
  let output: Awaited<ReturnType<typeof import('./sql-engine').executeSql>>
  if (import.meta.env.MODE === 'test') output = await (await import('./sql-engine')).executeSql(input)
  else {
    const { default: SqlWorker } = await import('./sql-worker?worker')
    const worker = new SqlWorker()
    try {
      output = await new Promise((resolve, reject) => {
        const finish = (error: string) => reject(new Error(error))
        const timer = setTimeout(() => finish('The SQL query exceeded 30 seconds and was stopped.'), 30_000)
        worker.onmessage = (e) => { clearTimeout(timer); e.data.error ? finish(e.data.error) : resolve(e.data.result) }
        worker.onerror = (e) => { clearTimeout(timer); finish(e.message || 'SQL worker failed.') }
        // Clear the timer on cancellation as well.
        pending.set(doc, () => { clearTimeout(timer); finish('SQL query cancelled.') })
        worker.postMessage(input)
      })
    } finally { worker.terminate(); pending.delete(doc) }
  }
  if (doc.version !== version) throw new Error('The workbook changed while SQL was running. Run the query again.')
  const changed = output.changed.map((c) => ({ table: tables.find((t) => t.name === c.name)!, data: c.data }))
  for (const c of changed) if (c.table.id === -1 && doc.wb.tables.some((t) => overlaps(t.area, c.table.area))) throw new Error('This range overlaps a named table. Edit that table instead.')
  return { result: output.result, changed }
}

const overlaps = (a: Area, b: Area) => a.sheet === b.sheet && a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0

/** What a query would do, without changing the workbook: its first 100 rows, and the tables it would change. */
export async function previewSql(doc: Doc, sql: string, source?: SqlRange) {
  const { result, changed } = await exec(doc, sql, source)
  return {
    columns: result?.columns ?? null,
    rows: (result?.rows.slice(0, 100) ?? []).map((row) => row.map((v) => (v instanceof CellError ? v.code : v))),
    rowCount: result?.rows.length ?? 0,
    changed: changed.map((c) => ({ table: c.table.name, rows: c.data.rows.length }))
  }
}

/** Run a query and apply it as one undoable step: changed tables are rewritten, the query's rows go to `out`. */
export async function runSql(doc: Doc, sql: string, out: SqlOut, source?: SqlRange) {
  const { result, changed } = await exec(doc, sql, source)
  if (!result && !changed.length) throw new Error('The query returned no rows and changed no tables.')
  let sheet: number | undefined
  let table: string | undefined
  doc.transact('SQL query', (tx) => {
    for (const c of changed) writeTableTx(doc, tx, c.table, c.data)
    if (!result) return
    let at = out.kind === 'cell' ? out : null
    if (!at) {
      let name = 'Query'
      for (let k = 2; doc.wb.sheetByName(name); k++) name = `Query ${k}`
      const sh = doc.wb.createSheet(name)
      const idx = doc.wb.sheets.length
      tx.custom(
        () => doc.wb.sheets.splice(idx, 0, sh),
        () => doc.wb.sheets.splice(doc.wb.sheetIndex(sh.id), 1)
      )
      tx.structural = true
      at = { kind: 'cell', sheet: sh.id, r: 0, c: 0 }
    }
    const names = uniqueNames(result.columns)
    const area = { sheet: at.sheet, r0: at.r, c0: at.c, r1: at.r + result.rows.length, c1: at.c + names.length - 1 }
    if (source && overlaps(area, source.area)) throw new Error('The output overlaps the source range. Choose another destination.')
    ensureRoom(doc, area, null)
    const sh = doc.sheet(at.sheet)
    names.forEach((n, i) => tx.setCell(at.sheet, at.r, at.c + i, { v: n, s: sh.get(at.r, at.c + i)?.s }))
    // A column named like a table column takes that column's format (dates stay dates).
    const fmt = names.map((n) => sourceStyle(doc, n))
    result.rows.forEach((row, j) => row.forEach((v, i) => v !== null && tx.setCell(at.sheet, at.r + 1 + j, at.c + i, { v, s: sh.get(at.r + 1 + j, at.c + i)?.s ?? fmt[i] })))
    sheet = at.sheet
    table = doc.sheet(sheet).name
  })
  return { changed: changed.map((c) => c.table.name), rows: result?.rows.length ?? null, sheet, table }
}

/** The style of the first data cell of a table column with this name, if any. */
function sourceStyle(doc: Doc, name: string): number | undefined {
  for (const t of doc.wb.tables) {
    const i = t.columns.findIndex((c) => c.name.toLowerCase() === name.toLowerCase())
    if (i >= 0) return doc.sheet(t.sheet).get(dataArea(t).r0, t.area.c0 + i)?.s
  }
}

function uniqueNames(columns: string[]): string[] {
  const names: string[] = []
  columns.forEach((col, i) => {
    const base = col.trim() || `Column${i + 1}`
    let name = base
    for (let k = 2; names.some((x) => x.toLowerCase() === name.toLowerCase()); k++) name = `${base}${k}`
    names.push(name)
  })
  return names
}

/** Refuse to spread over data, other tables or merged cells (cells inside `own` are free). */
function ensureRoom(doc: Doc, a: Area, own: Table | null) {
  if (a.r0 < 0 || a.c0 < 0 || !Number.isInteger(a.r0) || !Number.isInteger(a.c0) || a.r1 >= MAX_ROWS || a.c1 >= MAX_COLS) throw new Error('The result doesn’t fit on the sheet.')
  const sh = doc.sheet(a.sheet)
  const mine = (r: number, c: number) => !!own && r >= own.area.r0 && r <= own.area.r1 && c >= own.area.c0 && c <= own.area.c1
  for (let r = a.r0; r <= a.r1; r++)
    for (let c = a.c0; c <= a.c1; c++) {
      if (mine(r, c)) continue
      const cell = sh.get(r, c)
      const other = doc.wb.tableAt(a.sheet, r, c)
      if ((cell && (cell.v !== null || cell.f)) || (other && other !== own)) throw new Error(`There’s already data in ${cellName(r, c)}${own ? `, next to ${own.name}` : ''}. Move it, or put the result somewhere else.`)
    }
  if (sh.merges.some((m) => m.r0 <= a.r1 && m.r1 >= a.r0 && m.c0 <= a.c1 && m.c1 >= a.c0 && !(own && mine(m.r0, m.c0)))) throw new Error('The result would cover merged cells. Unmerge them first.')
}

/**
 * Rewrite a table with new columns and rows, growing or shrinking it from its top-left corner. Cells that keep their
 * value are left alone, so their formulas and formats survive.
 * ponytail: a changed cell is written as a value (a calculated column's formula goes once its value changes).
 */
function writeTableTx(doc: Doc, tx: Tx, t: Table, data: Rows) {
  if (t.id !== -1 && !t.headerRow) throw new Error(`${t.name} has no header row. Turn it on in Table Design first.`)
  if (t.totalsRow) throw new Error(`Turn off ${t.name}’s total row first (Table Design ▸ Total row).`)
  const sh = doc.sheet(t.sheet)
  const old = { ...t.area }
  const names = uniqueNames(data.columns)
  const area = { sheet: t.sheet, r0: old.r0, c0: old.c0, r1: old.r0 + (t.headerRow ? 1 : 0) + Math.max(t.id === -1 ? 0 : 1, data.rows.length) - 1, c1: old.c0 + names.length - 1 }
  if (t.id === -1 && sh.merges.some((m) => overlaps({ ...m, sheet: t.sheet }, old))) throw new Error('Unmerge the source cells before editing them with SQL.')
  ensureRoom(doc, area, t)
  const rewritten = new Set<number>()
  const put = (r: number, c: number, v: Scalar) => {
    const cell = sh.get(r, c)
    if (same(cell?.v ?? null, v)) return
    tx.setCell(t.sheet, r, c, v === null && !cell?.s ? undefined : { v, s: cell?.s })
    if (r > area.r0) rewritten.add(c - area.c0)
  }
  if (t.headerRow) names.forEach((n, i) => put(area.r0, area.c0 + i, n))
  const first = area.r0 + (t.headerRow ? 1 : 0)
  for (let r = first; r <= area.r1; r++) for (let i = 0; i < names.length; i++) put(r, area.c0 + i, data.rows[r - first]?.[i] ?? null)
  for (let r = old.r0; r <= old.r1; r++)
    for (let c = old.c0; c <= old.c1; c++) if ((r > area.r1 || c > area.c1) && sh.get(r, c)) tx.setCell(t.sheet, r, c, undefined)
  if (t.id === -1) { tx.refetch = true; return }
  const byName = new Map(t.columns.map((c) => [c.name.toLowerCase(), c]))
  tx.prop(
    t,
    'columns',
    names.map((name, i) => {
      const { formula, ...col } = byName.get(name.toLowerCase()) ?? { name }
      return rewritten.has(i) || !formula ? { ...col, name } : { ...col, name, formula }
    })
  )
  tx.prop(t, 'area', area)
  if (t.filter?.length) {
    tx.prop(t, 'filter', [])
    doc.refilterTx(tx, t.sheet, [old, area])
  }
  tx.refetch = true
  doc.engine.tableChanged(t.name)
}
