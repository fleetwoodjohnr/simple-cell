// Pure SQLite execution over value snapshots. No workbook, filesystem or network access.
import initSqlJs from 'sql.js/dist/sql-wasm-inline.js'
import type { SqlValue, Statement } from 'sql.js'

type Value = string | number | null
type Rows = { columns: string[]; rows: Value[][] }
export type SqlInput = { sql: string; tables: { name: string; columns: string[]; rows: SqlValue[][] }[] }
let sqlite: ReturnType<typeof initSqlJs> | undefined
const quote = (s: string) => `"${s.replace(/"/g, '""')}"`
function drain(st: Statement): Rows {
  try {
    const columns = st.getColumnNames(), rows: Value[][] = []
    while (st.step()) {
      if (rows.length * columns.length >= 1_000_000) throw new Error('SQL results exceed one million cells. Use a filter or LIMIT.')
      rows.push(st.get().map((v) => v instanceof Uint8Array ? `[${v.length} bytes]` : v))
    }
    return { columns, rows }
  } finally { st.free() }
}
export async function executeSql(input: SqlInput) {
  if (!input.sql.trim()) throw new Error('Type a query first.')
  const SQL = await (sqlite ??= initSqlJs()), db = new SQL.Database()
  try {
    const read = (name: string) => drain(db.prepare(`SELECT * FROM ${quote(name)}`))
    const before = new Map<string, string>()
    for (const t of input.tables) {
      db.run(`CREATE TABLE ${quote(t.name)} (${t.columns.map(quote).join(',')})`)
      const st = db.prepare(`INSERT INTO ${quote(t.name)} VALUES (${t.columns.map(() => '?').join(',')})`)
      try { db.run('BEGIN'); for (const row of t.rows) st.run(row); db.run('COMMIT') } finally { st.free() }
      before.set(t.name, JSON.stringify(read(t.name)))
    }
    let result: Rows | null = null
    for (const st of db.iterateStatements(input.sql)) {
      const text = st.getSQL().replace(/^(?:\s|--[^\n]*(?:\n|$)|\/\*[\s\S]*?\*\/)+/, '')
      if (!/^(SELECT|WITH|UPDATE|INSERT|DELETE|EXPLAIN)\b/i.test(text)) { st.free(); throw new Error('Use SELECT, WITH, UPDATE, INSERT or DELETE. Schema changes are not supported.') }
      const rows = drain(st)
      if (rows.columns.length) result = rows
    }
    const changed = input.tables.flatMap((t) => { const data = read(t.name); return JSON.stringify(data) === before.get(t.name) ? [] : [{ name: t.name, data }] })
    return { result, changed }
  } finally { db.close() }
}
