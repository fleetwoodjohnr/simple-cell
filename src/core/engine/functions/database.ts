// Database functions (DSUM, DCOUNT…), and the few web/add-in functions that need no network: ENCODEURL, EUROCONVERT.
import { defineFunctions, toArray, type EvalContext } from '../evaluate'
import { CellError, ERR, isError, type ArrayValue, type Scalar, type Value } from '../values'
import { criteria, fn, mean, sum, variance } from './helpers'

/** Rows of the database that match the criteria range: OR across criteria rows, AND across columns. */
function selected(db: ArrayValue, crit: ArrayValue): number[] | CellError {
  const head = db.data.slice(0, db.cols).map((h) => String(h ?? '').toLowerCase())
  const critHead = crit.data.slice(0, crit.cols).map((h) => String(h ?? '').toLowerCase())
  const colOf = critHead.map((h) => head.indexOf(h))
  // ponytail: computed criteria (a formula under a blank or non-matching header) are not supported yet.
  if (colOf.some((c, i) => c < 0 && critHead[i] !== '')) return ERR.VALUE
  const rows: number[] = []
  for (let r = 1; r < db.rows; r++) {
    let any = false
    for (let cr = 1; cr < crit.rows && !any; cr++) {
      let all = true
      for (let c = 0; c < crit.cols && all; c++) {
        const cond = crit.data[cr * crit.cols + c]
        if (cond === null || cond === '' || colOf[c] < 0) continue
        if (!criteria(cond)(db.data[r * db.cols + colOf[c]])) all = false
      }
      any = all
    }
    if (crit.rows < 2) any = true
    if (any) rows.push(r)
  }
  return rows
}

function fieldIndex(db: ArrayValue, field: Scalar): number | CellError {
  if (typeof field === 'number') {
    const i = Math.trunc(field) - 1
    return i >= 0 && i < db.cols ? i : ERR.VALUE
  }
  if (typeof field === 'string') {
    const i = db.data.slice(0, db.cols).findIndex((h) => String(h ?? '').toLowerCase() === field.toLowerCase())
    return i < 0 ? ERR.VALUE : i
  }
  return ERR.VALUE
}

function dbValues(args: Value[], ctx: EvalContext, needField = true): { vals: Scalar[] } | CellError {
  const db = toArray(args[0], ctx)
  const crit = toArray(args[2], ctx)
  if (isError(db)) return db
  if (isError(crit)) return crit
  const rows = selected(db, crit)
  if (isError(rows)) return rows
  if (!needField && (args[1] === undefined || args[1] === null)) return { vals: rows.map(() => 1) }
  const f = fieldIndex(db, args[1] as Scalar)
  if (isError(f)) return f
  return { vals: rows.map((r) => db.data[r * db.cols + f]) }
}

function dfn(name: string, f: (xs: number[], vals: Scalar[]) => Scalar, needField = true) {
  return fn(name, 3, 3, 'rvr', (args, ctx) => {
    const d = dbValues(args, ctx, needField)
    if (isError(d)) return d
    const xs = d.vals.filter((v): v is number => typeof v === 'number')
    return f(xs, d.vals)
  })
}

// Fixed conversion rates of the euro (Council Regulation (EC) No 2866/98 and later accessions).
const EURO: Record<string, number> = { BEF: 40.3399, LUF: 40.3399, DEM: 1.95583, ESP: 166.386, FRF: 6.55957, IEP: 0.787564, ITL: 1936.27, NLG: 2.20371, ATS: 13.7603, PTE: 200.482, FIM: 5.94573, GRD: 340.75, SIT: 239.64, CYP: 0.585274, MTL: 0.4293, SKK: 30.126, EEK: 15.6466, LVL: 0.702804, LTL: 3.4528, HRK: 7.5345, EUR: 1 }
// Decimals Excel rounds each currency to.
const EURO_DEC: Record<string, number> = { BEF: 0, LUF: 0, DEM: 2, ESP: 0, FRF: 2, IEP: 2, ITL: 0, NLG: 2, ATS: 2, PTE: 0, FIM: 2, GRD: 0, SIT: 2, CYP: 2, MTL: 2, SKK: 2, EEK: 2, LVL: 2, LTL: 2, HRK: 2, EUR: 2 }

defineFunctions([
  dfn('DSUM', (xs) => sum(xs)),
  dfn('DAVERAGE', (xs) => (xs.length ? mean(xs) : ERR.DIV0)),
  dfn('DCOUNT', (xs) => xs.length, false),
  dfn('DCOUNTA', (_xs, vals) => vals.filter((v) => v !== null && v !== '').length, false),
  dfn('DMAX', (xs) => (xs.length ? Math.max(...xs) : 0)),
  dfn('DMIN', (xs) => (xs.length ? Math.min(...xs) : 0)),
  dfn('DPRODUCT', (xs) => (xs.length ? xs.reduce((a, b) => a * b, 1) : 0)),
  dfn('DSTDEV', (xs) => {
    const v = variance(xs, true)
    return isError(v) ? v : Math.sqrt(v)
  }),
  dfn('DSTDEVP', (xs) => {
    const v = variance(xs, false)
    return isError(v) ? v : Math.sqrt(v)
  }),
  dfn('DVAR', (xs) => variance(xs, true)),
  dfn('DVARP', (xs) => variance(xs, false)),
  dfn('DGET', (_xs, vals) => (vals.length === 0 ? ERR.VALUE : vals.length > 1 ? ERR.NUM : (vals[0] ?? 0))),
  fn('ENCODEURL', 1, 1, 's', ([s]) => encodeURIComponent(s).replace(/[!'()*]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase())),
  fn('EUROCONVERT', 3, 5, 'nssbn', ([v, from, to, full = false, tri]) => {
    const a = EURO[from.toUpperCase()]
    const b = EURO[to.toUpperCase()]
    if (!a || !b) return ERR.VALUE
    if (tri !== undefined && Math.trunc(tri) < 3) return ERR.VALUE
    let euros = v / a
    if (tri !== undefined && from.toUpperCase() !== 'EUR' && to.toUpperCase() !== 'EUR') euros = Number(euros.toFixed(Math.trunc(tri)))
    const out = euros * b
    if (full) return out
    const d = EURO_DEC[to.toUpperCase()]
    return Math.round(out * 10 ** d) / 10 ** d
  })
])
