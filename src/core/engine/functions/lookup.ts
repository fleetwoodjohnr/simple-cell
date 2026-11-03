// Lookup and reference functions, including the dynamic-array functions.
import { applySpec, areaArray, defineFunctions, getFunction, invokeLambda, toArray, toScalar, type EvalContext } from '../evaluate'
import { parseFormula, printFormula } from '../parser'
import { evaluate } from '../evaluate'
import { CellError, ERR, MAX_COLS, MAX_ROWS, compareScalars, isArray, isError, isLambda, isRef, makeArray, toBool, type Area, type ArrayValue, type RefValue, type Scalar, type Value } from '../values'
import { colName, quoteSheet } from '../../refs'
import { fn, hasWildcard, MANY, wildcardRegex } from './helpers'

// ---------- matching ----------

function exactTest(v: Scalar, wildcards: boolean): (x: Scalar) => boolean {
  if (typeof v === 'string') {
    if (wildcards && (hasWildcard(v) || v.includes('~'))) {
      const re = wildcardRegex(v)
      return (x) => typeof x === 'string' && re.test(x)
    }
    const low = v.toLowerCase()
    return (x) => typeof x === 'string' && x.toLowerCase() === low
  }
  if (typeof v === 'number') return (x) => typeof x === 'number' && compareScalars(x, v) === 0
  if (v === null) return (x) => x === null || x === ''
  return (x) => x === v
}

const comparable = (a: Scalar, b: Scalar) => a !== null && !(a instanceof CellError) && typeof a === typeof b

/**
 * Binary search like Excel's approximate lookups: the last position whose value is ≤ v (asc) or ≥ v (desc).
 * Values of another type than v are skipped.
 */
function binarySearch(list: Scalar[], v: Scalar, desc = false): number {
  let lo = 0
  let hi = list.length - 1
  let best = -1
  while (lo <= hi) {
    let mid = (lo + hi) >> 1
    // step left over items of another type
    let m = mid
    while (m >= lo && !comparable(list[m], v)) m--
    if (m < lo) {
      lo = mid + 1
      continue
    }
    mid = m
    const c = compareScalars(list[mid], v)
    if (desc ? c >= 0 : c <= 0) {
      best = mid
      lo = mid + 1
    } else hi = mid - 1
  }
  return best
}

/** XMATCH/XLOOKUP matching. Returns the 0-based index or -1. */
function xmatch(list: Scalar[], v: Scalar, matchMode: number, searchMode: number): number {
  if (searchMode === 2 || searchMode === -2) {
    const desc = searchMode === -2
    // binary search for exact or neighbour
    let lo = 0
    let hi = list.length - 1
    while (lo <= hi) {
      const mid = (lo + hi) >> 1
      const c = compareScalars(list[mid], v)
      if (c === 0 && comparable(list[mid], v)) return mid
      if (desc ? c > 0 : c < 0) lo = mid + 1
      else hi = mid - 1
    }
    if (matchMode === -1) return desc ? (lo < list.length ? lo : -1) : hi >= 0 ? hi : -1
    if (matchMode === 1) return desc ? (hi >= 0 ? hi : -1) : lo < list.length ? lo : -1
    return -1
  }
  const order = searchMode === -1 ? [...list.keys()].reverse() : [...list.keys()]
  const test = exactTest(v, matchMode === 2)
  for (const i of order) if (test(list[i])) return i
  if (matchMode === -1 || matchMode === 1) {
    let best = -1
    for (const i of order) {
      const x = list[i]
      if (!comparable(x, v)) continue
      const c = compareScalars(x, v)
      if (matchMode === -1 ? c < 0 : c > 0) {
        if (best < 0) best = i
        else {
          const cb = compareScalars(x, list[best])
          if (matchMode === -1 ? cb > 0 : cb < 0) best = i
        }
      }
    }
    return best
  }
  return -1
}

/** One row or column of an array as a flat list (lookup vectors). */
function vector(a: ArrayValue): Scalar[] | null {
  if (a.rows === 1 || a.cols === 1) return a.data
  return null
}

const row = (a: ArrayValue, r: number) => makeArray(1, a.cols, a.data.slice(r * a.cols, (r + 1) * a.cols))
const col = (a: ArrayValue, c: number) =>
  makeArray(
    a.rows,
    1,
    Array.from({ length: a.rows }, (_, r) => a.data[r * a.cols + c])
  )

function refArea(v: Value): Area | null {
  return isRef(v) && v.areas.length === 1 ? v.areas[0] : null
}

/** The values to look at for lookups: a reference is read (clipped to the used range for whole columns). */
function lookupArray(v: Value, ctx: EvalContext): ArrayValue | CellError {
  if (isRef(v)) {
    if (v.areas.length !== 1) return ERR.NA
    const a = v.areas[0]
    const sh = ctx.wb.sheetById(a.sheet)
    if (!sh) return ERR.REF
    const u = sh.used()
    // Lookups over whole columns/rows only need the used part (the rest is empty).
    const r1 = a.r1 - a.r0 > 100_000 ? Math.max(a.r0, Math.min(a.r1, u.r)) : a.r1
    const c1 = a.c1 - a.c0 > 1000 ? Math.max(a.c0, Math.min(a.c1, u.c)) : a.c1
    return areaArray({ ...a, r1, c1 }, ctx)
  }
  return toArray(v, ctx)
}

// ---------- reference helpers ----------

const refOf = (a: Area): RefValue => ({ kind: 'ref', areas: [a] })

function offsetRef(base: Area, dr: number, dc: number, h: number, w: number): RefValue | CellError {
  const r0 = base.r0 + dr
  const c0 = base.c0 + dc
  const r1 = r0 + h - 1
  const c1 = c0 + w - 1
  if (h === 0 || w === 0) return ERR.REF
  const a = { sheet: base.sheet, r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) }
  if (a.r0 < 0 || a.c0 < 0 || a.r1 >= MAX_ROWS || a.c1 >= MAX_COLS) return ERR.REF
  return refOf(a)
}

function indirect(text: string, a1: boolean, ctx: EvalContext): Value {
  let src = text.trim()
  if (!src) return ERR.REF
  if (!a1) src = r1c1ToA1(src, ctx.row, ctx.col) ?? ''
  if (!src) return ERR.REF
  try {
    const ast = parseFormula(src, { row: ctx.row, col: ctx.col, sheetId: (n) => ctx.wb.sheetByName(n)?.id ?? null })
    if (ast.t !== 'ref' && ast.t !== 'name' && ast.t !== 'struct' && !(ast.t === 'bin' && ast.op === ':')) return ERR.REF
    const v = evaluate(ast, ctx)
    return isRef(v) ? v : ERR.REF
  } catch {
    return ERR.REF
  }
}

/** R1C1 reference text (R2C3, R[-1]C, R1C1:R2C2, Sheet1!R1C1) as A1 text; null when invalid. */
export function r1c1ToA1(s: string, row: number, colIdx: number): string | null {
  const bang = s.lastIndexOf('!')
  const prefix = bang >= 0 ? s.slice(0, bang + 1) : ''
  const body = bang >= 0 ? s.slice(bang + 1) : s
  const part = (p: string): string | null => {
    const m = /^R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?$/i.exec(p)
    if (!m) return null
    const rr = m[1] === undefined ? row : m[1].startsWith('[') ? row + Number(m[1].slice(1, -1)) : Number(m[1]) - 1
    const cc = m[2] === undefined ? colIdx : m[2].startsWith('[') ? colIdx + Number(m[2].slice(1, -1)) : Number(m[2]) - 1
    if (rr < 0 || cc < 0 || rr >= MAX_ROWS || cc >= MAX_COLS) return null
    return `${m[2] && !m[2].startsWith('[') ? '$' : ''}${colName(cc)}${m[1] && !m[1].startsWith('[') ? '$' : ''}${rr + 1}`
  }
  const pieces = body.split(':').map(part)
  if (pieces.some((p) => p === null) || pieces.length > 2) return null
  return prefix + pieces.join(':')
}

// ---------- array helpers ----------

function sliceRows(a: ArrayValue, idx: number[]): ArrayValue | CellError {
  if (!idx.length) return ERR.CALC
  const data: Scalar[] = []
  for (const r of idx) data.push(...a.data.slice(r * a.cols, (r + 1) * a.cols))
  return makeArray(idx.length, a.cols, data)
}

function sliceCols(a: ArrayValue, idx: number[]): ArrayValue | CellError {
  if (!idx.length) return ERR.CALC
  const data: Scalar[] = []
  for (let r = 0; r < a.rows; r++) for (const c of idx) data.push(a.data[r * a.cols + c])
  return makeArray(a.rows, idx.length, data)
}

const transpose = (a: ArrayValue): ArrayValue => {
  const data: Scalar[] = new Array(a.rows * a.cols)
  for (let r = 0; r < a.rows; r++) for (let c = 0; c < a.cols; c++) data[c * a.rows + r] = a.data[r * a.cols + c]
  return makeArray(a.cols, a.rows, data)
}

function stack(arrays: ArrayValue[], vertical: boolean): ArrayValue {
  if (!vertical) return transpose(stack(arrays.map(transpose), true))
  const cols = Math.max(...arrays.map((a) => a.cols))
  const data: Scalar[] = []
  let rows = 0
  for (const a of arrays) {
    for (let r = 0; r < a.rows; r++) for (let c = 0; c < cols; c++) data.push(c < a.cols ? a.data[r * a.cols + c] : ERR.NA)
    rows += a.rows
  }
  return makeArray(rows, cols, data)
}

function arrays(args: Value[], ctx: EvalContext): ArrayValue[] | CellError {
  const out: ArrayValue[] = []
  for (const v of args) {
    const a = toArray(v, ctx)
    if (isError(a)) return a
    out.push(a)
  }
  return out
}

/** Row keys for UNIQUE/GROUPBY: case-insensitive like Excel. */
const keyOf = (xs: Scalar[]) => JSON.stringify(xs.map((x) => (typeof x === 'string' ? 's' + x.toLowerCase() : x instanceof CellError ? 'e' + x.code : x)))

function sortIndexes(keys: Scalar[][], orders: number[]): number[] {
  const idx = keys[0] ? keys[0].map((_, i) => i) : []
  return idx.sort((i, j) => {
    for (let k = 0; k < keys.length; k++) {
      const a = keys[k][i]
      const b = keys[k][j]
      // Errors sort last, empties after everything (Excel puts blanks last in either order)
      const ea = a instanceof CellError
      const eb = b instanceof CellError
      if (ea || eb) {
        if (ea && eb) continue
        return ea ? 1 : -1
      }
      if (a === null || b === null) {
        if (a === b) continue
        return a === null ? 1 : -1
      }
      const c = compareScalars(a, b)
      if (c) return c * orders[k]
    }
    return i - j
  })
}

// ---------- definitions ----------

defineFunctions([
  fn('VLOOKUP', 3, 4, 'vrnb', ([v, table, c, approx = true], ctx) => lookupIn(v, table, c, approx, ctx, true)),
  fn('HLOOKUP', 3, 4, 'vrnb', ([v, table, r, approx = true], ctx) => lookupIn(v, table, r, approx, ctx, false)),
  fn('LOOKUP', 2, 3, 'vrr', ([v, lv, rv], ctx) => {
    if (v instanceof CellError) return v
    const a = lookupArray(lv, ctx)
    if (isError(a)) return a
    let keys: Scalar[]
    let results: Scalar[]
    if (rv === undefined) {
      // Array form: search the first row or column, return from the last.
      const byRow = a.cols > a.rows
      keys = byRow ? row(a, 0).data : col(a, 0).data
      results = byRow ? row(a, a.rows - 1).data : col(a, a.cols - 1).data
    } else {
      const k = vector(a)
      const ra = lookupArray(rv, ctx)
      if (isError(ra)) return ra
      const r = vector(ra)
      if (!k || !r) return ERR.NA
      keys = k
      results = r
    }
    const i = binarySearch(keys, v)
    if (i < 0) return ERR.NA
    return i < results.length ? (results[i] ?? 0) : ERR.NA
  }),
  fn('MATCH', 2, 3, 'vrn', ([v, lv, type = 1], ctx) => {
    if (v instanceof CellError) return v
    const a = lookupArray(lv, ctx)
    if (isError(a)) return a
    const list = vector(a)
    if (!list) return ERR.NA
    const t = Math.sign(type)
    const i = t === 0 ? list.findIndex(exactTest(v, true)) : binarySearch(list, v, t < 0)
    return i < 0 ? ERR.NA : i + 1
  }),
  fn('XMATCH', 2, 4, 'vrnn', ([v, lv, mm = 0, sm = 1], ctx) => {
    if (v instanceof CellError) return v
    const a = lookupArray(lv, ctx)
    if (isError(a)) return a
    const list = vector(a)
    if (!list) return ERR.VALUE
    if (![0, -1, 1, 2].includes(mm) || ![1, -1, 2, -2].includes(sm)) return ERR.VALUE
    const i = xmatch(list, v, mm, sm)
    return i < 0 ? ERR.NA : i + 1
  }),
  fn('XLOOKUP', 3, 6, 'vrrrnn', ([v, lv, rv, notFound, mm = 0, sm = 1], ctx) => {
    if (v instanceof CellError) return v
    const a = lookupArray(lv, ctx)
    if (isError(a)) return a
    const list = vector(a)
    if (!list) return ERR.VALUE
    if (![0, -1, 1, 2].includes(mm) || ![1, -1, 2, -2].includes(sm)) return ERR.VALUE
    const i = xmatch(list, v, mm, sm)
    if (i < 0) return notFound === undefined ? ERR.NA : notFound
    // Return the matching row/column of the return array; a reference stays a reference.
    const area = refArea(rv)
    const vertical = a.cols === 1
    if (area) {
      if (vertical) {
        if (area.r1 - area.r0 + 1 !== list.length) return ERR.VALUE
        return refOf({ ...area, r0: area.r0 + i, r1: area.r0 + i })
      }
      if (area.c1 - area.c0 + 1 !== list.length) return ERR.VALUE
      return refOf({ ...area, c0: area.c0 + i, c1: area.c0 + i })
    }
    const r = toArray(rv, ctx)
    if (isError(r)) return r
    if (vertical) return r.rows === list.length ? row(r, i) : ERR.VALUE
    return r.cols === list.length ? col(r, i) : ERR.VALUE
  }),
  fn('INDEX', 2, 4, 'rnnn', ([ref, r, c, areaNum = 1], ctx) => {
    if (isError(ref)) return ref
    r = Math.trunc(r ?? 0)
    if (c === undefined && !isRef(ref) && isArray(ref) && (ref.rows === 1 || ref.cols === 1)) {
      // INDEX(single row or column, n)
      const list = ref.data
      if (r < 0 || r > list.length) return ERR.REF
      return r === 0 ? ref : (list[r - 1] ?? 0)
    }
    c = Math.trunc(c ?? 0)
    if (r < 0 || c < 0) return ERR.VALUE
    if (isRef(ref)) {
      const a = ref.areas[Math.trunc(areaNum) - 1]
      if (!a) return ERR.REF
      const h = a.r1 - a.r0 + 1
      const w = a.c1 - a.c0 + 1
      // A single row/column with one index picks along it.
      if (c === 0 && (h === 1 || w === 1) && r > 0 && (h === 1 ? r <= w : r <= h)) {
        return h === 1 ? refOf({ ...a, c0: a.c0 + r - 1, c1: a.c0 + r - 1 }) : refOf({ ...a, r0: a.r0 + r - 1, r1: a.r0 + r - 1 })
      }
      if (r > h || c > w) return ERR.REF
      return refOf({
        sheet: a.sheet,
        r0: r === 0 ? a.r0 : a.r0 + r - 1,
        r1: r === 0 ? a.r1 : a.r0 + r - 1,
        c0: c === 0 ? a.c0 : a.c0 + c - 1,
        c1: c === 0 ? a.c1 : a.c0 + c - 1
      })
    }
    const arr = toArray(ref, ctx)
    if (isError(arr)) return arr
    if (arr.rows === 1 && c === 0 && r > 0) {
      c = r
      r = 1
    }
    if (r > arr.rows || c > arr.cols) return ERR.REF
    if (r === 0 && c === 0) return arr
    if (r === 0) return col(arr, c - 1)
    if (c === 0) return row(arr, r - 1)
    return arr.data[(r - 1) * arr.cols + (c - 1)]
  }),
  fn('OFFSET', 3, 5, 'rnnnn', ([ref, dr, dc, h, w]) => {
    if (isError(ref)) return ref
    if (!isRef(ref) || ref.areas.length !== 1) return ERR.VALUE
    const a = ref.areas[0]
    return offsetRef(a, Math.trunc(dr), Math.trunc(dc), Math.trunc(h ?? a.r1 - a.r0 + 1), Math.trunc(w ?? a.c1 - a.c0 + 1))
  }, { volatile: true }),
  fn('INDIRECT', 1, 2, 'sb', ([text, a1 = true], ctx) => indirect(text, a1, ctx), { volatile: true }),
  fn('ADDRESS', 2, 5, 'nnnbs', ([r, c, abs = 1, a1 = true, sheet]) => {
    r = Math.trunc(r)
    c = Math.trunc(c)
    abs = Math.trunc(abs)
    if (r < 1 || c < 1 || r > MAX_ROWS || c > MAX_COLS || abs < 1 || abs > 4) return ERR.VALUE
    const rowAbs = abs === 1 || abs === 2
    const colAbs = abs === 1 || abs === 3
    const pre = sheet === undefined ? '' : sheet === '' ? '!' : `${quoteSheet(sheet)}!`
    if (a1) return `${pre}${colAbs ? '$' : ''}${colName(c - 1)}${rowAbs ? '$' : ''}${r}`
    return `${pre}R${rowAbs ? r : `[${r}]`}C${colAbs ? c : `[${c}]`}`
  }),
  fn('AREAS', 1, 1, 'r', ([v]) => (isRef(v) ? v.areas.length : isError(v) ? v : ERR.VALUE)),
  fn('CHOOSE', 2, MANY, 'n|r', ([i, ...opts]) => {
    const k = Math.trunc(i)
    if (k < 1 || k > opts.length) return ERR.VALUE
    return opts[k - 1] ?? 0
  }),
  fn('ROW', 0, 1, 'r', ([v], ctx) => {
    if (v === undefined) return ctx.row + 1
    if (!isRef(v)) return isError(v) ? v : ERR.VALUE
    const a = v.areas[0]
    if (a.r0 === a.r1) return a.r0 + 1
    if (a.r1 - a.r0 >= MAX_ROWS - 1) return a.r0 + 1
    return makeArray(a.r1 - a.r0 + 1, 1, Array.from({ length: a.r1 - a.r0 + 1 }, (_, i) => a.r0 + 1 + i))
  }),
  fn('COLUMN', 0, 1, 'r', ([v], ctx) => {
    if (v === undefined) return ctx.col + 1
    if (!isRef(v)) return isError(v) ? v : ERR.VALUE
    const a = v.areas[0]
    if (a.c0 === a.c1) return a.c0 + 1
    return makeArray(1, a.c1 - a.c0 + 1, Array.from({ length: a.c1 - a.c0 + 1 }, (_, i) => a.c0 + 1 + i))
  }),
  fn('ROWS', 1, 1, 'r', ([v], ctx) => {
    if (isRef(v)) return v.areas.length === 1 ? v.areas[0].r1 - v.areas[0].r0 + 1 : ERR.REF
    const a = toArray(v, ctx)
    return isError(a) ? a : a.rows
  }),
  fn('COLUMNS', 1, 1, 'r', ([v], ctx) => {
    if (isRef(v)) return v.areas.length === 1 ? v.areas[0].c1 - v.areas[0].c0 + 1 : ERR.REF
    const a = toArray(v, ctx)
    return isError(a) ? a : a.cols
  }),
  fn('TRANSPOSE', 1, 1, 'a', ([a]) => transpose(a)),
  fn('FORMULATEXT', 1, 1, 'r', ([v], ctx) => {
    if (!isRef(v)) return isError(v) ? v : ERR.VALUE
    const a = v.areas[0]
    const f = ctx.wb.sheetById(a.sheet)?.get(a.r0, a.c0)?.f
    if (!f) return ERR.NA
    return '=' + printFormula(f.ast, { row: a.r0, col: a.c0, sheetName: (id) => ctx.wb.sheetById(id)?.name ?? null })
  }),
  fn('HYPERLINK', 1, 2, 'vv', ([link, friendly]) => (friendly === undefined ? link : friendly ?? 0)),
  fn('FILTER', 2, 3, 'aar', ([a, inc, empty], ctx) => {
    const byRow = inc.cols === 1 && inc.rows === a.rows
    const byCol = inc.rows === 1 && inc.cols === a.cols
    if (!byRow && !byCol) return ERR.VALUE
    const keep: number[] = []
    for (let i = 0; i < inc.data.length; i++) {
      const x = inc.data[i]
      if (x instanceof CellError) return x
      const b = toBool(x)
      if (isError(b)) return ERR.VALUE
      if (b) keep.push(i)
    }
    if (!keep.length) return empty === undefined ? ERR.CALC : toScalarOrArray(empty, ctx)
    return byRow && !(byCol && a.rows === 1) ? sliceRows(a, keep) : sliceCols(a, keep)
  }),
  fn('SORT', 1, 4, 'arrb', ([a, idx, ord, byCol = false], ctx) => {
    const ia = idx === undefined ? [1] : numList(idx, ctx)
    const oa = ord === undefined ? [1] : numList(ord, ctx)
    if (isError(ia)) return ia
    if (isError(oa)) return oa
    const src = byCol ? transpose(a) : a
    if (ia.some((i) => i < 1 || i > src.cols) || oa.some((o) => o !== 1 && o !== -1)) return ERR.VALUE
    const keys = ia.map((i) => col(src, i - 1).data)
    const orders = ia.map((_, k) => oa[Math.min(k, oa.length - 1)])
    const out = sliceRows(src, sortIndexes(keys, orders))
    return isError(out) ? out : byCol ? transpose(out) : out
  }),
  fn('SORTBY', 2, MANY, 'a|ar', ([a, ...rest], ctx) => {
    const keys: Scalar[][] = []
    const orders: number[] = []
    let byCol: boolean | null = null
    for (let i = 0; i < rest.length; i += 2) {
      const k = toArray(rest[i], ctx)
      if (isError(k)) return k
      const o = rest[i + 1] === undefined ? 1 : toScalar(rest[i + 1], ctx)
      if (o !== 1 && o !== -1) return ERR.VALUE
      const vertical = k.cols === 1 && k.rows === a.rows
      const horizontal = k.rows === 1 && k.cols === a.cols
      if (!vertical && !horizontal) return ERR.VALUE
      const bc = !vertical
      if (byCol !== null && byCol !== bc) return ERR.VALUE
      byCol = bc
      keys.push(k.data)
      orders.push(o)
    }
    const src = byCol ? transpose(a) : a
    const out = sliceRows(src, sortIndexes(keys, orders))
    return isError(out) ? out : byCol ? transpose(out) : out
  }),
  fn('UNIQUE', 1, 3, 'abb', ([a, byCol = false, once = false]) => {
    const src = byCol ? transpose(a) : a
    const counts = new Map<string, { first: number; n: number }>()
    for (let r = 0; r < src.rows; r++) {
      const k = keyOf(src.data.slice(r * src.cols, (r + 1) * src.cols))
      const e = counts.get(k)
      if (e) e.n++
      else counts.set(k, { first: r, n: 1 })
    }
    const idx = [...counts.values()].filter((e) => !once || e.n === 1).map((e) => e.first)
    const out = sliceRows(src, idx)
    return isError(out) ? out : byCol ? transpose(out) : out
  }),
  fn('TAKE', 2, 3, 'ann', ([a, r, c]) => {
    const rr = r === undefined ? a.rows : Math.trunc(r)
    const cc = c === undefined ? a.cols : Math.trunc(c)
    if (rr === 0 || cc === 0) return ERR.CALC
    const rows = rr > 0 ? range(0, Math.min(rr, a.rows)) : range(Math.max(0, a.rows + rr), a.rows)
    const cols = cc > 0 ? range(0, Math.min(cc, a.cols)) : range(Math.max(0, a.cols + cc), a.cols)
    const x = sliceRows(a, rows)
    return isError(x) ? x : sliceCols(x, cols)
  }),
  fn('DROP', 2, 3, 'ann', ([a, r = 0, c = 0]) => {
    const rr = Math.trunc(r)
    const cc = Math.trunc(c)
    const rows = rr >= 0 ? range(rr, a.rows) : range(0, Math.max(0, a.rows + rr))
    const cols = cc >= 0 ? range(cc, a.cols) : range(0, Math.max(0, a.cols + cc))
    const x = sliceRows(a, rows)
    return isError(x) ? x : sliceCols(x, cols)
  }),
  fn('CHOOSEROWS', 2, MANY, 'a|r', ([a, ...ns], ctx) => {
    const idx: number[] = []
    for (const n of ns) {
      const list = numList(n, ctx)
      if (isError(list)) return list
      for (const k of list) {
        const i = Math.trunc(k)
        if (i === 0 || Math.abs(i) > a.rows) return ERR.VALUE
        idx.push(i > 0 ? i - 1 : a.rows + i)
      }
    }
    return sliceRows(a, idx)
  }),
  fn('CHOOSECOLS', 2, MANY, 'a|r', ([a, ...ns], ctx) => {
    const idx: number[] = []
    for (const n of ns) {
      const list = numList(n, ctx)
      if (isError(list)) return list
      for (const k of list) {
        const i = Math.trunc(k)
        if (i === 0 || Math.abs(i) > a.cols) return ERR.VALUE
        idx.push(i > 0 ? i - 1 : a.cols + i)
      }
    }
    return sliceCols(a, idx)
  }),
  fn('EXPAND', 2, 4, 'annv', ([a, r, c, pad]) => {
    const rows = r === undefined ? a.rows : Math.trunc(r)
    const cols = c === undefined ? a.cols : Math.trunc(c)
    if (rows < a.rows || cols < a.cols) return ERR.VALUE
    const p: Scalar = pad === undefined ? ERR.NA : pad
    const data: Scalar[] = []
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) data.push(i < a.rows && j < a.cols ? a.data[i * a.cols + j] : p)
    return makeArray(rows, cols, data)
  }),
  fn('VSTACK', 1, MANY, 'r', (args, ctx) => {
    const as = arrays(args, ctx)
    return isError(as) ? as : stack(as, true)
  }),
  fn('HSTACK', 1, MANY, 'r', (args, ctx) => {
    const as = arrays(args, ctx)
    return isError(as) ? as : stack(as, false)
  }),
  fn('TOCOL', 1, 3, 'annb', ([a, ignore = 0, byCol = false]) => toLine(a, ignore, byCol, true)),
  fn('TOROW', 1, 3, 'annb', ([a, ignore = 0, byCol = false]) => toLine(a, ignore, byCol, false)),
  fn('WRAPROWS', 2, 3, 'anv', ([a, n, pad]) => wrap(a, n, pad, true)),
  fn('WRAPCOLS', 2, 3, 'anv', ([a, n, pad]) => wrap(a, n, pad, false)),
  fn('TRIMRANGE', 1, 3, 'rnn', ([v, tr = 3, tc = 3], ctx) => {
    const a = refArea(v)
    if (!a) return isError(v) ? v : ERR.VALUE
    const sh = ctx.wb.sheetById(a.sheet)
    if (!sh) return ERR.REF
    let r0 = Infinity
    let r1 = -1
    let c0 = Infinity
    let c1 = -1
    ctx.ensure(a)
    sh.forEachIn(a, (r, c, cell) => {
      if (cell.v === null) return
      r0 = Math.min(r0, r)
      r1 = Math.max(r1, r)
      c0 = Math.min(c0, c)
      c1 = Math.max(c1, c)
    })
    if (r1 < 0) return ERR.REF
    const out = { ...a }
    if (tr === 1 || tr === 3) out.r0 = r0
    if (tr === 2 || tr === 3) out.r1 = r1
    if (tc === 1 || tc === 3) out.c0 = c0
    if (tc === 2 || tc === 3) out.c1 = c1
    return refOf(out)
  }),
  fn('GROUPBY', 3, 8, 'aarnnnrn', ([keys, vals, f, headers, totalDepth, sortOrder, filt], ctx) => groupBy(keys, vals, f, headers, totalDepth, sortOrder, filt, ctx))
])

const range = (a: number, b: number) => Array.from({ length: Math.max(0, b - a) }, (_, i) => a + i)

function numList(v: Value, ctx: EvalContext): number[] | CellError {
  const a = toArray(v, ctx)
  if (isError(a)) return a
  const out: number[] = []
  for (const x of a.data) {
    if (x instanceof CellError) return x
    if (typeof x !== 'number') return ERR.VALUE
    out.push(x)
  }
  return out
}

function toScalarOrArray(v: Value, ctx: EvalContext): Value {
  return isRef(v) ? toArray(v, ctx) : v
}

function toLine(a: ArrayValue, ignore: number, byCol: boolean, asCol: boolean): Value {
  const src = byCol ? transpose(a) : a
  const out = src.data.filter((x) => !((ignore === 1 || ignore === 3) && x === null) && !((ignore === 2 || ignore === 3) && x instanceof CellError))
  if (!out.length) return ERR.CALC
  return asCol ? makeArray(out.length, 1, out) : makeArray(1, out.length, out)
}

function wrap(a: ArrayValue, n: number, pad: Scalar | undefined, rows: boolean): Value {
  if (a.rows !== 1 && a.cols !== 1) return ERR.VALUE
  const k = Math.trunc(n)
  if (k < 1) return ERR.NUM
  const list = a.data
  const lines = Math.ceil(list.length / k)
  const p: Scalar = pad === undefined ? ERR.NA : pad
  const data: Scalar[] = []
  for (let i = 0; i < lines * k; i++) data.push(i < list.length ? list[i] : p)
  const out = makeArray(lines, k, data)
  return rows ? out : transpose(out)
}

function lookupIn(v: Scalar, table: Value, n: number, approx: boolean, ctx: EvalContext, vertical: boolean): Value {
  if (v instanceof CellError) return v
  const a = lookupArray(table, ctx)
  if (isError(a)) return a
  const k = Math.trunc(n)
  const width = vertical ? a.cols : a.rows
  if (k < 1) return ERR.VALUE
  if (k > width) return ERR.REF
  const keys = vertical ? col(a, 0).data : row(a, 0).data
  const i = approx ? binarySearch(keys, v) : keys.findIndex(exactTest(v, true))
  if (i < 0) return ERR.NA
  const r = vertical ? a.data[i * a.cols + (k - 1)] : a.data[(k - 1) * a.cols + i]
  return r === null ? 0 : r
}

/** GROUPBY: group rows by key columns and aggregate the value columns with a function (or LAMBDA). */
function groupBy(keys: ArrayValue, vals: ArrayValue, f: Value, headers: number | undefined, totalDepth: number | undefined, sortOrder: number | undefined, filt: Value, ctx: EvalContext): Value {
  if (keys.rows !== vals.rows) return ERR.VALUE
  const agg = isLambda(f) ? f : null
  if (!agg) return ERR.VALUE
  const hasHeaders = headers === 1 || headers === 3
  const showHeaders = headers === 2 || headers === 3
  const start = hasHeaders ? 1 : 0
  let filter: Scalar[] | null = null
  if (filt !== undefined) {
    const fa = toArray(filt, ctx)
    if (isError(fa)) return fa
    filter = fa.data
  }
  const groups = new Map<string, { key: Scalar[]; rows: number[] }>()
  for (let r = start; r < keys.rows; r++) {
    if (filter && !toBool(filter[r] ?? false)) continue
    const k = keys.data.slice(r * keys.cols, (r + 1) * keys.cols)
    const id = keyOf(k)
    const g = groups.get(id)
    if (g) g.rows.push(r)
    else groups.set(id, { key: k, rows: [r] })
  }
  const list = [...groups.values()]
  const order = sortOrder === undefined ? 1 : Math.trunc(sortOrder)
  if (order !== 0) {
    const ki = Math.abs(order) - 1
    list.sort((a, b) => {
      const x = ki < keys.cols ? a.key[ki] : null
      const y = ki < keys.cols ? b.key[ki] : null
      return compareScalars(x, y) * Math.sign(order)
    })
  }
  const out: Scalar[][] = []
  if (showHeaders) {
    const head: Scalar[] = []
    for (let c = 0; c < keys.cols; c++) head.push(hasHeaders ? keys.data[c] : `Key${c + 1}`)
    for (let c = 0; c < vals.cols; c++) head.push(hasHeaders ? vals.data[c] : `Value${c + 1}`)
    out.push(head)
  }
  const aggregate = (rows: number[]) => {
    const res: Scalar[] = []
    for (let c = 0; c < vals.cols; c++) {
      const column = rows.map((r) => vals.data[r * vals.cols + c])
      const v = invokeLambda(agg, [makeArray(column.length, 1, column)], ctx)
      res.push(isArray(v) ? v.data[0] : isRef(v) ? toScalar(v, ctx) : (v as Scalar))
    }
    return res
  }
  for (const g of list) out.push([...g.key, ...aggregate(g.rows)])
  if ((totalDepth ?? 1) !== 0) {
    const all = list.flatMap((g) => g.rows)
    out.push(['Total', ...Array(keys.cols - 1).fill(null), ...aggregate(all)])
  }
  const cols = keys.cols + vals.cols
  return makeArray(out.length, cols, out.flat())
}

export { getFunction, applySpec }
