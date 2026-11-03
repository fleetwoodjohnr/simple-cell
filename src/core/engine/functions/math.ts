// Math and trigonometry functions.
import { defineFunctions, toArray, type EvalContext } from '../evaluate'
import { CellError, ERR, isArray, isError, isRef, makeArray, round15, type ArrayValue, type Scalar, type Value } from '../values'
import { checkNum, fn, fromMatrix, grid, matchIndexes, matrix, MANY, numbers, rangeValues, roundHalfAway, sum, variance, mean } from './helpers'
import { aggregateStat } from './stats-core'

const n1 = (name: string, f: (x: number) => number | CellError) => fn(name, 1, 1, 'n', ([x]) => {
  const r = f(x)
  return typeof r === 'number' ? checkNum(r) : r
})

function fact(n: number): number {
  if (n < 0) return NaN
  let r = 1
  for (let i = 2; i <= Math.floor(n); i++) r *= i
  return r
}

export function combin(n: number, k: number): number {
  if (k < 0 || n < 0 || k > n) return NaN
  k = Math.min(k, n - k)
  let r = 1
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i
  return Math.round(r)
}

const gcd2 = (a: number, b: number): number => (b === 0 ? a : gcd2(b, a % b))

function intArgs(args: Value[], ctx: EvalContext): number[] | CellError {
  const xs = numbers(args, ctx)
  if (isError(xs)) return xs
  if (xs.some((x) => x < 0)) return ERR.NUM
  return xs.map(Math.trunc)
}

function ceilingLegacy(x: number, s: number): number | CellError {
  if (s === 0 || x === 0) return 0
  if (x > 0 && s < 0) return ERR.NUM
  // Negative numbers with a positive significance round toward zero; both negative round away from zero.
  return round15(Math.ceil(round15(x / s)) * s)
}

function floorLegacy(x: number, s: number): number | CellError {
  if (x === 0) return 0
  if (s === 0) return ERR.DIV0
  if (x > 0 && s < 0) return ERR.NUM
  return round15(Math.floor(round15(x / s)) * s)
}

function mathRound(x: number, s: number, mode: number, up: boolean): number {
  if (s === 0 || x === 0) return 0
  s = Math.abs(s)
  let q = round15(x / s)
  if (x < 0 && mode) q = up ? Math.floor(q) : Math.ceil(q) // away from / toward zero for negatives
  else q = up ? Math.ceil(q) : Math.floor(q)
  return round15(q * s)
}

// ---------- SUBTOTAL / AGGREGATE ----------

const SUBTOTAL_FN: Record<number, string> = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P' }
const AGG_FN: Record<number, string> = { ...SUBTOTAL_FN, 7: 'STDEV.S', 8: 'STDEV.P', 12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC' }

function aggregateValues(name: string, xs: number[], count: number, k?: number): Scalar {
  switch (name) {
    case 'SUM':
      return sum(xs)
    case 'COUNT':
      return xs.length
    case 'COUNTA':
      return count
    case 'AVERAGE':
      return xs.length ? mean(xs) : ERR.DIV0
    case 'MAX':
      return xs.length ? Math.max(...xs) : 0
    case 'MIN':
      return xs.length ? Math.min(...xs) : 0
    case 'PRODUCT':
      return xs.length ? xs.reduce((a, b) => a * b, 1) : 0
    case 'STDEV.S': {
      const v = variance(xs, true)
      return isError(v) ? v : Math.sqrt(v)
    }
    case 'STDEV.P': {
      const v = variance(xs, false)
      return isError(v) ? v : Math.sqrt(v)
    }
    case 'VAR.S':
      return variance(xs, true)
    case 'VAR.P':
      return variance(xs, false)
    default:
      return aggregateStat(name, xs, k)
  }
}

function countNonEmpty(args: Value[], ctx: EvalContext, skipHidden: false | 'filtered' | 'all'): number {
  let n = 0
  for (const v of args) {
    if (!isRef(v)) continue
    for (const a of v.areas) {
      const sh = ctx.wb.sheetById(a.sheet)
      ctx.ensure(a)
      sh?.forEachIn(a, (r, _c, cell) => {
        if (cell.v === null) return
        const ri = sh.rowInfo.get(r)
        if (skipHidden && ri?.hidden && (skipHidden === 'all' || ri.filtered)) return
        if (cell.f && cell.f.ast.t === 'func' && (cell.f.ast.name === 'SUBTOTAL' || cell.f.ast.name === 'AGGREGATE')) return
        n++
      })
    }
  }
  return n
}

defineFunctions([
  n1('ABS', Math.abs),
  n1('ACOS', (x) => (x < -1 || x > 1 ? ERR.NUM : Math.acos(x))),
  n1('ACOSH', (x) => (x < 1 ? ERR.NUM : Math.acosh(x))),
  n1('ACOT', (x) => Math.PI / 2 - Math.atan(x)),
  n1('ACOTH', (x) => (Math.abs(x) <= 1 ? ERR.NUM : 0.5 * Math.log((x + 1) / (x - 1)))),
  n1('ASIN', (x) => (x < -1 || x > 1 ? ERR.NUM : Math.asin(x))),
  n1('ASINH', Math.asinh),
  n1('ATAN', Math.atan),
  fn('ATAN2', 2, 2, 'n', ([x, y]) => (x === 0 && y === 0 ? ERR.DIV0 : Math.atan2(y, x))),
  n1('ATANH', (x) => (x <= -1 || x >= 1 ? ERR.NUM : Math.atanh(x))),
  n1('COS', (x) => (Math.abs(x) >= 2 ** 27 ? ERR.NUM : Math.cos(x))),
  n1('COSH', Math.cosh),
  n1('COT', (x) => (x === 0 ? ERR.DIV0 : Math.abs(x) >= 2 ** 27 ? ERR.NUM : 1 / Math.tan(x))),
  n1('COTH', (x) => (x === 0 ? ERR.DIV0 : 1 / Math.tanh(x))),
  n1('CSC', (x) => (x === 0 ? ERR.DIV0 : Math.abs(x) >= 2 ** 27 ? ERR.NUM : 1 / Math.sin(x))),
  n1('CSCH', (x) => (x === 0 ? ERR.DIV0 : 1 / Math.sinh(x))),
  n1('SEC', (x) => (Math.abs(x) >= 2 ** 27 ? ERR.NUM : 1 / Math.cos(x))),
  n1('SECH', (x) => 1 / Math.cosh(x)),
  n1('SIN', (x) => (Math.abs(x) >= 2 ** 27 ? ERR.NUM : Math.sin(x))),
  n1('SINH', Math.sinh),
  n1('TAN', (x) => (Math.abs(x) >= 2 ** 27 ? ERR.NUM : Math.tan(x))),
  n1('TANH', Math.tanh),
  n1('DEGREES', (x) => (x * 180) / Math.PI),
  n1('RADIANS', (x) => (x * Math.PI) / 180),
  n1('EXP', Math.exp),
  n1('LN', (x) => (x <= 0 ? ERR.NUM : Math.log(x))),
  n1('LOG10', (x) => (x <= 0 ? ERR.NUM : Math.log10(x))),
  fn('LOG', 1, 2, 'n', ([x, b = 10]) => (x <= 0 || b <= 0 ? ERR.NUM : b === 1 ? ERR.DIV0 : Math.log(x) / Math.log(b))),
  n1('SQRT', (x) => (x < 0 ? ERR.NUM : Math.sqrt(x))),
  n1('SQRTPI', (x) => (x < 0 ? ERR.NUM : Math.sqrt(x * Math.PI))),
  fn('PI', 0, 0, '', () => Math.PI),
  n1('INT', Math.floor),
  n1('SIGN', Math.sign),
  n1('EVEN', (x) => {
    const a = Math.ceil(round15(Math.abs(x)))
    return Math.sign(x) * (a % 2 === 0 ? a : a + 1)
  }),
  n1('ODD', (x) => {
    const a = Math.ceil(round15(Math.abs(x)))
    const r = a % 2 === 1 ? a : a + 1
    return x < 0 ? -r : r
  }),
  n1('FACT', (x) => (x < 0 ? ERR.NUM : x > 170 ? ERR.NUM : fact(x))),
  n1('FACTDOUBLE', (x) => {
    if (x < -1) return ERR.NUM
    let r = 1
    for (let i = Math.floor(x); i > 1; i -= 2) r *= i
    return r
  }),
  fn('POWER', 2, 2, 'n', ([x, y]) => {
    if (x === 0 && y === 0) return ERR.NUM
    if (x === 0 && y < 0) return ERR.DIV0
    const r = Math.pow(x, y)
    return Number.isFinite(r) ? r : ERR.NUM
  }),
  fn('MOD', 2, 2, 'n', ([x, d]) => {
    if (d === 0) return ERR.DIV0
    const q = x / d
    if (Math.abs(q) >= 2 ** 50) return ERR.NUM
    const r = x - d * Math.floor(round15(q))
    return round15(r) === round15(d) ? 0 : r
  }),
  fn('QUOTIENT', 2, 2, 'n', ([x, d]) => (d === 0 ? ERR.DIV0 : Math.trunc(x / d))),
  fn('ROUND', 2, 2, 'n', ([x, d]) => roundHalfAway(x, d)),
  fn('ROUNDUP', 2, 2, 'n', ([x, d]) => {
    const p = Math.pow(10, Math.trunc(d))
    const v = round15(Math.abs(x) * p)
    return (Math.sign(x) * Math.ceil(v)) / p
  }),
  fn('ROUNDDOWN', 2, 2, 'n', ([x, d]) => {
    const p = Math.pow(10, Math.trunc(d))
    const v = round15(Math.abs(x) * p)
    return (Math.sign(x) * Math.floor(v)) / p
  }),
  fn('TRUNC', 1, 2, 'n', ([x, d = 0]) => {
    const p = Math.pow(10, Math.trunc(d))
    return Math.trunc(round15(x * p)) / p
  }),
  fn('MROUND', 2, 2, 'n', ([x, m]) => {
    if (m === 0) return 0
    if (x * m < 0) return ERR.NUM
    return round15(Math.round(round15(Math.abs(x / m))) * Math.abs(m)) * Math.sign(x)
  }),
  fn('CEILING', 2, 2, 'n', ([x, s]) => ceilingLegacy(x, s)),
  fn('FLOOR', 2, 2, 'n', ([x, s]) => floorLegacy(x, s)),
  fn('CEILING.MATH', 1, 3, 'n', ([x, s = 1, mode = 0]) => mathRound(x, s, mode, true)),
  fn('FLOOR.MATH', 1, 3, 'n', ([x, s = 1, mode = 0]) => mathRound(x, s, mode, false)),
  fn('CEILING.PRECISE', 1, 2, 'n', ([x, s = 1]) => mathRound(x, s, 0, true)),
  fn('ISO.CEILING', 1, 2, 'n', ([x, s = 1]) => mathRound(x, s, 0, true)),
  fn('FLOOR.PRECISE', 1, 2, 'n', ([x, s = 1]) => mathRound(x, s, 0, false)),
  fn('COMBIN', 2, 2, 'n', ([n, k]) => {
    n = Math.trunc(n)
    k = Math.trunc(k)
    return n < 0 || k < 0 || n < k ? ERR.NUM : checkNum(combin(n, k))
  }),
  fn('COMBINA', 2, 2, 'n', ([n, k]) => {
    n = Math.trunc(n)
    k = Math.trunc(k)
    if (n < 0 || k < 0 || (n === 0 && k > 0)) return ERR.NUM
    return k === 0 ? 1 : checkNum(combin(n + k - 1, k))
  }),
  fn('GCD', 1, MANY, 'r', (args, ctx) => {
    const xs = intArgs(args, ctx)
    if (isError(xs)) return xs
    return xs.reduce((a, b) => gcd2(a, b), 0)
  }),
  fn('LCM', 1, MANY, 'r', (args, ctx) => {
    const xs = intArgs(args, ctx)
    if (isError(xs)) return xs
    if (xs.some((x) => x === 0)) return 0
    return xs.reduce((a, b) => (a * b) / gcd2(a, b), 1)
  }),
  fn('MULTINOMIAL', 1, MANY, 'r', (args, ctx) => {
    const xs = intArgs(args, ctx)
    if (isError(xs)) return xs
    let r = 1
    let total = 0
    for (const x of xs) {
      for (let i = 1; i <= x; i++) r = (r * ++total) / i
    }
    return checkNum(Math.round(r))
  }),
  fn('SUM', 1, MANY, 'r', (args, ctx) => {
    const xs = numbers(args, ctx)
    return isError(xs) ? xs : sum(xs)
  }),
  fn('PRODUCT', 1, MANY, 'r', (args, ctx) => {
    const xs = numbers(args, ctx)
    if (isError(xs)) return xs
    return xs.length ? checkNum(xs.reduce((a, b) => a * b, 1)) : 0
  }),
  fn('SUMSQ', 1, MANY, 'r', (args, ctx) => {
    const xs = numbers(args, ctx)
    return isError(xs) ? xs : sum(xs.map((x) => x * x))
  }),
  fn('SUMPRODUCT', 1, MANY, 'r', (args, ctx) => {
    const arrs = args.map((a) => toArray(a, ctx))
    for (const a of arrs) if (isError(a)) return a
    const list = arrs as ArrayValue[]
    const { rows, cols } = list[0]
    if (list.some((a) => a.rows !== rows || a.cols !== cols)) return ERR.VALUE
    let total = 0
    for (let i = 0; i < rows * cols; i++) {
      let p = 1
      for (const a of list) {
        const v = a.data[i]
        if (v instanceof CellError) return v
        p *= typeof v === 'number' ? v : 0
      }
      total += p
    }
    return total
  }),
  ...(['SUMX2MY2', 'SUMX2PY2', 'SUMXMY2'] as const).map((name) =>
    fn(name, 2, 2, 'r', ([a, b], ctx) => {
      const x = toArray(a, ctx)
      const y = toArray(b, ctx)
      if (isError(x)) return x
      if (isError(y)) return y
      if (x.data.length !== y.data.length) return ERR.NA
      let s = 0
      let any = false
      for (let i = 0; i < x.data.length; i++) {
        const p = x.data[i]
        const q = y.data[i]
        if (p instanceof CellError) return p
        if (q instanceof CellError) return q
        if (typeof p !== 'number' || typeof q !== 'number') continue
        any = true
        s += name === 'SUMX2MY2' ? p * p - q * q : name === 'SUMX2PY2' ? p * p + q * q : (p - q) * (p - q)
      }
      return any ? s : ERR.DIV0
    })
  ),
  fn('SUMIF', 2, 3, 'rvr', ([range, crit, sumRange], ctx) => {
    const m = matchIndexes([[range, crit]], ctx)
    if (isError(m)) return m
    const vals = sumRange === undefined ? rangeValues(range, ctx) : resized(sumRange, m.rows, m.cols, ctx)
    if (isError(vals)) return vals
    let s = 0
    for (const i of m.idx) {
      const v = vals.data[i]
      if (v instanceof CellError) return v
      if (typeof v === 'number') s += v
    }
    return s
  }),
  fn('SUMIFS', 3, MANY, 'r|rv', ([sumRange, ...rest], ctx) => {
    if (rest.length % 2) return ERR.VALUE
    const vals = rangeValues(sumRange, ctx)
    if (isError(vals)) return vals
    const pairs: [Value, Value][] = []
    for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]])
    const m = matchIndexes(pairs, ctx, { rows: vals.rows, cols: vals.cols })
    if (isError(m)) return m
    let s = 0
    for (const i of m.idx) {
      const v = vals.data[i]
      if (v instanceof CellError) return v
      if (typeof v === 'number') s += v
    }
    return s
  }),
  fn('SUBTOTAL', 2, MANY, 'n|r', ([code, ...refs], ctx) => {
    const c = Math.trunc(code)
    const base = c > 100 ? c - 100 : c
    const name = SUBTOTAL_FN[base]
    if (!name || (c > 11 && c < 101) || c > 111) return ERR.VALUE
    if (refs.some((r: Value) => !isRef(r))) return ERR.VALUE
    const hidden = c > 100 ? 'all' : 'filtered'
    const xs = numbers(refs, ctx, { skipHidden: hidden, skipSubtotals: true })
    if (isError(xs)) return xs
    return aggregateValues(name, xs, name === 'COUNTA' ? countNonEmpty(refs, ctx, hidden) : 0)
  }),
  fn('AGGREGATE', 3, MANY, 'nn|r', ([code, opt, ...refs], ctx) => {
    const c = Math.trunc(code)
    const o = Math.trunc(opt)
    const name = AGG_FN[c]
    if (!name || o < 0 || o > 7) return ERR.VALUE
    const skipHidden = o === 1 || o === 3 || o === 5 || o === 7 ? 'all' : undefined
    const skipErrors = o === 2 || o === 3 || o === 6 || o === 7
    const skipSubtotals = o <= 3
    let k: number | undefined
    let data = refs
    if (c >= 14) {
      if (refs.length < 2) return ERR.VALUE
      const kv = refs[refs.length - 1]
      const kk = typeof kv === 'number' ? kv : isRef(kv) ? numbers([kv], ctx) : NaN
      k = typeof kk === 'number' ? kk : Array.isArray(kk) ? kk[0] : NaN
      data = refs.slice(0, -1)
    }
    const xs = numbers(data, ctx, { skipHidden, skipErrors, skipSubtotals })
    if (isError(xs)) return xs
    return aggregateValues(name, xs, name === 'COUNTA' ? countNonEmpty(data, ctx, skipHidden ?? false) : 0, k)
  }),
  fn('RAND', 0, 0, '', () => Math.random(), { volatile: true }),
  fn('RANDBETWEEN', 2, 2, 'n', ([lo, hi]) => {
    const a = Math.ceil(lo)
    const b = Math.floor(hi)
    if (a > b) return ERR.NUM
    return a + Math.floor(Math.random() * (b - a + 1))
  }, { volatile: true }),
  fn('RANDARRAY', 0, 5, 'nnnnb', ([rows = 1, cols = 1, lo = 0, hi = 1, whole = false]) => {
    rows = Math.trunc(rows)
    cols = Math.trunc(cols)
    if (rows < 1 || cols < 1) return ERR.CALC
    if (lo > hi) return ERR.VALUE
    if (rows * cols > 32_000_000) return ERR.NUM
    const data: Scalar[] = []
    for (let i = 0; i < rows * cols; i++) data.push(whole ? Math.ceil(lo) + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo))
    return makeArray(rows, cols, data)
  }, { volatile: true }),
  fn('SEQUENCE', 1, 4, 'n', ([rows, cols = 1, start = 1, step = 1]) => {
    rows = Math.trunc(rows)
    cols = Math.trunc(cols)
    if (rows < 1 || cols < 1) return ERR.CALC
    if (rows * cols > 32_000_000) return ERR.NUM
    const data: Scalar[] = new Array(rows * cols)
    for (let i = 0; i < rows * cols; i++) data[i] = start + i * step
    return makeArray(rows, cols, data)
  }),
  fn('MMULT', 2, 2, 'r', ([a, b], ctx) => {
    const x = matrix(a, ctx)
    const y = matrix(b, ctx)
    if (isError(x)) return x
    if (isError(y)) return y
    if (x[0].length !== y.length) return ERR.VALUE
    const out = x.map((row) => y[0].map((_, j) => row.reduce((s, v, k) => s + v * y[k][j], 0)))
    return fromMatrix(out)
  }),
  fn('MDETERM', 1, 1, 'r', ([a], ctx) => {
    const m = matrix(a, ctx)
    if (isError(m)) return m
    if (m.length !== m[0].length) return ERR.VALUE
    return det(m)
  }),
  fn('MINVERSE', 1, 1, 'r', ([a], ctx) => {
    const m = matrix(a, ctx)
    if (isError(m)) return m
    if (m.length !== m[0].length) return ERR.VALUE
    const inv = inverse(m)
    return inv ? fromMatrix(inv) : ERR.NUM
  }),
  fn('MUNIT', 1, 1, 'n', ([n]) => {
    n = Math.trunc(n)
    if (n < 1) return ERR.VALUE
    const data: Scalar[] = []
    for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) data.push(i === j ? 1 : 0)
    return makeArray(n, n, data)
  }),
  fn('SERIESSUM', 4, 4, 'nnnr', ([x, n, m, coef], ctx) => {
    const cs = numbers([coef], ctx)
    if (isError(cs)) return cs
    let s = 0
    cs.forEach((a, i) => (s += a * Math.pow(x, n + i * m)))
    return checkNum(s)
  }),
  fn('ROMAN', 1, 2, 'nv', ([x, form = 0]) => {
    const mode = form === true ? 0 : form === false ? 4 : typeof form === 'number' ? Math.trunc(form) : 0
    if (x < 0 || x > 3999 || mode < 0 || mode > 4) return ERR.VALUE
    return roman(Math.trunc(x), mode)
  }),
  fn('ARABIC', 1, 1, 's', ([s]) => arabic(s)),
  fn('BASE', 2, 3, 'n', ([x, radix, minLen = 0]) => {
    x = Math.trunc(x)
    radix = Math.trunc(radix)
    if (x < 0 || x >= 2 ** 53 || radix < 2 || radix > 36 || minLen < 0 || minLen > 255) return ERR.NUM
    return x.toString(radix).toUpperCase().padStart(Math.trunc(minLen), '0')
  }),
  fn('DECIMAL', 2, 2, 'sn', ([s, radix]) => {
    radix = Math.trunc(radix)
    if (radix < 2 || radix > 36 || s.length > 255) return ERR.NUM
    let t = s.trim().toLowerCase()
    if (radix === 16 && t.startsWith('0x')) t = t.slice(2)
    if (!t) return 0
    let v = 0
    for (const ch of t) {
      const d = parseInt(ch, 36)
      if (Number.isNaN(d) || d >= radix) return ERR.NUM
      v = v * radix + d
    }
    return v
  }),
  fn('PERCENTOF', 2, 2, 'r', ([a, b], ctx) => {
    const x = numbers([a], ctx)
    const y = numbers([b], ctx)
    if (isError(x)) return x
    if (isError(y)) return y
    const t = sum(y)
    return t === 0 ? ERR.DIV0 : sum(x) / t
  })
])

/** SUMIF's sum range takes the shape of the criteria range, starting at its own top-left cell. */
function resized(v: Value, rows: number, cols: number, ctx: EvalContext) {
  if (isRef(v) && v.areas.length === 1) {
    const a = v.areas[0]
    return rangeValues({ kind: 'ref', areas: [{ ...a, r1: a.r0 + rows - 1, c1: a.c0 + cols - 1 }] }, ctx)
  }
  return isArray(v) ? v : grid(v, ctx)
}

function det(m: number[][]): number {
  const n = m.length
  const a = m.map((r) => [...r])
  let d = 1
  for (let i = 0; i < n; i++) {
    let p = i
    for (let r = i + 1; r < n; r++) if (Math.abs(a[r][i]) > Math.abs(a[p][i])) p = r
    if (a[p][i] === 0) return 0
    if (p !== i) {
      ;[a[p], a[i]] = [a[i], a[p]]
      d = -d
    }
    d *= a[i][i]
    for (let r = i + 1; r < n; r++) {
      const f = a[r][i] / a[i][i]
      for (let c = i; c < n; c++) a[r][c] -= f * a[i][c]
    }
  }
  return d
}

export function inverse(m: number[][]): number[][] | null {
  const n = m.length
  const a = m.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))])
  for (let i = 0; i < n; i++) {
    let p = i
    for (let r = i + 1; r < n; r++) if (Math.abs(a[r][i]) > Math.abs(a[p][i])) p = r
    if (Math.abs(a[p][i]) < 1e-300) return null
    ;[a[p], a[i]] = [a[i], a[p]]
    const piv = a[i][i]
    for (let c = 0; c < 2 * n; c++) a[i][c] /= piv
    for (let r = 0; r < n; r++) {
      if (r === i) continue
      const f = a[r][i]
      if (f) for (let c = 0; c < 2 * n; c++) a[r][c] -= f * a[i][c]
    }
  }
  return a.map((r) => r.slice(n))
}

// ROMAN forms 0 (classic) to 4 (most simplified), the algorithm LibreOffice uses to match Excel.
function roman(n: number, mode: number): string {
  const chars = ['M', 'D', 'C', 'L', 'X', 'V', 'I']
  const vals = [1000, 500, 100, 50, 10, 5, 1]
  const max = vals.length - 1
  let out = ''
  let v = n
  for (let i = 0; i <= max / 2; i++) {
    let idx = 2 * i
    const digit = Math.floor(v / vals[idx])
    if (digit % 5 === 4) {
      const idx2 = digit === 4 ? idx - 1 : idx - 2
      let steps = 0
      while (steps < mode && idx < max) {
        steps++
        if (vals[idx2] - vals[idx + 1] <= v) idx++
        else steps = mode
      }
      out += chars[idx] + chars[idx2]
      v = v + vals[idx] - vals[idx2]
    } else {
      if (digit > 4) out += chars[idx - 1]
      out += chars[idx].repeat(digit % 5)
      v %= vals[idx]
    }
  }
  return out
}

function arabic(s: string): number | CellError {
  let t = s.trim().toUpperCase()
  if (t.length > 255) return ERR.VALUE
  let sign = 1
  if (t.startsWith('-')) {
    sign = -1
    t = t.slice(1)
  }
  const map: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 }
  let total = 0
  for (let i = 0; i < t.length; i++) {
    const v = map[t[i]]
    if (!v) return ERR.VALUE
    const next = map[t[i + 1]] ?? 0
    total += v < next ? -v : v
  }
  return sign * total
}

export { mean }
