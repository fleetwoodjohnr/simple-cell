// Statistical functions, and the pre-2010 compatibility names that map onto them.
import { defineFunctions, toArray, type EvalContext } from '../evaluate'
import { CellError, ERR, isArray, isError, isRef, makeArray, toNumber, type ArrayValue, type Scalar, type Value } from '../values'
import { allValues, checkNum, fn, matchIndexes, MANY, mean, numbers, rangeValues, sum, variance } from './helpers'
import type { FnSpec } from '../evaluate'
import {
  betaI,
  betaPdf,
  binomCdf,
  binomPmf,
  chiCdf,
  chiPdf,
  erf,
  fCdf,
  fPdf,
  gamma,
  gammaCdf,
  gammaPdf,
  invert,
  kth,
  lnGamma,
  median,
  modes,
  normCdf,
  normInv,
  normPdf,
  percentileExc,
  percentileInc,
  sorted,
  tCdf,
  tPdf
} from './stats-core'
import { combin, inverse } from './math'

const nums = (args: Value[], ctx: EvalContext) => numbers(args, ctx)

/** Paired numeric data (CORREL, SLOPE…): both arrays the same size; pairs with a non-number on either side are skipped. */
function pairs(a: Value, b: Value, ctx: EvalContext): { x: number[]; y: number[] } | CellError {
  const A = toArray(a, ctx)
  const B = toArray(b, ctx)
  if (isError(A)) return A
  if (isError(B)) return B
  if (A.data.length !== B.data.length) return ERR.NA
  const x: number[] = []
  const y: number[] = []
  for (let i = 0; i < A.data.length; i++) {
    const p = A.data[i]
    const q = B.data[i]
    if (p instanceof CellError) return p
    if (q instanceof CellError) return q
    if (typeof p === 'number' && typeof q === 'number') {
      x.push(p)
      y.push(q)
    }
  }
  return { x, y }
}

function covariance(x: number[], y: number[], sample: boolean): number | CellError {
  const n = x.length
  if (n < (sample ? 2 : 1)) return ERR.DIV0
  const mx = mean(x)
  const my = mean(y)
  let s = 0
  for (let i = 0; i < n; i++) s += (x[i] - mx) * (y[i] - my)
  return s / (sample ? n - 1 : n)
}

function correl(x: number[], y: number[]): number | CellError {
  if (x.length < 2) return ERR.DIV0
  const mx = mean(x)
  const my = mean(y)
  let sxy = 0
  let sxx = 0
  let syy = 0
  for (let i = 0; i < x.length; i++) {
    sxy += (x[i] - mx) * (y[i] - my)
    sxx += (x[i] - mx) ** 2
    syy += (y[i] - my) ** 2
  }
  if (sxx === 0 || syy === 0) return ERR.DIV0
  return sxy / Math.sqrt(sxx * syy)
}

function slopeIntercept(ys: number[], xs: number[]): { slope: number; intercept: number } | CellError {
  if (xs.length < 1) return ERR.DIV0
  const mx = mean(xs)
  const my = mean(ys)
  let sxy = 0
  let sxx = 0
  for (let i = 0; i < xs.length; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my)
    sxx += (xs[i] - mx) ** 2
  }
  if (sxx === 0) return ERR.DIV0
  const slope = sxy / sxx
  return { slope, intercept: my - slope * mx }
}

// ---------- counting ----------

function count(args: Value[], ctx: EvalContext, a: boolean): number {
  let n = 0
  for (const v of args) {
    if (v === undefined) continue
    if (isRef(v) || isArray(v)) {
      const vals = allValues([v], ctx)
      if (isError(vals)) continue
      for (const x of vals) if (a ? x !== null : typeof x === 'number') n++
    } else if (a) n++
    else if (typeof v === 'number' || typeof v === 'boolean') n++
    else if (typeof v === 'string' && !isError(toNumber(v))) n++
  }
  return n
}

// ---------- regression (LINEST, LOGEST, TREND, GROWTH) ----------

interface Fit {
  coef: number[] // b0 (intercept), b1..bk
  se?: number[]
  r2: number
  sey: number
  F: number
  df: number
  ssreg: number
  ssresid: number
}

/** Ordinary least squares y = b0 + Σ bi·xi (b0 = 0 when const is false), via the normal equations. */
function ols(y: number[], X: number[][], withConst: boolean, stats: boolean): Fit | CellError {
  const n = y.length
  const k = X[0]?.length ?? 0
  const cols = withConst ? k + 1 : k
  const A = X.map((row) => (withConst ? [1, ...row] : row))
  if (n < cols) return ERR.NUM
  const AtA = Array.from({ length: cols }, (_, i) => Array.from({ length: cols }, (_, j) => A.reduce((s, row) => s + row[i] * row[j], 0)))
  const Aty = Array.from({ length: cols }, (_, i) => A.reduce((s, row, r) => s + row[i] * y[r], 0))
  const inv = inverse(AtA)
  if (!inv) return ERR.NUM
  const b = inv.map((row) => row.reduce((s, v, j) => s + v * Aty[j], 0))
  const coef = withConst ? b : [0, ...b]
  const fitted = A.map((row) => row.reduce((s, v, j) => s + v * b[j], 0))
  const my = withConst ? mean(y) : 0
  const ssresid = y.reduce((s, v, i) => s + (v - fitted[i]) ** 2, 0)
  const sstot = y.reduce((s, v) => s + (v - my) ** 2, 0)
  const ssreg = sstot - ssresid
  const df = n - cols
  const r2 = sstot === 0 ? 1 : ssreg / sstot
  const sey = df > 0 ? Math.sqrt(ssresid / df) : 0
  const F = df > 0 && cols - (withConst ? 1 : 0) > 0 ? ssreg / (cols - (withConst ? 1 : 0)) / (ssresid / df) : 0
  let se: number[] | undefined
  if (stats) {
    const s2 = df > 0 ? ssresid / df : 0
    const raw = inv.map((row, i) => Math.sqrt(Math.max(0, row[i] * s2)))
    se = withConst ? raw : [NaN, ...raw]
  }
  return { coef, se, r2, sey, F, df, ssreg, ssresid }
}

/** known_y / known_x arguments as y vector and X matrix (columns = variables). */
function regressionData(yv: Value, xv: Value | undefined, ctx: EvalContext): { y: number[]; X: number[][]; byRow: boolean } | CellError {
  const Y = toArray(yv, ctx)
  if (isError(Y)) return Y
  const y: number[] = []
  for (const v of Y.data) {
    if (typeof v !== 'number') return ERR.VALUE
    y.push(v)
  }
  const n = y.length
  if (xv === undefined || xv === null) return { y, X: y.map((_, i) => [i + 1]), byRow: Y.rows === 1 && Y.cols > 1 }
  const Xa = toArray(xv, ctx)
  if (isError(Xa)) return Xa
  for (const v of Xa.data) if (typeof v !== 'number') return ERR.VALUE
  const vals = Xa.data as number[]
  // One variable with the same shape, or several: as columns (y vertical) or rows (y horizontal).
  if (Xa.rows * Xa.cols === n) return { y, X: vals.map((v) => [v]), byRow: Y.rows === 1 && Y.cols > 1 }
  if (Y.cols === 1 && Xa.rows === n) return { y, X: Array.from({ length: n }, (_, r) => vals.slice(r * Xa.cols, (r + 1) * Xa.cols)), byRow: false }
  if (Y.rows === 1 && Xa.cols === n) return { y, X: Array.from({ length: n }, (_, c) => Array.from({ length: Xa.rows }, (_, r) => vals[r * Xa.cols + c])), byRow: true }
  return ERR.REF
}

function linest(yv: Value, xv: Value | undefined, withConst: boolean, stats: boolean, ctx: EvalContext, log: boolean): Value {
  const d = regressionData(yv, xv, ctx)
  if (isError(d)) return d
  let y = d.y
  if (log) {
    if (y.some((v) => v <= 0)) return ERR.NUM
    y = y.map(Math.log)
  }
  const fit = ols(y, d.X, withConst, stats)
  if (isError(fit)) return fit
  const k = d.X[0].length
  // Excel orders coefficients mk, …, m1, b
  const row1 = [...fit.coef.slice(1).reverse(), fit.coef[0]].map((v) => (log ? Math.exp(v) : v))
  if (!stats) return makeArray(1, k + 1, row1)
  const se = fit.se!
  const row2: Scalar[] = [...se.slice(1).reverse(), withConst ? se[0] : ERR.NA]
  const na = Array(Math.max(0, k - 1)).fill(ERR.NA) as Scalar[]
  const data: Scalar[] = [...row1, ...row2, fit.r2, fit.sey, ...na, fit.F, fit.df, ...na, fit.ssreg, fit.ssresid, ...na]
  return makeArray(5, k + 1, data)
}

function trend(yv: Value, xv: Value | undefined, newv: Value | undefined, withConst: boolean, ctx: EvalContext, log: boolean): Value {
  const d = regressionData(yv, xv, ctx)
  if (isError(d)) return d
  let y = d.y
  if (log) {
    if (y.some((v) => v <= 0)) return ERR.NUM
    y = y.map(Math.log)
  }
  const fit = ols(y, d.X, withConst, false)
  if (isError(fit)) return fit
  const k = d.X[0].length
  let points: number[][]
  let shape: { rows: number; cols: number }
  if (newv === undefined || newv === null) {
    points = d.X
    const Y = toArray(yv, ctx) as ArrayValue
    shape = { rows: Y.rows, cols: Y.cols }
  } else {
    const N = toArray(newv, ctx)
    if (isError(N)) return N
    if (N.data.some((v) => typeof v !== 'number')) return ERR.VALUE
    const vals = N.data as number[]
    if (k === 1) {
      points = vals.map((v) => [v])
      shape = { rows: N.rows, cols: N.cols }
    } else if (N.cols === k) {
      points = Array.from({ length: N.rows }, (_, r) => vals.slice(r * k, (r + 1) * k))
      shape = { rows: N.rows, cols: 1 }
    } else if (N.rows === k) {
      points = Array.from({ length: N.cols }, (_, c) => Array.from({ length: k }, (_, r) => vals[r * N.cols + c]))
      shape = { rows: 1, cols: N.cols }
    } else return ERR.REF
  }
  const out = points.map((p) => {
    const v = fit.coef[0] + p.reduce((s, x, i) => s + x * fit.coef[i + 1], 0)
    return log ? Math.exp(v) : v
  })
  return makeArray(shape.rows, shape.cols, out)
}

// ---------- ranking ----------

function rank(x: number, ref: Value, desc: boolean, avg: boolean, ctx: EvalContext): Scalar {
  const xs = numbers([ref], ctx)
  if (isError(xs)) return xs
  if (!xs.includes(x)) return ERR.NA
  const better = xs.filter((v) => (desc ? v > x : v < x)).length
  if (!avg) return better + 1
  const ties = xs.filter((v) => v === x).length
  return better + (ties + 1) / 2
}

function percentRank(xs: number[], x: number, sig: number, exc: boolean): Scalar {
  if (!xs.length || sig < 1) return ERR.NUM
  const s = sorted(xs)
  const n = s.length
  if (x < s[0] || x > s[n - 1]) return ERR.NA
  let pos: number
  const i = s.findIndex((v) => v >= x)
  if (s[i] === x) {
    // first occurrence rank
    pos = i
  } else {
    // interpolate between s[i-1] and s[i]
    pos = i - 1 + (x - s[i - 1]) / (s[i] - s[i - 1])
  }
  const r = exc ? (pos + 1) / (n + 1) : n === 1 ? 1 : pos / (n - 1)
  const p = Math.pow(10, Math.trunc(sig))
  return Math.floor(r * p + 1e-9) / p
}

// ---------- distributions ----------

const cumOrPdf = (cum: boolean, c: number, p: number) => checkNum(cum ? c : p)

function hypgeom(s: number, n: number, K: number, N: number): number {
  return Math.exp(lnC(K, s) + lnC(N - K, n - s) - lnC(N, n))
}
const lnC = (n: number, k: number) => lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1)

function tTest(a: number[], b: number[], tails: number, type: number): Scalar {
  if (tails !== 1 && tails !== 2) return ERR.NUM
  let t: number
  let df: number
  if (type === 1) {
    if (a.length !== b.length || a.length < 2) return ERR.NA
    const d = a.map((v, i) => v - b[i])
    const sd = Math.sqrt(variance(d, true) as number)
    if (!sd) return ERR.DIV0
    t = mean(d) / (sd / Math.sqrt(d.length))
    df = d.length - 1
  } else {
    if (a.length < 2 || b.length < 2) return ERR.DIV0
    const va = variance(a, true) as number
    const vb = variance(b, true) as number
    const na = a.length
    const nb = b.length
    if (type === 2) {
      df = na + nb - 2
      const sp = ((na - 1) * va + (nb - 1) * vb) / df
      t = (mean(a) - mean(b)) / Math.sqrt(sp * (1 / na + 1 / nb))
    } else if (type === 3) {
      const se2 = va / na + vb / nb
      t = (mean(a) - mean(b)) / Math.sqrt(se2)
      df = (se2 * se2) / ((va / na) ** 2 / (na - 1) + (vb / nb) ** 2 / (nb - 1))
    } else return ERR.NUM
  }
  const p = 1 - tCdf(Math.abs(t), df)
  return tails === 2 ? 2 * p : p
}

function chiTest(actual: ArrayValue, expected: ArrayValue): Scalar {
  if (actual.rows !== expected.rows || actual.cols !== expected.cols) return ERR.NA
  let chi = 0
  for (let i = 0; i < actual.data.length; i++) {
    const a = actual.data[i]
    const e = expected.data[i]
    if (typeof a !== 'number' || typeof e !== 'number') continue
    if (e === 0) return ERR.DIV0
    chi += (a - e) ** 2 / e
  }
  const df = actual.rows > 1 && actual.cols > 1 ? (actual.rows - 1) * (actual.cols - 1) : Math.max(actual.rows, actual.cols) - 1
  if (df < 1) return ERR.NA
  return 1 - chiCdf(chi, df)
}

function fTest(a: number[], b: number[]): Scalar {
  if (a.length < 2 || b.length < 2) return ERR.DIV0
  const va = variance(a, true) as number
  const vb = variance(b, true) as number
  if (!va || !vb) return ERR.DIV0
  const f = va / vb
  const d1 = a.length - 1
  const d2 = b.length - 1
  const p = fCdf(f, d1, d2)
  return 2 * Math.min(p, 1 - p)
}

function zTest(xs: number[], x: number, sigma?: number): Scalar {
  if (xs.length < 1) return ERR.NA
  const s = sigma ?? Math.sqrt(variance(xs, true) as number)
  if (!s) return ERR.DIV0
  return 1 - normCdf((mean(xs) - x) / (s / Math.sqrt(xs.length)))
}

function binomInv(n: number, p: number, alpha: number): Scalar {
  n = Math.trunc(n)
  if (n < 0 || p < 0 || p > 1 || alpha <= 0 || alpha >= 1) return ERR.NUM
  for (let k = 0; k <= n; k++) if (binomCdf(k, n, p) >= alpha - 1e-12) return k
  return n
}

const forecastLinear = (x: number, yv: Value, xv: Value, ctx: EvalContext): Scalar => {
  const p = pairs(yv, xv, ctx)
  if (isError(p)) return p
  const si = slopeIntercept(p.x, p.y)
  if (isError(si)) return si
  return si.intercept + si.slope * x
}

const stdevS = (xs: number[]) => {
  const v = variance(xs, true)
  return isError(v) ? v : Math.sqrt(v)
}
const stdevP = (xs: number[]) => {
  const v = variance(xs, false)
  return isError(v) ? v : Math.sqrt(v)
}

const A = (args: Value[], ctx: EvalContext) => numbers(args, ctx, { a: true })

function withNums(f: (xs: number[]) => Scalar | number, a = false) {
  return (args: Value[], ctx: EvalContext) => {
    const xs = a ? A(args, ctx) : nums(args, ctx)
    return isError(xs) ? xs : f(xs)
  }
}

const statFns: FnSpec[] = [
  fn('AVERAGE', 1, MANY, 'r', withNums((xs) => (xs.length ? mean(xs) : ERR.DIV0))),
  fn('AVERAGEA', 1, MANY, 'r', withNums((xs) => (xs.length ? mean(xs) : ERR.DIV0), true)),
  fn('AVEDEV', 1, MANY, 'r', withNums((xs) => {
    if (!xs.length) return ERR.NUM
    const m = mean(xs)
    return mean(xs.map((x) => Math.abs(x - m)))
  })),
  fn('DEVSQ', 1, MANY, 'r', withNums((xs) => {
    if (!xs.length) return ERR.NUM
    const m = mean(xs)
    return sum(xs.map((x) => (x - m) ** 2))
  })),
  fn('GEOMEAN', 1, MANY, 'r', withNums((xs) => (!xs.length || xs.some((x) => x <= 0) ? ERR.NUM : Math.exp(mean(xs.map(Math.log)))))),
  fn('HARMEAN', 1, MANY, 'r', withNums((xs) => (!xs.length || xs.some((x) => x <= 0) ? ERR.NUM : xs.length / sum(xs.map((x) => 1 / x))))),
  fn('COUNT', 1, MANY, 'r', (args, ctx) => count(args, ctx, false)),
  fn('COUNTA', 1, MANY, 'r', (args, ctx) => count(args, ctx, true)),
  fn('COUNTBLANK', 1, 1, 'r', ([v], ctx) => {
    const a = rangeValues(v, ctx)
    if (isError(a)) return a
    return a.data.filter((x) => x === null || x === '').length
  }),
  fn('COUNTIF', 2, 2, 'rv', ([range, crit], ctx) => {
    const m = matchIndexes([[range, crit]], ctx)
    return isError(m) ? m : m.idx.length
  }),
  fn('COUNTIFS', 2, MANY, '|rv', (args, ctx) => {
    if (args.length % 2) return ERR.VALUE
    const ps: [Value, Value][] = []
    for (let i = 0; i < args.length; i += 2) ps.push([args[i], args[i + 1]])
    const m = matchIndexes(ps, ctx)
    return isError(m) ? m : m.idx.length
  }),
  fn('AVERAGEIF', 2, 3, 'rvr', ([range, crit, avgRange], ctx) => {
    const m = matchIndexes([[range, crit]], ctx)
    if (isError(m)) return m
    let vals = rangeValues(avgRange ?? range, ctx)
    if (isError(vals)) return vals
    if (avgRange !== undefined && isRef(avgRange)) {
      const a = avgRange.areas[0]
      vals = rangeValues({ kind: 'ref', areas: [{ ...a, r1: a.r0 + m.rows - 1, c1: a.c0 + m.cols - 1 }] }, ctx) as ArrayValue
    }
    const xs: number[] = []
    for (const i of m.idx) {
      const v = (vals as ArrayValue).data[i]
      if (v instanceof CellError) return v
      if (typeof v === 'number') xs.push(v)
    }
    return xs.length ? mean(xs) : ERR.DIV0
  }),
  ...(['AVERAGEIFS', 'MAXIFS', 'MINIFS'] as const).map((name) =>
    fn(name, 3, MANY, 'r|rv', ([target, ...rest], ctx) => {
      if (rest.length % 2) return ERR.VALUE
      const vals = rangeValues(target, ctx)
      if (isError(vals)) return vals
      const ps: [Value, Value][] = []
      for (let i = 0; i < rest.length; i += 2) ps.push([rest[i], rest[i + 1]])
      const m = matchIndexes(ps, ctx, { rows: vals.rows, cols: vals.cols })
      if (isError(m)) return m
      const xs: number[] = []
      for (const i of m.idx) {
        const v = vals.data[i]
        if (v instanceof CellError) return v
        if (typeof v === 'number') xs.push(v)
      }
      if (name === 'AVERAGEIFS') return xs.length ? mean(xs) : ERR.DIV0
      if (!xs.length) return 0
      return name === 'MAXIFS' ? Math.max(...xs) : Math.min(...xs)
    })
  ),
  fn('MAX', 1, MANY, 'r', withNums((xs) => (xs.length ? xs.reduce((a, b) => (b > a ? b : a), -Infinity) : 0))),
  fn('MIN', 1, MANY, 'r', withNums((xs) => (xs.length ? xs.reduce((a, b) => (b < a ? b : a), Infinity) : 0))),
  fn('MAXA', 1, MANY, 'r', withNums((xs) => (xs.length ? xs.reduce((a, b) => (b > a ? b : a), -Infinity) : 0), true)),
  fn('MINA', 1, MANY, 'r', withNums((xs) => (xs.length ? xs.reduce((a, b) => (b < a ? b : a), Infinity) : 0), true)),
  fn('MEDIAN', 1, MANY, 'r', withNums((xs) => median(xs))),
  fn('MODE.SNGL', 1, MANY, 'r', withNums((xs) => modes(xs)[0] ?? ERR.NA)),
  fn('MODE', 1, MANY, 'r', withNums((xs) => modes(xs)[0] ?? ERR.NA)),
  fn('MODE.MULT', 1, MANY, 'r', (args, ctx) => {
    const xs = nums(args, ctx)
    if (isError(xs)) return xs
    const m = modes(xs)
    return m.length ? makeArray(m.length, 1, m) : ERR.NA
  }),
  fn('LARGE', 2, 2, 'rn', ([a, k], ctx) => {
    const xs = nums([a], ctx)
    return isError(xs) ? xs : kth(xs, k, true)
  }),
  fn('SMALL', 2, 2, 'rn', ([a, k], ctx) => {
    const xs = nums([a], ctx)
    return isError(xs) ? xs : kth(xs, k, false)
  }),
  ...(['PERCENTILE.INC', 'PERCENTILE'] as const).map((name) =>
    fn(name, 2, 2, 'rn', ([a, k], ctx) => {
      const xs = nums([a], ctx)
      return isError(xs) ? xs : percentileInc(xs, k)
    })
  ),
  fn('PERCENTILE.EXC', 2, 2, 'rn', ([a, k], ctx) => {
    const xs = nums([a], ctx)
    return isError(xs) ? xs : percentileExc(xs, k)
  }),
  ...(['QUARTILE.INC', 'QUARTILE'] as const).map((name) =>
    fn(name, 2, 2, 'rn', ([a, q], ctx) => {
      const xs = nums([a], ctx)
      if (isError(xs)) return xs
      const k = Math.trunc(q)
      return k < 0 || k > 4 ? ERR.NUM : percentileInc(xs, k / 4)
    })
  ),
  fn('QUARTILE.EXC', 2, 2, 'rn', ([a, q], ctx) => {
    const xs = nums([a], ctx)
    if (isError(xs)) return xs
    const k = Math.trunc(q)
    return k <= 0 || k >= 4 ? ERR.NUM : percentileExc(xs, k / 4)
  }),
  ...(['PERCENTRANK.INC', 'PERCENTRANK'] as const).map((name) =>
    fn(name, 2, 3, 'rnn', ([a, x, sig = 3], ctx) => {
      const xs = nums([a], ctx)
      return isError(xs) ? xs : percentRank(xs, x, sig, false)
    })
  ),
  fn('PERCENTRANK.EXC', 2, 3, 'rnn', ([a, x, sig = 3], ctx) => {
    const xs = nums([a], ctx)
    return isError(xs) ? xs : percentRank(xs, x, sig, true)
  }),
  ...(['RANK.EQ', 'RANK'] as const).map((name) => fn(name, 2, 3, 'nrn', ([x, ref, order = 0], ctx) => rank(x, ref, order === 0, false, ctx))),
  fn('RANK.AVG', 2, 3, 'nrn', ([x, ref, order = 0], ctx) => rank(x, ref, order === 0, true, ctx)),
  ...(['STDEV.S', 'STDEV'] as const).map((n) => fn(n, 1, MANY, 'r', withNums(stdevS))),
  ...(['STDEV.P', 'STDEVP'] as const).map((n) => fn(n, 1, MANY, 'r', withNums(stdevP))),
  fn('STDEVA', 1, MANY, 'r', withNums(stdevS, true)),
  fn('STDEVPA', 1, MANY, 'r', withNums(stdevP, true)),
  ...(['VAR.S', 'VAR'] as const).map((n) => fn(n, 1, MANY, 'r', withNums((xs) => variance(xs, true)))),
  ...(['VAR.P', 'VARP'] as const).map((n) => fn(n, 1, MANY, 'r', withNums((xs) => variance(xs, false)))),
  fn('VARA', 1, MANY, 'r', withNums((xs) => variance(xs, true), true)),
  fn('VARPA', 1, MANY, 'r', withNums((xs) => variance(xs, false), true)),
  fn('KURT', 1, MANY, 'r', withNums((xs) => {
    const n = xs.length
    if (n < 4) return ERR.DIV0
    const m = mean(xs)
    const s = Math.sqrt(variance(xs, true) as number)
    if (!s) return ERR.DIV0
    const k = xs.reduce((a, x) => a + ((x - m) / s) ** 4, 0)
    return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * k - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
  })),
  fn('SKEW', 1, MANY, 'r', withNums((xs) => {
    const n = xs.length
    if (n < 3) return ERR.DIV0
    const m = mean(xs)
    const s = Math.sqrt(variance(xs, true) as number)
    if (!s) return ERR.DIV0
    return (n / ((n - 1) * (n - 2))) * xs.reduce((a, x) => a + ((x - m) / s) ** 3, 0)
  })),
  fn('SKEW.P', 1, MANY, 'r', withNums((xs) => {
    const n = xs.length
    if (n < 3) return ERR.DIV0
    const m = mean(xs)
    const s = Math.sqrt(variance(xs, false) as number)
    if (!s) return ERR.DIV0
    return xs.reduce((a, x) => a + ((x - m) / s) ** 3, 0) / n
  })),
  fn('TRIMMEAN', 2, 2, 'rn', ([a, pct], ctx) => {
    const xs = nums([a], ctx)
    if (isError(xs)) return xs
    if (pct < 0 || pct >= 1 || !xs.length) return ERR.NUM
    const cut = Math.floor((xs.length * pct) / 2)
    const s = sorted(xs).slice(cut, xs.length - cut)
    return mean(s)
  }),
  fn('STANDARDIZE', 3, 3, 'n', ([x, m, s]) => (s <= 0 ? ERR.NUM : (x - m) / s)),
  fn('FISHER', 1, 1, 'n', ([x]) => (x <= -1 || x >= 1 ? ERR.NUM : 0.5 * Math.log((1 + x) / (1 - x)))),
  fn('FISHERINV', 1, 1, 'n', ([y]) => Math.tanh(y)),
  fn('GAMMA', 1, 1, 'n', ([x]) => (x === 0 || (Number.isInteger(x) && x < 0) ? ERR.NUM : checkNum(gamma(x)))),
  ...(['GAMMALN', 'GAMMALN.PRECISE'] as const).map((n) => fn(n, 1, 1, 'n', ([x]) => (x <= 0 ? ERR.NUM : lnGamma(x)))),
  fn('GAUSS', 1, 1, 'n', ([z]) => normCdf(z) - 0.5),
  fn('PHI', 1, 1, 'n', ([x]) => normPdf(x)),
  fn('PERMUT', 2, 2, 'n', ([n, k]) => {
    n = Math.trunc(n)
    k = Math.trunc(k)
    if (n < 0 || k < 0 || n < k) return ERR.NUM
    let r = 1
    for (let i = 0; i < k; i++) r *= n - i
    return checkNum(r)
  }),
  fn('PERMUTATIONA', 2, 2, 'n', ([n, k]) => (n < 0 || k < 0 ? ERR.NUM : checkNum(Math.pow(Math.trunc(n), Math.trunc(k))))),
  // paired
  fn('CORREL', 2, 2, 'r', ([a, b], ctx) => {
    const p = pairs(a, b, ctx)
    return isError(p) ? p : correl(p.x, p.y)
  }),
  fn('PEARSON', 2, 2, 'r', ([a, b], ctx) => {
    const p = pairs(a, b, ctx)
    return isError(p) ? p : correl(p.x, p.y)
  }),
  fn('RSQ', 2, 2, 'r', ([a, b], ctx) => {
    const p = pairs(a, b, ctx)
    if (isError(p)) return p
    const r = correl(p.x, p.y)
    return isError(r) ? r : r * r
  }),
  ...(['COVARIANCE.P', 'COVAR'] as const).map((n) =>
    fn(n, 2, 2, 'r', ([a, b], ctx) => {
      const p = pairs(a, b, ctx)
      return isError(p) ? p : covariance(p.x, p.y, false)
    })
  ),
  fn('COVARIANCE.S', 2, 2, 'r', ([a, b], ctx) => {
    const p = pairs(a, b, ctx)
    return isError(p) ? p : covariance(p.x, p.y, true)
  }),
  fn('SLOPE', 2, 2, 'r', ([y, x], ctx) => {
    const p = pairs(y, x, ctx)
    if (isError(p)) return p
    const r = slopeIntercept(p.x, p.y)
    return isError(r) ? r : r.slope
  }),
  fn('INTERCEPT', 2, 2, 'r', ([y, x], ctx) => {
    const p = pairs(y, x, ctx)
    if (isError(p)) return p
    const r = slopeIntercept(p.x, p.y)
    return isError(r) ? r : r.intercept
  }),
  fn('STEYX', 2, 2, 'r', ([y, x], ctx) => {
    const p = pairs(y, x, ctx)
    if (isError(p)) return p
    const n = p.x.length
    if (n < 3) return ERR.DIV0
    const r = slopeIntercept(p.x, p.y)
    if (isError(r)) return r
    const ss = p.x.reduce((s, ys, i) => s + (ys - (r.intercept + r.slope * p.y[i])) ** 2, 0)
    return Math.sqrt(ss / (n - 2))
  }),
  ...(['FORECAST', 'FORECAST.LINEAR'] as const).map((n) => fn(n, 3, 3, 'nrr', ([x, y, xs], ctx) => forecastLinear(x, y, xs, ctx))),
  fn('LINEST', 1, 4, 'rrbb', ([y, x, c = true, s = false], ctx) => linest(y, x, c, s, ctx, false)),
  fn('LOGEST', 1, 4, 'rrbb', ([y, x, c = true, s = false], ctx) => linest(y, x, c, s, ctx, true)),
  fn('TREND', 1, 4, 'rrrb', ([y, x, nx, c = true], ctx) => trend(y, x, nx, c, ctx, false)),
  fn('GROWTH', 1, 4, 'rrrb', ([y, x, nx, c = true], ctx) => trend(y, x, nx, c, ctx, true)),
  fn('FREQUENCY', 2, 2, 'r', ([data, bins], ctx) => {
    const xs = nums([data], ctx)
    const bs = nums([bins], ctx)
    if (isError(xs)) return xs
    if (isError(bs)) return bs
    const order = bs.map((b, i) => [b, i]).sort((a, b) => a[0] - b[0])
    const counts = new Array(bs.length + 1).fill(0)
    for (const x of xs) {
      const k = order.findIndex(([b]) => x <= b)
      counts[k < 0 ? bs.length : order[k][1]]++
    }
    return makeArray(counts.length, 1, counts)
  }),
  fn('PROB', 3, 4, 'rrnn', ([xr, pr, lo, hi], ctx) => {
    const p = pairs(xr, pr, ctx)
    if (isError(p)) return p
    if (p.y.some((v) => v < 0 || v > 1) || Math.abs(sum(p.y) - 1) > 1e-9) return ERR.NUM
    const up = hi ?? lo
    return sum(p.x.map((x, i) => (x >= lo && x <= up ? p.y[i] : 0)))
  }),
  // distributions
  fn('NORM.DIST', 4, 4, 'nnnb', ([x, m, s, cum]) => (s <= 0 ? ERR.NUM : cumOrPdf(cum, normCdf((x - m) / s), normPdf((x - m) / s) / s))),
  fn('NORMDIST', 4, 4, 'nnnb', ([x, m, s, cum]) => (s <= 0 ? ERR.NUM : cumOrPdf(cum, normCdf((x - m) / s), normPdf((x - m) / s) / s))),
  ...(['NORM.INV', 'NORMINV'] as const).map((n) => fn(n, 3, 3, 'n', ([p, m, s]) => (p <= 0 || p >= 1 || s <= 0 ? ERR.NUM : m + s * normInv(p)))),
  fn('NORM.S.DIST', 2, 2, 'nb', ([z, cum]) => cumOrPdf(cum, normCdf(z), normPdf(z))),
  fn('NORMSDIST', 1, 1, 'n', ([z]) => normCdf(z)),
  ...(['NORM.S.INV', 'NORMSINV'] as const).map((n) => fn(n, 1, 1, 'n', ([p]) => (p <= 0 || p >= 1 ? ERR.NUM : normInv(p)))),
  ...(['LOGNORM.DIST'] as const).map((n) => fn(n, 4, 4, 'nnnb', ([x, m, s, cum]) => (x <= 0 || s <= 0 ? ERR.NUM : cumOrPdf(cum, normCdf((Math.log(x) - m) / s), normPdf((Math.log(x) - m) / s) / (x * s))))),
  fn('LOGNORMDIST', 3, 3, 'n', ([x, m, s]) => (x <= 0 || s <= 0 ? ERR.NUM : normCdf((Math.log(x) - m) / s))),
  ...(['LOGNORM.INV', 'LOGINV'] as const).map((n) => fn(n, 3, 3, 'n', ([p, m, s]) => (p <= 0 || p >= 1 || s <= 0 ? ERR.NUM : Math.exp(m + s * normInv(p))))),
  fn('EXPON.DIST', 3, 3, 'nnb', ([x, l, cum]) => (x < 0 || l <= 0 ? ERR.NUM : cumOrPdf(cum, 1 - Math.exp(-l * x), l * Math.exp(-l * x)))),
  fn('EXPONDIST', 3, 3, 'nnb', ([x, l, cum]) => (x < 0 || l <= 0 ? ERR.NUM : cumOrPdf(cum, 1 - Math.exp(-l * x), l * Math.exp(-l * x)))),
  ...(['WEIBULL.DIST', 'WEIBULL'] as const).map((n) =>
    fn(n, 4, 4, 'nnnb', ([x, a, b, cum]) => (x < 0 || a <= 0 || b <= 0 ? ERR.NUM : cumOrPdf(cum, 1 - Math.exp(-((x / b) ** a)), (a / b ** a) * x ** (a - 1) * Math.exp(-((x / b) ** a)))))
  ),
  ...(['GAMMA.DIST', 'GAMMADIST'] as const).map((n) => fn(n, 4, 4, 'nnnb', ([x, a, b, cum]) => (x < 0 || a <= 0 || b <= 0 ? ERR.NUM : cumOrPdf(cum, gammaCdf(x, a, b), gammaPdf(x, a, b))))),
  ...(['GAMMA.INV', 'GAMMAINV'] as const).map((n) => fn(n, 3, 3, 'n', ([p, a, b]) => (p < 0 || p > 1 || a <= 0 || b <= 0 ? ERR.NUM : p === 0 ? 0 : invert((x) => gammaCdf(x, a, b), p, 0, a * b * 2 + 10)))),
  fn('BETA.DIST', 4, 6, 'nnnbnn', ([x, a, b, cum, lo = 0, hi = 1]) => {
    if (a <= 0 || b <= 0 || x < lo || x > hi || lo === hi) return ERR.NUM
    const t = (x - lo) / (hi - lo)
    return cumOrPdf(cum, betaI(t, a, b), betaPdf(t, a, b) / (hi - lo))
  }),
  fn('BETADIST', 3, 5, 'n', ([x, a, b, lo = 0, hi = 1]) => (a <= 0 || b <= 0 || x < lo || x > hi || lo === hi ? ERR.NUM : betaI((x - lo) / (hi - lo), a, b))),
  ...(['BETA.INV', 'BETAINV'] as const).map((n) =>
    fn(n, 3, 5, 'n', ([p, a, b, lo = 0, hi = 1]) => (p <= 0 || p > 1 || a <= 0 || b <= 0 || lo >= hi ? ERR.NUM : lo + (hi - lo) * invert((t) => betaI(t, a, b), p, 0, 1)))
  ),
  fn('BINOM.DIST', 4, 4, 'nnnb', ([k, n, p, cum]) => {
    k = Math.trunc(k)
    n = Math.trunc(n)
    if (k < 0 || k > n || p < 0 || p > 1) return ERR.NUM
    if (p === 0 || p === 1) return cum ? (p === 0 ? 1 : k === n ? 1 : 0) : (p === 0 ? (k === 0 ? 1 : 0) : k === n ? 1 : 0)
    return cumOrPdf(cum, binomCdf(k, n, p), binomPmf(k, n, p))
  }),
  fn('BINOMDIST', 4, 4, 'nnnb', ([k, n, p, cum]) => {
    k = Math.trunc(k)
    n = Math.trunc(n)
    if (k < 0 || k > n || p < 0 || p > 1) return ERR.NUM
    if (p === 0 || p === 1) return cum ? (p === 0 ? 1 : k === n ? 1 : 0) : (p === 0 ? (k === 0 ? 1 : 0) : k === n ? 1 : 0)
    return cumOrPdf(cum, binomCdf(k, n, p), binomPmf(k, n, p))
  }),
  fn('BINOM.DIST.RANGE', 3, 4, 'n', ([n, p, s, s2]) => {
    n = Math.trunc(n)
    s = Math.trunc(s)
    const e = s2 === undefined ? s : Math.trunc(s2)
    if (n < 0 || p < 0 || p > 1 || s < 0 || s > n || e < s || e > n) return ERR.NUM
    let t = 0
    for (let k = s; k <= e; k++) t += combin(n, k) * p ** k * (1 - p) ** (n - k)
    return t
  }),
  ...(['BINOM.INV', 'CRITBINOM'] as const).map((name) => fn(name, 3, 3, 'n', ([n, p, a]) => binomInv(n, p, a))),
  fn('NEGBINOM.DIST', 4, 4, 'nnnb', ([f, s, p, cum]) => {
    f = Math.trunc(f)
    s = Math.trunc(s)
    if (f < 0 || s < 1 || p < 0 || p > 1) return ERR.NUM
    const pmf = (k: number) => combin(k + s - 1, s - 1) * p ** s * (1 - p) ** k
    if (!cum) return pmf(f)
    return 1 - betaI(1 - p, f + 1, s)
  }),
  fn('NEGBINOMDIST', 3, 3, 'n', ([f, s, p]) => {
    f = Math.trunc(f)
    s = Math.trunc(s)
    if (f < 0 || s < 1 || p < 0 || p > 1) return ERR.NUM
    return combin(f + s - 1, s - 1) * p ** s * (1 - p) ** f
  }),
  ...(['POISSON.DIST', 'POISSON'] as const).map((n) =>
    fn(n, 3, 3, 'nnb', ([x, m, cum]) => {
      x = Math.trunc(x)
      if (x < 0 || m < 0) return ERR.NUM
      if (cum) return m === 0 ? 1 : 1 - gammaCdf(m, x + 1, 1)
      return Math.exp(x * Math.log(m) - m - lnGamma(x + 1))
    })
  ),
  fn('HYPGEOM.DIST', 5, 5, 'nnnnb', ([s, n, K, N, cum]) => {
    s = Math.trunc(s)
    n = Math.trunc(n)
    K = Math.trunc(K)
    N = Math.trunc(N)
    if (s < 0 || s > n || s > K || n > N || K > N || n - s > N - K || n <= 0 || K <= 0 || N <= 0) return ERR.NUM
    if (!cum) return hypgeom(s, n, K, N)
    let t = 0
    for (let k = Math.max(0, n - N + K); k <= s; k++) t += hypgeom(k, n, K, N)
    return t
  }),
  fn('HYPGEOMDIST', 4, 4, 'n', ([s, n, K, N]) => {
    s = Math.trunc(s)
    n = Math.trunc(n)
    K = Math.trunc(K)
    N = Math.trunc(N)
    if (s < 0 || s > n || s > K || n > N || K > N || n - s > N - K) return ERR.NUM
    return hypgeom(s, n, K, N)
  }),
  fn('CHISQ.DIST', 3, 3, 'nnb', ([x, df, cum]) => {
    df = Math.trunc(df)
    if (x < 0 || df < 1 || df > 1e10) return ERR.NUM
    return cumOrPdf(cum, chiCdf(x, df), chiPdf(x, df))
  }),
  ...(['CHISQ.DIST.RT', 'CHIDIST'] as const).map((n) => fn(n, 2, 2, 'n', ([x, df]) => (x < 0 || Math.trunc(df) < 1 ? ERR.NUM : 1 - chiCdf(x, Math.trunc(df))))),
  fn('CHISQ.INV', 2, 2, 'n', ([p, df]) => (p < 0 || p > 1 || Math.trunc(df) < 1 ? ERR.NUM : p === 0 ? 0 : invert((x) => chiCdf(x, Math.trunc(df)), p, 0, df * 3 + 10))),
  ...(['CHISQ.INV.RT', 'CHIINV'] as const).map((n) => fn(n, 2, 2, 'n', ([p, df]) => (p <= 0 || p > 1 || Math.trunc(df) < 1 ? ERR.NUM : invert((x) => chiCdf(x, Math.trunc(df)), 1 - p, 0, df * 3 + 10)))),
  ...(['CHISQ.TEST', 'CHITEST'] as const).map((n) => fn(n, 2, 2, 'aa', ([a, e]) => chiTest(a, e))),
  fn('F.DIST', 4, 4, 'nnnb', ([x, d1, d2, cum]) => {
    d1 = Math.trunc(d1)
    d2 = Math.trunc(d2)
    if (x < 0 || d1 < 1 || d2 < 1) return ERR.NUM
    return cumOrPdf(cum, fCdf(x, d1, d2), fPdf(x, d1, d2))
  }),
  ...(['F.DIST.RT', 'FDIST'] as const).map((n) => fn(n, 3, 3, 'n', ([x, d1, d2]) => (x < 0 || Math.trunc(d1) < 1 || Math.trunc(d2) < 1 ? ERR.NUM : 1 - fCdf(x, Math.trunc(d1), Math.trunc(d2))))),
  fn('F.INV', 3, 3, 'n', ([p, d1, d2]) => (p < 0 || p > 1 || Math.trunc(d1) < 1 || Math.trunc(d2) < 1 ? ERR.NUM : p === 0 ? 0 : invert((x) => fCdf(x, Math.trunc(d1), Math.trunc(d2)), p, 0, 100))),
  ...(['F.INV.RT', 'FINV'] as const).map((n) => fn(n, 3, 3, 'n', ([p, d1, d2]) => (p <= 0 || p > 1 || Math.trunc(d1) < 1 || Math.trunc(d2) < 1 ? ERR.NUM : invert((x) => fCdf(x, Math.trunc(d1), Math.trunc(d2)), 1 - p, 0, 100)))),
  ...(['F.TEST', 'FTEST'] as const).map((n) =>
    fn(n, 2, 2, 'r', ([a, b], ctx) => {
      const x = nums([a], ctx)
      const y = nums([b], ctx)
      if (isError(x)) return x
      if (isError(y)) return y
      return fTest(x, y)
    })
  ),
  fn('T.DIST', 3, 3, 'nnb', ([t, df, cum]) => (Math.trunc(df) < 1 ? ERR.DIV0 : cumOrPdf(cum, tCdf(t, Math.trunc(df)), tPdf(t, Math.trunc(df))))),
  fn('T.DIST.2T', 2, 2, 'n', ([t, df]) => (t < 0 || Math.trunc(df) < 1 ? ERR.NUM : 2 * (1 - tCdf(t, Math.trunc(df))))),
  fn('T.DIST.RT', 2, 2, 'n', ([t, df]) => (Math.trunc(df) < 1 ? ERR.NUM : 1 - tCdf(t, Math.trunc(df)))),
  fn('TDIST', 3, 3, 'n', ([t, df, tails]) => {
    tails = Math.trunc(tails)
    if (t < 0 || Math.trunc(df) < 1 || (tails !== 1 && tails !== 2)) return ERR.NUM
    const p = 1 - tCdf(t, Math.trunc(df))
    return tails === 2 ? 2 * p : p
  }),
  fn('T.INV', 2, 2, 'n', ([p, df]) => (p <= 0 || p >= 1 || Math.trunc(df) < 1 ? ERR.NUM : invert((x) => tCdf(x, Math.trunc(df)), p, -100, 100))),
  ...(['T.INV.2T', 'TINV'] as const).map((n) => fn(n, 2, 2, 'n', ([p, df]) => (p <= 0 || p > 1 || Math.trunc(df) < 1 ? ERR.NUM : invert((x) => tCdf(x, Math.trunc(df)), 1 - p / 2, -100, 100)))),
  ...(['T.TEST', 'TTEST'] as const).map((n) =>
    fn(n, 4, 4, 'rrnn', ([a, b, tails, type], ctx) => {
      const ta = Math.trunc(type)
      if (ta === 1) {
        const p = pairs(a, b, ctx)
        if (isError(p)) return p
        return tTest(p.x, p.y, Math.trunc(tails), 1)
      }
      const x = nums([a], ctx)
      const y = nums([b], ctx)
      if (isError(x)) return x
      if (isError(y)) return y
      return tTest(x, y, Math.trunc(tails), ta)
    })
  ),
  ...(['CONFIDENCE.NORM', 'CONFIDENCE'] as const).map((n) => fn(n, 3, 3, 'n', ([alpha, s, size]) => (alpha <= 0 || alpha >= 1 || s <= 0 || Math.trunc(size) < 1 ? ERR.NUM : (normInv(1 - alpha / 2) * s) / Math.sqrt(Math.trunc(size))))),
  fn('CONFIDENCE.T', 3, 3, 'n', ([alpha, s, size]) => {
    size = Math.trunc(size)
    if (alpha <= 0 || alpha >= 1 || s <= 0 || size < 1) return ERR.NUM
    if (size === 1) return ERR.DIV0
    return (invert((x) => tCdf(x, size - 1), 1 - alpha / 2, -100, 100) * s) / Math.sqrt(size)
  }),
  ...(['Z.TEST', 'ZTEST'] as const).map((n) =>
    fn(n, 2, 3, 'rnn', ([a, x, sigma], ctx) => {
      const xs = nums([a], ctx)
      return isError(xs) ? xs : zTest(xs, x, sigma)
    })
  )
]

defineFunctions(statFns)

void erf
