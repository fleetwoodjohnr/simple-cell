// Financial functions. Formulas follow Microsoft's documentation; the bond and depreciation algorithms follow
// LibreOffice's analysis add-in, which was written to match Excel.
import { defineFunctions, toArray, type EvalContext } from '../evaluate'
import { CellError, ERR, isError, type Scalar, type Value } from '../values'
import { checkNum, fn, MANY, numbers } from './helpers'
import { days360, yearFrac } from './date'
import { dateSerial, daysInMonth, serialToDate, type YMD } from '../../numfmt/dates'

// ---------- time value of money ----------

export function pv(r: number, n: number, pmt: number, fv = 0, type = 0): number {
  if (r === 0) return -(fv + pmt * n)
  const f = Math.pow(1 + r, n)
  return -(fv + (pmt * (1 + r * type) * (f - 1)) / r) / f
}
export function fv(r: number, n: number, pmt: number, pv0 = 0, type = 0): number {
  if (r === 0) return -(pv0 + pmt * n)
  const f = Math.pow(1 + r, n)
  return -(pv0 * f + (pmt * (1 + r * type) * (f - 1)) / r)
}
export function pmt(r: number, n: number, pv0: number, fv0 = 0, type = 0): number {
  if (r === 0) return -(pv0 + fv0) / n
  const f = Math.pow(1 + r, n)
  return -(r * (fv0 + pv0 * f)) / ((1 + r * type) * (f - 1))
}
function ipmt(r: number, per: number, n: number, pv0: number, fv0 = 0, type = 0): number {
  const p = pmt(r, n, pv0, fv0, type)
  let interest: number
  if (per === 1) interest = type === 1 ? 0 : -pv0
  else interest = type === 1 ? fv(r, per - 2, p, pv0, 1) - p : fv(r, per - 1, p, pv0, 0)
  return interest * r
}

/** Newton's method with a bisection fallback; null when it doesn't converge. */
function solve(f: (x: number) => number, guess: number, lo = -0.999999, hi = 1e6): number | null {
  let x = guess
  for (let i = 0; i < 200; i++) {
    const y = f(x)
    if (!Number.isFinite(y)) break
    if (Math.abs(y) < 1e-10) return x
    const h = Math.max(1e-7, Math.abs(x) * 1e-7)
    const d = (f(x + h) - y) / h
    if (!d || !Number.isFinite(d)) break
    const nx = x - y / d
    if (Math.abs(nx - x) < 1e-12) return nx
    x = nx
  }
  // Bisection on a sign change.
  const fl = f(lo)
  let a = lo
  let b = hi
  let fa = fl
  const fb = f(b)
  if (!Number.isFinite(fa) || !Number.isFinite(fb) || fa * fb > 0) {
    // scan for a bracket
    let prev = lo
    let pf = fl
    for (let t = lo; t < 100; t = t < 1 ? t + 0.01 : t * 1.5) {
      const ft = f(t)
      if (Number.isFinite(ft) && Number.isFinite(pf) && pf * ft <= 0) {
        a = prev
        b = t
        fa = pf
        break
      }
      prev = t
      pf = ft
    }
    if (fa * f(b) > 0) return null
  }
  for (let i = 0; i < 300; i++) {
    const m = (a + b) / 2
    const fm = f(m)
    if (Math.abs(fm) < 1e-10 || b - a < 1e-14) return m
    if (fa * fm < 0) b = m
    else {
      a = m
      fa = fm
    }
  }
  return (a + b) / 2
}

function npv(r: number, vals: number[]): number {
  let s = 0
  vals.forEach((v, i) => (s += v / Math.pow(1 + r, i + 1)))
  return s
}

function cashflows(v: Value, ctx: EvalContext): number[] | CellError {
  return numbers([v], ctx)
}

function datesAndValues(vals: Value, dates: Value, ctx: EvalContext): { v: number[]; d: number[] } | CellError {
  const a = toArray(vals, ctx)
  const b = toArray(dates, ctx)
  if (isError(a)) return a
  if (isError(b)) return b
  if (a.data.length !== b.data.length) return ERR.NUM
  const v: number[] = []
  const d: number[] = []
  for (let i = 0; i < a.data.length; i++) {
    const x = a.data[i]
    const y = b.data[i]
    if (x instanceof CellError) return x
    if (y instanceof CellError) return y
    if (typeof x !== 'number' || typeof y !== 'number') return ERR.VALUE
    v.push(x)
    d.push(Math.floor(y))
  }
  if (d.some((x) => x < d[0])) return ERR.NUM
  return { v, d }
}

const xnpv = (r: number, v: number[], d: number[]) => v.reduce((s, x, i) => s + x / Math.pow(1 + r, (d[i] - d[0]) / 365), 0)

// ---------- depreciation ----------

function db(cost: number, salvage: number, life: number, period: number, month: number): number | CellError {
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12 || period > life + (month < 12 ? 1 : 0)) return ERR.NUM
  if (cost === 0) return 0
  const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000
  let total = 0
  let dep = 0
  for (let p = 1; p <= period; p++) {
    if (p === 1) dep = (cost * rate * month) / 12
    else if (p === Math.floor(life) + 1) dep = ((cost - total) * rate * (12 - month)) / 12
    else dep = (cost - total) * rate
    total += dep
  }
  return dep
}

/** Double-declining depreciation of one period (LibreOffice ScGetDDB; equals Excel's DDB). */
function getDDB(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life
  let oldValue: number
  if (rate >= 1) {
    rate = 1
    oldValue = period === 1 ? cost : 0
  } else oldValue = cost * Math.pow(1 - rate, period - 1)
  const newValue = cost * Math.pow(1 - rate, period)
  const d = newValue < salvage ? oldValue - salvage : oldValue - newValue
  return d < 0 ? 0 : d
}

function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number | CellError {
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0 || period > life) return ERR.NUM
  return getDDB(cost, salvage, life, period, factor)
}

/** Declining balance up to `period`, switching to straight line when that is larger (LibreOffice ScInterVDB). */
function interVDB(cost: number, salvage: number, life: number, life1: number, period: number, factor: number): number {
  let v = 0
  const intEnd = Math.ceil(period - 1e-12)
  let sln = 0
  let rest = cost - salvage
  let nowSln = false
  for (let i = 1; i <= intEnd; i++) {
    let term: number
    if (!nowSln) {
      const d = getDDB(cost, salvage, life, i, factor)
      sln = rest / (life1 - (i - 1))
      if (sln > d) {
        term = sln
        nowSln = true
      } else {
        term = d
        rest -= d
      }
    } else term = sln
    if (i === intEnd) term *= period + 1 - intEnd
    v += term
  }
  return v
}

/** VDB (LibreOffice ScVDB, which matches Excel). */
function vdb(cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean): number | CellError {
  if (start < 0 || end < start || end > life || cost < 0 || salvage > cost || factor <= 0) return ERR.NUM
  const intStart = Math.floor(start + 1e-12)
  const intEnd = Math.ceil(end - 1e-12)
  const eq = (a: number, b: number) => Math.abs(a - b) < 1e-12
  if (noSwitch) {
    let v = 0
    for (let i = intStart + 1; i <= intEnd; i++) {
      let term = getDDB(cost, salvage, life, i, factor)
      if (i === intStart + 1) term *= Math.min(end, intStart + 1) - start
      else if (i === intEnd) term *= end + 1 - intEnd
      v += term
    }
    return v
  }
  let part = 0
  if (!eq(start, intStart)) {
    const temp = cost - interVDB(cost, salvage, life, life, intStart, factor)
    part += (start - intStart) * interVDB(temp, salvage, life, life - intStart, 1, factor)
  }
  if (!eq(end, intEnd)) {
    const tStart = intEnd - 1
    const temp = cost - interVDB(cost, salvage, life, life, tStart, factor)
    part += (intEnd - end) * interVDB(temp, salvage, life, life - tStart, intEnd - tStart, factor)
  }
  const c = cost - interVDB(cost, salvage, life, life, intStart, factor)
  return interVDB(c, salvage, life, life - intStart, intEnd - intStart, factor) - part
}

// ---------- day counts and coupons ----------

function dayCount(a: number, b: number, basis: number, d1904: boolean): number {
  if (basis === 0) return days360(serialToDate(a, d1904), serialToDate(b, d1904), false)
  if (basis === 4) return days360(serialToDate(a, d1904), serialToDate(b, d1904), true)
  return b - a
}

const yearBasis = (basis: number, y: number) => (basis === 1 ? (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0) ? 366 : 365) : basis === 3 ? 365 : 360)

function addMonths(d: YMD, months: number, eom: boolean): YMD {
  const t = d.m - 1 + months
  const y = d.y + Math.floor(t / 12)
  const m = (((t % 12) + 12) % 12) + 1
  const dim = daysInMonth(y, m)
  return { y, m, d: eom ? dim : Math.min(d.d, dim) }
}

interface Coupon {
  pcd: number
  ncd: number
  num: number
}

function coupons(settle: number, mat: number, freq: number, d1904: boolean): Coupon | CellError {
  if (![1, 2, 4].includes(freq) || settle >= mat) return ERR.NUM
  const m = serialToDate(mat, d1904)
  const eom = m.d === daysInMonth(m.y, m.m)
  const step = 12 / freq
  let k = 0
  let date = mat
  // Step back from maturity until on or before settlement.
  while (date > settle) {
    k++
    const d = addMonths(m, -k * step, eom)
    date = dateSerial(d.y, d.m, d.d, d1904)
    if (k > 4000) return ERR.NUM
  }
  const nd = addMonths(m, -(k - 1) * step, eom)
  return { pcd: date, ncd: dateSerial(nd.y, nd.m, nd.d, d1904), num: k }
}

function coupDays(c: Coupon, basis: number, freq: number): number {
  if (basis === 1) return c.ncd - c.pcd
  return (basis === 3 ? 365 : 360) / freq
}

function couponArgs(args: number[], n: number): { s: number; m: number; f: number; b: number } | CellError {
  const s = Math.floor(args[0])
  const m = Math.floor(args[1])
  const f = Math.trunc(args[n])
  const b = Math.trunc(args[n + 1] ?? 0)
  if (b < 0 || b > 4 || ![1, 2, 4].includes(f) || s >= m) return ERR.NUM
  return { s, m, f, b }
}

function priceOf(s: number, mt: number, rate: number, yld: number, redemp: number, f: number, b: number, d1904: boolean): number | CellError {
  const c = coupons(s, mt, f, d1904)
  if (isError(c)) return c
  const E = coupDays(c, b, f)
  const A = dayCount(c.pcd, s, b, d1904)
  const DSC = b === 0 || b === 4 ? E - A : c.ncd - s
  const N = c.num
  const cpn = (100 * rate) / f
  if (N === 1) return (redemp + cpn) / (1 + ((DSC / E) * yld) / f) - (cpn * A) / E
  let p = redemp / Math.pow(1 + yld / f, N - 1 + DSC / E)
  for (let k = 1; k <= N; k++) p += cpn / Math.pow(1 + yld / f, k - 1 + DSC / E)
  return p - (cpn * A) / E
}

/** Macaulay duration in years, discounting each coupon over coupon periods (as PRICE does). */
function duration(s: number, mt: number, cpnRate: number, yld: number, f: number, b: number, d1904: boolean): number | CellError {
  const c = coupons(s, mt, f, d1904)
  if (isError(c)) return c
  const E = coupDays(c, b, f)
  const A = dayCount(c.pcd, s, b, d1904)
  const DSC = b === 0 || b === 4 ? E - A : c.ncd - s
  const off = DSC / E
  const cpn = (cpnRate * 100) / f
  const y = 1 + yld / f
  let dur = 0
  let p = 0
  for (let k = 1; k <= c.num; k++) {
    const t = k - 1 + off
    const cf = k === c.num ? cpn + 100 : cpn
    const pv = cf / Math.pow(y, t)
    dur += t * pv
    p += pv
  }
  return dur / p / f
}

const yf = (a: number, b: number, basis: number, ctx: EvalContext) => yearFrac(a, b, basis, ctx.wb.date1904)

function checkBasis(b: number | undefined): number | CellError {
  const x = Math.trunc(b ?? 0)
  return x < 0 || x > 4 ? ERR.NUM : x
}

// ---------- definitions ----------

defineFunctions([
  fn('PV', 3, 5, 'n', ([r, n, p, f = 0, t = 0]) => checkNum(pv(r, n, p, f, t ? 1 : 0))),
  fn('FV', 3, 5, 'n', ([r, n, p, v = 0, t = 0]) => checkNum(fv(r, n, p, v, t ? 1 : 0))),
  fn('PMT', 3, 5, 'n', ([r, n, v, f = 0, t = 0]) => (n === 0 ? ERR.NUM : checkNum(pmt(r, n, v, f, t ? 1 : 0)))),
  fn('NPER', 3, 5, 'n', ([r, p, v, f = 0, t = 0]) => {
    const ty = t ? 1 : 0
    if (r === 0) return p === 0 ? ERR.NUM : -(v + f) / p
    const num = p * (1 + r * ty) - f * r
    const den = v * r + p * (1 + r * ty)
    if (num / den <= 0) return ERR.NUM
    return checkNum(Math.log(num / den) / Math.log(1 + r))
  }),
  fn('IPMT', 4, 6, 'n', ([r, per, n, v, f = 0, t = 0]) => (per < 1 || per > n ? ERR.NUM : checkNum(ipmt(r, per, n, v, f, t ? 1 : 0)))),
  fn('PPMT', 4, 6, 'n', ([r, per, n, v, f = 0, t = 0]) => (per < 1 || per > n ? ERR.NUM : checkNum(pmt(r, n, v, f, t ? 1 : 0) - ipmt(r, per, n, v, f, t ? 1 : 0)))),
  fn('CUMIPMT', 6, 6, 'n', ([r, n, v, s, e, t]) => {
    s = Math.trunc(s)
    e = Math.trunc(e)
    if (r <= 0 || n <= 0 || v <= 0 || s < 1 || e < 1 || s > e || (t !== 0 && t !== 1)) return ERR.NUM
    let total = 0
    for (let p = s; p <= e; p++) total += ipmt(r, p, n, v, 0, t)
    return total
  }),
  fn('CUMPRINC', 6, 6, 'n', ([r, n, v, s, e, t]) => {
    s = Math.trunc(s)
    e = Math.trunc(e)
    if (r <= 0 || n <= 0 || v <= 0 || s < 1 || e < 1 || s > e || (t !== 0 && t !== 1)) return ERR.NUM
    const p = pmt(r, n, v, 0, t)
    let total = 0
    for (let k = s; k <= e; k++) total += p - ipmt(r, k, n, v, 0, t)
    return total
  }),
  fn('RATE', 3, 6, 'n', ([n, p, v, f = 0, t = 0, guess = 0.1]) => {
    const ty = t ? 1 : 0
    const g = (r: number) => (Math.abs(r) < 1e-12 ? v + p * n + f : v * Math.pow(1 + r, n) + (p * (1 + r * ty) * (Math.pow(1 + r, n) - 1)) / r + f)
    const r = solve(g, guess)
    return r === null ? ERR.NUM : r
  }),
  fn('NPV', 2, MANY, 'n|r', ([r, ...vals], ctx) => {
    const xs = numbers(vals, ctx)
    if (isError(xs)) return xs
    if (r === -1) return ERR.DIV0
    return checkNum(npv(r, xs))
  }),
  fn('XNPV', 3, 3, 'nrr', ([r, vals, dates], ctx) => {
    const d = datesAndValues(vals, dates, ctx)
    if (isError(d)) return d
    if (r <= -1) return ERR.NUM
    return checkNum(xnpv(r, d.v, d.d))
  }),
  fn('IRR', 1, 2, 'rn', ([vals, guess = 0.1], ctx) => {
    const xs = cashflows(vals, ctx)
    if (isError(xs)) return xs
    if (!xs.some((x) => x > 0) || !xs.some((x) => x < 0)) return ERR.NUM
    const r = solve((x) => xs.reduce((s, v, i) => s + v / Math.pow(1 + x, i), 0), guess)
    return r === null ? ERR.NUM : r
  }),
  fn('XIRR', 2, 3, 'rrn', ([vals, dates, guess = 0.1], ctx) => {
    const d = datesAndValues(vals, dates, ctx)
    if (isError(d)) return d
    if (!d.v.some((x) => x > 0) || !d.v.some((x) => x < 0)) return ERR.NUM
    const r = solve((x) => xnpv(x, d.v, d.d), guess)
    return r === null ? ERR.NUM : r
  }),
  fn('MIRR', 3, 3, 'rnn', ([vals, fr, rr], ctx) => {
    const xs = cashflows(vals, ctx)
    if (isError(xs)) return xs
    const n = xs.length
    let fvPos = 0
    let pvNeg = 0
    xs.forEach((v, i) => {
      if (v > 0) fvPos += v * Math.pow(1 + rr, n - 1 - i)
      else pvNeg += v / Math.pow(1 + fr, i)
    })
    if (pvNeg === 0 || fvPos === 0) return ERR.DIV0
    return Math.pow(-fvPos / pvNeg, 1 / (n - 1)) - 1
  }),
  fn('EFFECT', 2, 2, 'n', ([nom, n]) => {
    n = Math.trunc(n)
    return nom <= 0 || n < 1 ? ERR.NUM : Math.pow(1 + nom / n, n) - 1
  }),
  fn('NOMINAL', 2, 2, 'n', ([eff, n]) => {
    n = Math.trunc(n)
    return eff <= 0 || n < 1 ? ERR.NUM : n * (Math.pow(1 + eff, 1 / n) - 1)
  }),
  fn('FVSCHEDULE', 2, 2, 'nr', ([p, sched], ctx) => {
    const a = toArray(sched, ctx)
    if (isError(a)) return a
    let v = p
    for (const r of a.data) {
      if (r instanceof CellError) return r
      if (r === null) continue
      if (typeof r !== 'number') return ERR.VALUE
      v *= 1 + r
    }
    return v
  }),
  fn('PDURATION', 3, 3, 'n', ([r, p, f]) => (r <= 0 || p <= 0 || f <= 0 ? ERR.NUM : (Math.log(f) - Math.log(p)) / Math.log(1 + r))),
  fn('RRI', 3, 3, 'n', ([n, p, f]) => (n <= 0 || p === 0 ? ERR.NUM : checkNum(Math.pow(f / p, 1 / n) - 1))),
  fn('ISPMT', 4, 4, 'n', ([r, per, n, v]) => (n === 0 ? ERR.DIV0 : v * r * (per / n - 1))),
  fn('SLN', 3, 3, 'n', ([c, s, l]) => (l === 0 ? ERR.DIV0 : (c - s) / l)),
  fn('SYD', 4, 4, 'n', ([c, s, l, per]) => (l <= 0 || per <= 0 || per > l ? ERR.NUM : ((c - s) * (l - per + 1) * 2) / (l * (l + 1)))),
  fn('DB', 4, 5, 'n', ([c, s, l, per, m = 12]) => db(c, s, l, Math.trunc(per), Math.trunc(m))),
  fn('DDB', 4, 5, 'n', ([c, s, l, per, f = 2]) => ddb(c, s, l, per, f)),
  fn('VDB', 5, 7, 'nnnnnnb', ([c, s, l, a, b, f = 2, noSwitch = false]) => vdb(c, s, l, a, b, f, noSwitch)),
  fn('DOLLARDE', 2, 2, 'n', ([d, f]) => {
    f = Math.trunc(f)
    if (f < 0) return ERR.NUM
    if (f === 0) return ERR.DIV0
    const int = Math.trunc(d)
    const frac = d - int
    return int + (frac * Math.pow(10, Math.ceil(Math.log10(f)))) / f
  }),
  fn('DOLLARFR', 2, 2, 'n', ([d, f]) => {
    f = Math.trunc(f)
    if (f < 0) return ERR.NUM
    if (f === 0) return ERR.DIV0
    const int = Math.trunc(d)
    const frac = d - int
    return int + (frac * f) / Math.pow(10, Math.ceil(Math.log10(f)))
  }),
  // securities
  fn('COUPPCD', 3, 4, 'n', (args, ctx) => {
    const a = couponArgs(args, 2)
    if (isError(a)) return a
    const c = coupons(a.s, a.m, a.f, ctx.wb.date1904)
    return isError(c) ? c : c.pcd
  }),
  fn('COUPNCD', 3, 4, 'n', (args, ctx) => {
    const a = couponArgs(args, 2)
    if (isError(a)) return a
    const c = coupons(a.s, a.m, a.f, ctx.wb.date1904)
    return isError(c) ? c : c.ncd
  }),
  fn('COUPNUM', 3, 4, 'n', (args, ctx) => {
    const a = couponArgs(args, 2)
    if (isError(a)) return a
    const c = coupons(a.s, a.m, a.f, ctx.wb.date1904)
    return isError(c) ? c : c.num
  }),
  fn('COUPDAYS', 3, 4, 'n', (args, ctx) => {
    const a = couponArgs(args, 2)
    if (isError(a)) return a
    const c = coupons(a.s, a.m, a.f, ctx.wb.date1904)
    return isError(c) ? c : coupDays(c, a.b, a.f)
  }),
  fn('COUPDAYBS', 3, 4, 'n', (args, ctx) => {
    const a = couponArgs(args, 2)
    if (isError(a)) return a
    const c = coupons(a.s, a.m, a.f, ctx.wb.date1904)
    return isError(c) ? c : dayCount(c.pcd, a.s, a.b, ctx.wb.date1904)
  }),
  fn('COUPDAYSNC', 3, 4, 'n', (args, ctx) => {
    const a = couponArgs(args, 2)
    if (isError(a)) return a
    const c = coupons(a.s, a.m, a.f, ctx.wb.date1904)
    if (isError(c)) return c
    if (a.b === 0 || a.b === 4) return coupDays(c, a.b, a.f) - dayCount(c.pcd, a.s, a.b, ctx.wb.date1904)
    return c.ncd - a.s
  }),
  fn('PRICE', 6, 7, 'n', ([s, m, rate, yld, red, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (rate < 0 || yld < 0 || red <= 0 || ![1, 2, 4].includes(Math.trunc(f)) || Math.floor(s) >= Math.floor(m)) return ERR.NUM
    return priceOf(Math.floor(s), Math.floor(m), rate, yld, red, Math.trunc(f), bb, ctx.wb.date1904)
  }),
  fn('YIELD', 6, 7, 'n', ([s, m, rate, pr, red, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    f = Math.trunc(f)
    s = Math.floor(s)
    m = Math.floor(m)
    if (rate < 0 || pr <= 0 || red <= 0 || ![1, 2, 4].includes(f) || s >= m) return ERR.NUM
    const c = coupons(s, m, f, ctx.wb.date1904)
    if (isError(c)) return c
    if (c.num === 1) {
      const E = coupDays(c, bb, f)
      const A = dayCount(c.pcd, s, bb, ctx.wb.date1904)
      const DSR = E - A
      const cpn = (100 * rate) / f
      return ((red / 100 + cpn / 100 - (pr / 100 + ((A / E) * cpn) / 100)) / (pr / 100 + ((A / E) * cpn) / 100)) * ((f * E) / DSR)
    }
    const y = solve((x) => (priceOf(s, m, rate, x, red, f, bb, ctx.wb.date1904) as number) - pr, rate || 0.05, 0, 10)
    return y === null ? ERR.NUM : y
  }),
  fn('DURATION', 5, 6, 'n', ([s, m, c, y, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (c < 0 || y < 0 || ![1, 2, 4].includes(Math.trunc(f)) || Math.floor(s) >= Math.floor(m)) return ERR.NUM
    return duration(Math.floor(s), Math.floor(m), c, y, Math.trunc(f), bb, ctx.wb.date1904)
  }),
  fn('MDURATION', 5, 6, 'n', ([s, m, c, y, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (c < 0 || y < 0 || ![1, 2, 4].includes(Math.trunc(f)) || Math.floor(s) >= Math.floor(m)) return ERR.NUM
    const d = duration(Math.floor(s), Math.floor(m), c, y, Math.trunc(f), bb, ctx.wb.date1904)
    return isError(d) ? d : d / (1 + y / Math.trunc(f))
  }),
  fn('ACCRINT', 6, 8, 'nnnnnnnb', ([issue, _first, s, rate, par, f, b = 0], ctx) => {
    void _first
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (rate <= 0 || par <= 0 || ![1, 2, 4].includes(Math.trunc(f)) || issue >= s) return ERR.NUM
    const y = yf(issue, s, bb, ctx)
    return isError(y) ? y : par * rate * y
  }),
  fn('ACCRINTM', 4, 5, 'n', ([issue, s, rate, par = 1000, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (rate <= 0 || par <= 0 || issue >= s) return ERR.NUM
    const y = yf(issue, s, bb, ctx)
    return isError(y) ? y : par * rate * y
  }),
  fn('DISC', 4, 5, 'n', ([s, m, pr, red, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (pr <= 0 || red <= 0 || s >= m) return ERR.NUM
    const y = yf(s, m, bb, ctx)
    return isError(y) ? y : (red - pr) / red / y
  }),
  fn('INTRATE', 4, 5, 'n', ([s, m, inv, red, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (inv <= 0 || red <= 0 || s >= m) return ERR.NUM
    const y = yf(s, m, bb, ctx)
    return isError(y) ? y : (red - inv) / inv / y
  }),
  fn('RECEIVED', 4, 5, 'n', ([s, m, inv, disc, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (inv <= 0 || disc <= 0 || s >= m) return ERR.NUM
    const y = yf(s, m, bb, ctx)
    if (isError(y)) return y
    const d = 1 - disc * y
    return d <= 0 ? ERR.NUM : inv / d
  }),
  fn('PRICEDISC', 4, 5, 'n', ([s, m, disc, red, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (disc <= 0 || red <= 0 || s >= m) return ERR.NUM
    const y = yf(s, m, bb, ctx)
    return isError(y) ? y : red * (1 - disc * y)
  }),
  fn('YIELDDISC', 4, 5, 'n', ([s, m, pr, red, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (pr <= 0 || red <= 0 || s >= m) return ERR.NUM
    const y = yf(s, m, bb, ctx)
    return isError(y) ? y : (red / pr - 1) / y
  }),
  fn('PRICEMAT', 5, 6, 'n', ([s, m, issue, rate, yld, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (rate < 0 || yld < 0 || s >= m) return ERR.NUM
    const im = yf(issue, m, bb, ctx) as number
    const is = yf(issue, s, bb, ctx) as number
    const sm = yf(s, m, bb, ctx) as number
    return ((1 + im * rate) / (1 + sm * yld) - is * rate) * 100
  }),
  fn('YIELDMAT', 5, 6, 'n', ([s, m, issue, rate, pr, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (rate < 0 || pr <= 0 || s >= m) return ERR.NUM
    const im = yf(issue, m, bb, ctx) as number
    const is = yf(issue, s, bb, ctx) as number
    const sm = yf(s, m, bb, ctx) as number
    return ((1 + im * rate) / (pr / 100 + is * rate) - 1) / sm
  }),
  fn('TBILLPRICE', 3, 3, 'n', ([s, m, disc]) => {
    s = Math.floor(s)
    m = Math.floor(m)
    if (disc <= 0 || s >= m || m - s > 365) return ERR.NUM
    const p = 100 * (1 - (disc * (m - s)) / 360)
    return p <= 0 ? ERR.NUM : p
  }),
  fn('TBILLYIELD', 3, 3, 'n', ([s, m, pr]) => {
    s = Math.floor(s)
    m = Math.floor(m)
    if (pr <= 0 || s >= m || m - s > 365) return ERR.NUM
    return ((100 - pr) / pr) * (360 / (m - s))
  }),
  fn('TBILLEQ', 3, 3, 'n', ([s, m, disc]) => {
    s = Math.floor(s)
    m = Math.floor(m)
    const dsm = m - s
    if (disc <= 0 || s >= m || dsm > 365) return ERR.NUM
    if (dsm <= 182) return (365 * disc) / (360 - disc * dsm)
    const price = 100 * (1 - (disc * dsm) / 360)
    const t = dsm / 365
    return (-t + Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price))) / (t - 0.5)
  }),
  fn('ODDLPRICE', 7, 8, 'n', ([s, m, last, rate, yld, red, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    f = Math.trunc(f)
    if (rate < 0 || yld < 0 || red <= 0 || ![1, 2, 4].includes(f) || !(last < s && s < m)) return ERR.NUM
    const dci = (yf(last, m, bb, ctx) as number) * f
    const dsci = (yf(s, m, bb, ctx) as number) * f
    const ai = (yf(last, s, bb, ctx) as number) * f
    return (red + (dci * 100 * rate) / f) / ((dsci * yld) / f + 1) - (ai * 100 * rate) / f
  }),
  fn('ODDLYIELD', 7, 8, 'n', ([s, m, last, rate, pr, red, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    f = Math.trunc(f)
    if (rate < 0 || pr <= 0 || red <= 0 || ![1, 2, 4].includes(f) || !(last < s && s < m)) return ERR.NUM
    const dci = (yf(last, m, bb, ctx) as number) * f
    const dsci = (yf(s, m, bb, ctx) as number) * f
    const ai = (yf(last, s, bb, ctx) as number) * f
    return ((red + (dci * 100 * rate) / f) / (pr + (ai * 100 * rate) / f) - 1) * (f / dsci)
  }),
  fn('ODDFPRICE', 8, 9, 'n', ([s, m, issue, first, rate, yld, red, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    f = Math.trunc(f)
    if (rate < 0 || yld < 0 || red <= 0 || ![1, 2, 4].includes(f) || !(issue < s && s < first && first < m)) return ERR.NUM
    return oddFirstPrice(Math.floor(s), Math.floor(m), Math.floor(issue), Math.floor(first), rate, yld, red, f, bb, ctx.wb.date1904)
  }),
  fn('ODDFYIELD', 8, 9, 'n', ([s, m, issue, first, rate, pr, red, f, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    f = Math.trunc(f)
    if (rate < 0 || pr <= 0 || red <= 0 || ![1, 2, 4].includes(f) || !(issue < s && s < first && first < m)) return ERR.NUM
    const y = solve((x) => (oddFirstPrice(Math.floor(s), Math.floor(m), Math.floor(issue), Math.floor(first), rate, x, red, f, bb, ctx.wb.date1904) as number) - pr, rate || 0.05, 0, 10)
    return y === null ? ERR.NUM : y
  }),
  fn('AMORLINC', 6, 7, 'n', ([cost, bought, first, salvage, per, rate, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (cost < 0 || salvage < 0 || salvage > cost || rate <= 0 || per < 0) return ERR.NUM
    const one = cost * rate
    const delta = cost - salvage
    const r0 = (yf(bought, first, bb, ctx) as number) * rate * cost
    const full = Math.floor((cost - salvage - r0) / one)
    const p = Math.trunc(per)
    let res = 0
    if (p === 0) res = r0
    else if (p <= full) res = one
    else if (p === full + 1) res = delta - one * full - r0
    return res > 0 ? res : 0
  }),
  fn('AMORDEGRC', 6, 7, 'n', ([cost, bought, first, salvage, per, rate, b = 0], ctx) => {
    const bb = checkBasis(b)
    if (isError(bb)) return bb
    if (cost < 0 || salvage < 0 || rate <= 0 || per < 0) return ERR.NUM
    const usePer = 1 / rate
    const coeff = usePer < 3 ? 1 : usePer < 5 ? 1.5 : usePer <= 6 ? 2 : 2.5
    const r = rate * coeff
    let n = Math.round((yf(bought, first, bb, ctx) as number) * r * cost)
    let c = cost - n
    let rest = c - salvage
    const p = Math.trunc(per)
    for (let k = 0; k < p; k++) {
      n = Math.round(r * c)
      rest -= n
      if (rest < 0) return p - k <= 1 ? Math.round(c * 0.5) : 0
      c -= n
    }
    return n
  })
])

/** ODDFPRICE with a short or long odd first coupon period (Microsoft's documented formulas). */
function oddFirstPrice(s: number, m: number, issue: number, first: number, rate: number, yld: number, red: number, f: number, b: number, d1904: boolean): number | CellError {
  const months = 12 / f
  const fd = serialToDate(first, d1904)
  const eom = fd.d === daysInMonth(fd.y, fd.m)
  const shift = (k: number) => {
    const d = addMonths(fd, k * months, eom)
    return dateSerial(d.y, d.m, d.d, d1904)
  }
  // Coupons after the first one up to maturity
  let N = 1
  while (shift(N) <= m) N++
  const cpn = (100 * rate) / f
  const E = (b === 1 ? first - shift(-1) : (b === 3 ? 365 : 360) / f) as number
  const DFC = dayCount(issue, first, b, d1904)
  if (DFC <= E) {
    // Short first period
    const A = dayCount(issue, s, b, d1904)
    const DSC = dayCount(s, first, b, d1904)
    let p = red / Math.pow(1 + yld / f, N - 1 + DSC / E)
    p += (cpn * (DFC / E)) / Math.pow(1 + yld / f, DSC / E)
    for (let k = 2; k <= N; k++) p += cpn / Math.pow(1 + yld / f, k - 1 + DSC / E)
    return p - cpn * (A / E)
  }
  // Long first period: quasi-coupon periods before the first coupon.
  let NC = 0
  while (shift(-NC - 1) > issue) NC++
  NC++
  let sumDC = 0
  let sumA = 0
  let nq = 0
  for (let i = 1; i <= NC; i++) {
    const end = shift(-(NC - i))
    const start = shift(-(NC - i) - 1)
    const NL = b === 1 ? end - start : (b === 3 ? 365 : 360) / f
    const from = Math.max(start, issue)
    const DC = dayCount(from, end, b, d1904)
    sumDC += DC / NL
    if (s > start) sumA += dayCount(from, Math.min(s, end), b, d1904) / NL
    if (end <= s) nq++
  }
  const nextQuasi = (() => {
    for (let i = -NC; i <= 0; i++) if (shift(i) > s) return shift(i)
    return first
  })()
  const DSC = dayCount(s, nextQuasi, b, d1904)
  const Nq = NC - nq - 1
  let p = red / Math.pow(1 + yld / f, N - 1 + Nq + 1 + DSC / E - 1)
  p += (cpn * sumDC) / Math.pow(1 + yld / f, Nq + DSC / E)
  for (let k = 1; k < N; k++) p += cpn / Math.pow(1 + yld / f, k + Nq + DSC / E)
  return p - cpn * sumA
}

export { yearBasis }
export type { Scalar }
