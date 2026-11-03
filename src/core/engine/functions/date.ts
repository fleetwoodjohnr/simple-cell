// Date and time functions.
import { defineFunctions, type EvalContext } from '../evaluate'
import { ERR, isError, toNumber, type CellError, type Value } from '../values'
import { dateSerial, daysInMonth, isLeapYear, nowSerial, parseDateTime, serialToDate, timeParts, weekday, type YMD } from '../../numfmt/dates'
import { allValues, fn } from './helpers'

const MAX_SERIAL = 2958465 // 9999-12-31

function checkSerial(x: number): number | CellError {
  return x < 0 || x > MAX_SERIAL + 0.99999999 ? ERR.NUM : x
}

const ymd = (x: number, ctx: EvalContext) => serialToDate(x, ctx.wb.date1904)
const ser = (y: number, m: number, d: number, ctx: EvalContext) => dateSerial(y, m, d, ctx.wb.date1904)

export function days360(a: YMD, b: YMD, european: boolean, yearfrac = false): number {
  let sd = a.d
  let ed = b.d
  if (european) {
    sd = Math.min(sd, 30)
    ed = Math.min(ed, 30)
  } else {
    const lastFebA = a.m === 2 && a.d === daysInMonth(a.y, 2)
    const lastFebB = b.m === 2 && b.d === daysInMonth(b.y, 2)
    if (yearfrac && lastFebA && lastFebB) ed = 30
    if (lastFebA) sd = 30
    if (ed === 31 && sd >= 30) ed = 30
    if (sd === 31) sd = 30
  }
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (ed - sd)
}

export function yearFrac(s: number, e: number, basis: number, date1904: boolean): number | CellError {
  if (s > e) [s, e] = [e, s]
  s = Math.floor(s)
  e = Math.floor(e)
  const a = serialToDate(s, date1904)
  const b = serialToDate(e, date1904)
  switch (basis) {
    case 0:
      return days360(a, b, false, true) / 360
    case 1: {
      const feb29Between = () => {
        const mar1a = dateSerial(a.y, 3, 1, date1904)
        if (isLeapYear(a.y) && s < mar1a && e >= mar1a) return true
        const mar1b = dateSerial(b.y, 3, 1, date1904)
        return isLeapYear(b.y) && e >= mar1b && s < mar1b
      }
      if (a.y === b.y || (a.y + 1 === b.y && (a.m > b.m || (a.m === b.m && a.d >= b.d)))) {
        const len = (a.y === b.y && isLeapYear(a.y)) || feb29Between() || (b.m === 2 && b.d === 29) ? 366 : 365
        return (e - s) / len
      }
      const years = b.y - a.y + 1
      const days = dateSerial(b.y + 1, 1, 1, date1904) - dateSerial(a.y, 1, 1, date1904)
      return (e - s) / (days / years)
    }
    case 2:
      return (e - s) / 360
    case 3:
      return (e - s) / 365
    case 4:
      return days360(a, b, true) / 360
  }
  return ERR.NUM
}

/** Weekend mask Monday..Sunday from a NETWORKDAYS.INTL / WORKDAY.INTL weekend argument. */
function weekendMask(w: Value | undefined): boolean[] | CellError {
  if (w === undefined || w === null) return [false, false, false, false, false, true, true]
  if (typeof w === 'string') {
    if (!/^[01]{7}$/.test(w) || w === '1111111') return ERR.VALUE
    return [...w].map((c) => c === '1')
  }
  const n = toNumber(w as never)
  if (isError(n)) return n
  const codes: Record<number, number[]> = { 1: [5, 6], 2: [6, 0], 3: [0, 1], 4: [1, 2], 5: [2, 3], 6: [3, 4], 7: [4, 5], 11: [6], 12: [0], 13: [1], 14: [2], 15: [3], 16: [4], 17: [5] }
  const days = codes[Math.trunc(n)]
  if (!days) return ERR.NUM
  const mask = [false, false, false, false, false, false, false]
  for (const d of days) mask[d] = true
  return mask
}

/** Monday-first weekday index 0..6. */
const monFirst = (x: number, date1904: boolean) => (weekday(x, date1904) + 6) % 7

/** Holidays: numbers, or date text (Excel accepts {"2006/1/2","2006/1/16"}). Empty cells are skipped. */
function holidaySet(v: Value | undefined, ctx: EvalContext): Set<number> | CellError {
  if (v === undefined || v === null) return new Set()
  const vals = allValues([v], ctx)
  if (isError(vals)) return vals
  const out = new Set<number>()
  for (const x of vals) {
    if (x === null) continue
    const n = toNumber(x)
    if (isError(n)) return n
    out.add(Math.floor(n))
  }
  return out
}

function networkDays(s: number, e: number, mask: boolean[], hol: Set<number>, d1904: boolean): number {
  s = Math.floor(s)
  e = Math.floor(e)
  const sign = s <= e ? 1 : -1
  if (sign < 0) [s, e] = [e, s]
  let count = 0
  // Whole weeks first, then the remainder.
  const workPerWeek = mask.filter((x) => !x).length
  const weeks = Math.floor((e - s + 1) / 7)
  count += weeks * workPerWeek
  for (let d = s + weeks * 7; d <= e; d++) if (!mask[monFirst(d, d1904)]) count++
  for (const h of hol) if (h >= s && h <= e && !mask[monFirst(h, d1904)]) count--
  return sign * count
}

function workday(s: number, days: number, mask: boolean[], hol: Set<number>, d1904: boolean): number | CellError {
  if (mask.every(Boolean)) return ERR.VALUE
  let d = Math.floor(s)
  let n = Math.trunc(days)
  const step = n >= 0 ? 1 : -1
  while (n !== 0) {
    d += step
    if (d < 0 || d > MAX_SERIAL) return ERR.NUM
    if (!mask[monFirst(d, d1904)] && !hol.has(d)) n -= step
  }
  return d
}

defineFunctions([
  fn('DATE', 3, 3, 'n', ([y, m, d], ctx) => {
    y = Math.trunc(y)
    if (y < 0 || y >= 10000) return ERR.NUM
    if (y < 1900) y += 1900
    const s = ser(y, Math.trunc(m), Math.trunc(d), ctx)
    return checkSerial(s)
  }),
  fn('TIME', 3, 3, 'n', ([h, m, s]) => {
    const t = Math.trunc(h) * 3600 + Math.trunc(m) * 60 + Math.trunc(s)
    if (t < 0) return ERR.NUM
    return (t % 86400) / 86400
  }),
  fn('DATEVALUE', 1, 1, 's', ([s], ctx) => {
    const v = parseDateTime(s, ctx.wb.date1904)
    return v === null ? ERR.VALUE : Math.floor(v)
  }),
  fn('TIMEVALUE', 1, 1, 's', ([s], ctx) => {
    const v = parseDateTime(s, ctx.wb.date1904)
    return v === null ? ERR.VALUE : v - Math.floor(v)
  }),
  fn('YEAR', 1, 1, 'n', ([x], ctx) => (x < 0 ? ERR.NUM : ymd(x, ctx).y)),
  fn('MONTH', 1, 1, 'n', ([x], ctx) => (x < 0 ? ERR.NUM : ymd(x, ctx).m)),
  fn('DAY', 1, 1, 'n', ([x], ctx) => (x < 0 ? ERR.NUM : ymd(x, ctx).d)),
  fn('HOUR', 1, 1, 'n', ([x]) => (x < 0 ? ERR.NUM : timeParts(x).h)),
  fn('MINUTE', 1, 1, 'n', ([x]) => (x < 0 ? ERR.NUM : timeParts(x).mi)),
  fn('SECOND', 1, 1, 'n', ([x]) => (x < 0 ? ERR.NUM : timeParts(x).s + (timeParts(x).ms >= 500 ? 1 : 0))),
  fn('NOW', 0, 0, '', (_a, ctx) => nowSerial(ctx.wb.date1904), { volatile: true }),
  fn('TODAY', 0, 0, '', (_a, ctx) => Math.floor(nowSerial(ctx.wb.date1904)), { volatile: true }),
  fn('DAYS', 2, 2, 'n', ([e, s]) => Math.floor(e) - Math.floor(s)),
  fn('DAYS360', 2, 3, 'nnb', ([s, e, eu = false], ctx) => days360(ymd(s, ctx), ymd(e, ctx), eu)),
  fn('EDATE', 2, 2, 'n', ([s, m], ctx) => {
    if (s < 0) return ERR.NUM
    const d = ymd(s, ctx)
    const t = d.m + Math.trunc(m)
    const y = d.y + Math.floor((t - 1) / 12)
    const mm = ((((t - 1) % 12) + 12) % 12) + 1
    return checkSerial(ser(y, mm, Math.min(d.d, daysInMonth(y, mm)), ctx))
  }),
  fn('EOMONTH', 2, 2, 'n', ([s, m], ctx) => {
    if (s < 0) return ERR.NUM
    const d = ymd(s, ctx)
    const t = d.m + Math.trunc(m)
    const y = d.y + Math.floor((t - 1) / 12)
    const mm = ((((t - 1) % 12) + 12) % 12) + 1
    return checkSerial(ser(y, mm, daysInMonth(y, mm), ctx))
  }),
  fn('DATEDIF', 3, 3, 'nns', ([s, e, unit], ctx) => {
    s = Math.floor(s)
    e = Math.floor(e)
    if (s > e) return ERR.NUM
    const a = ymd(s, ctx)
    const b = ymd(e, ctx)
    let months = (b.y - a.y) * 12 + (b.m - a.m)
    if (b.d < a.d) months--
    switch (unit.toUpperCase()) {
      case 'Y':
        return Math.floor(months / 12)
      case 'M':
        return months
      case 'D':
        return e - s
      case 'MD': {
        if (b.d >= a.d) return b.d - a.d
        // days from the same day in the previous month
        const pm = b.m === 1 ? 12 : b.m - 1
        const py = b.m === 1 ? b.y - 1 : b.y
        return e - ser(py, pm, a.d, ctx)
      }
      case 'YM':
        return ((months % 12) + 12) % 12
      case 'YD': {
        let start = ser(b.y, a.m, a.d, ctx)
        if (start > e) start = ser(b.y - 1, a.m, a.d, ctx)
        return e - start
      }
    }
    return ERR.NUM
  }),
  fn('WEEKDAY', 1, 2, 'n', ([x, type = 1], ctx) => {
    if (x < 0) return ERR.NUM
    const w = weekday(x, ctx.wb.date1904) // 0 = Sunday
    const t = Math.trunc(type)
    if (t === 1 || t === 17) return w + 1
    if (t === 2 || t === 11) return ((w + 6) % 7) + 1
    if (t === 3) return (w + 6) % 7
    if (t >= 12 && t <= 16) return ((w - (t - 10) + 7) % 7) + 1
    return ERR.NUM
  }),
  fn('WEEKNUM', 1, 2, 'n', ([x, type = 1], ctx) => {
    if (x < 0) return ERR.NUM
    const t = Math.trunc(type)
    if (t === 21) return isoWeek(x, ctx)
    const startDay: Record<number, number> = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }
    const sd = startDay[t]
    if (sd === undefined) return ERR.NUM
    const d = ymd(x, ctx)
    const jan1 = ser(d.y, 1, 1, ctx)
    const offset = (weekday(jan1, ctx.wb.date1904) - sd + 7) % 7
    return Math.floor((Math.floor(x) - jan1 + offset) / 7) + 1
  }),
  fn('ISOWEEKNUM', 1, 1, 'n', ([x], ctx) => (x < 0 ? ERR.NUM : isoWeek(x, ctx))),
  fn('NETWORKDAYS', 2, 3, 'nnr', ([s, e, hol], ctx) => {
    const h = holidaySet(hol, ctx)
    return isError(h) ? h : networkDays(s, e, [false, false, false, false, false, true, true], h, ctx.wb.date1904)
  }),
  fn('NETWORKDAYS.INTL', 2, 4, 'nnvr', ([s, e, w, hol], ctx) => {
    const m = weekendMask(w)
    if (isError(m)) return m
    const h = holidaySet(hol, ctx)
    return isError(h) ? h : networkDays(s, e, m, h, ctx.wb.date1904)
  }),
  fn('WORKDAY', 2, 3, 'nnr', ([s, d, hol], ctx) => {
    const h = holidaySet(hol, ctx)
    return isError(h) ? h : workday(s, d, [false, false, false, false, false, true, true], h, ctx.wb.date1904)
  }),
  fn('WORKDAY.INTL', 2, 4, 'nnvr', ([s, d, w, hol], ctx) => {
    const m = weekendMask(w)
    if (isError(m)) return m
    const h = holidaySet(hol, ctx)
    return isError(h) ? h : workday(s, d, m, h, ctx.wb.date1904)
  }),
  fn('YEARFRAC', 2, 3, 'n', ([s, e, basis = 0], ctx) => {
    const b = Math.trunc(basis)
    if (b < 0 || b > 4 || s < 0 || e < 0) return ERR.NUM
    return yearFrac(s, e, b, ctx.wb.date1904)
  })
])

function isoWeek(x: number, ctx: EvalContext): number {
  const d = Math.floor(x)
  const wd = (weekday(d, ctx.wb.date1904) + 6) % 7 // Monday = 0
  const thursday = d - wd + 3
  const y = ymd(thursday, ctx).y
  const jan1 = ser(y, 1, 1, ctx)
  return Math.floor((thursday - jan1) / 7) + 1
}
