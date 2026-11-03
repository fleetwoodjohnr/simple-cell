// Shared building blocks for worksheet functions.
import { areaArray, eachInArea, toArray, type EvalContext, type FnSpec } from '../evaluate'
import {
  CellError,
  ERR,
  compareScalars,
  isArray,
  isError,
  isLambda,
  isRef,
  makeArray,
  parseNumber,
  round15,
  toNumber,
  toText,
  type ArrayValue,
  type Scalar,
  type Value
} from '../values'
import { parseDateTime } from '../../numfmt/dates'

export type Fn = FnSpec['fn']
export const fn = (name: string, min: number, max: number, params: string, f: Fn, extra: Partial<FnSpec> = {}): FnSpec => ({ name, min, max, params, fn: f, ...extra })

/** Unlimited argument count (Excel allows 255). */
export const MANY = 255

export interface NumOpts {
  /** Text and logicals inside references/arrays count (the …A functions: AVERAGEA, MAXA…): text as 0, TRUE as 1. */
  a?: boolean
  /** Errors are skipped instead of returned (AGGREGATE options). */
  skipErrors?: boolean
  /** Rows hidden by a filter or by the user are skipped (SUBTOTAL 101-111, AGGREGATE). */
  skipHidden?: 'filtered' | 'all'
  /** Cells holding SUBTOTAL/AGGREGATE formulas are skipped. */
  skipSubtotals?: boolean
}

/**
 * The numbers among the arguments, with Excel's aggregate rules: values typed directly into the function are coerced
 * (TRUE → 1, "3" → 3, other text → #VALUE!); inside references and arrays only numbers count.
 */
export function numbers(args: Value[], ctx: EvalContext, opts: NumOpts = {}): number[] | CellError {
  const out: number[] = []
  for (const v of args) {
    if (v === undefined) continue
    if (isRef(v)) {
      for (const a of v.areas) {
        let err: CellError | null = null
        const sheet = ctx.wb.sheetById(a.sheet)
        eachInArea(a, ctx, (x, r, c) => {
          if (err) return
          if (opts.skipHidden && sheet) {
            const ri = sheet.rowInfo.get(r)
            if (ri?.hidden && (opts.skipHidden === 'all' || ri.filtered)) return
          }
          if (opts.skipSubtotals) {
            const f = sheet?.get(r, c)?.f
            if (f && isSubtotalFormula(f.ast)) return
          }
          if (typeof x === 'number') out.push(x)
          else if (x instanceof CellError) {
            if (!opts.skipErrors) err = x
          } else if (opts.a) out.push(typeof x === 'boolean' ? (x ? 1 : 0) : 0)
        })
        if (err) return err
      }
    } else if (isArray(v)) {
      for (const x of v.data) {
        if (typeof x === 'number') out.push(x)
        else if (x instanceof CellError) {
          if (!opts.skipErrors) return x
        } else if (opts.a && x !== null) out.push(typeof x === 'boolean' ? (x ? 1 : 0) : 0)
      }
    } else if (isLambda(v)) return ERR.VALUE
    else {
      if (v === null) {
        out.push(0)
        continue
      }
      const n = toNumber(v)
      if (isError(n)) {
        if (opts.skipErrors && v instanceof CellError) continue
        return n
      }
      out.push(n)
    }
  }
  return out
}

function isSubtotalFormula(ast: unknown): boolean {
  const n = ast as { t: string; name?: string }
  return n.t === 'func' && (n.name === 'SUBTOTAL' || n.name === 'AGGREGATE')
}

/** Every value of the arguments (references read, arrays flattened), for COUNTA and similar. */
export function allValues(args: Value[], ctx: EvalContext, includeEmpty = false): Scalar[] | CellError {
  const out: Scalar[] = []
  for (const v of args) {
    if (v === undefined) continue
    if (isRef(v)) {
      for (const a of v.areas) {
        if (includeEmpty) out.push(...areaArray(a, ctx).data)
        else eachInArea(a, ctx, (x) => out.push(x))
      }
    } else if (isArray(v)) out.push(...v.data)
    else if (isLambda(v)) return ERR.VALUE
    else out.push(v)
  }
  return out
}

/** A 2-D argument as an array, or the error. */
export function grid(v: Value, ctx: EvalContext): ArrayValue | CellError {
  if (v === undefined) return ERR.VALUE
  return toArray(v, ctx)
}

export const num = (v: Scalar) => {
  const n = toNumber(v)
  return n
}

export function text(v: Scalar): string | CellError {
  return toText(v)
}

/** Round half away from zero to `digits` places, avoiding binary artefacts (ROUND(2.675, 2) = 2.68 in Excel). */
export function roundHalfAway(x: number, digits: number): number {
  if (!Number.isFinite(x)) return x
  const d = Math.trunc(digits)
  if (d > 15) return x
  if (d < -15) return 0
  const s = Math.sign(x) || 1
  const a = Math.abs(x)
  const p = Math.pow(10, Math.abs(d))
  // Work on the 15-significant-digit value Excel shows, then round.
  const v = round15(a)
  const r = d >= 0 ? Math.round(Number((v * p).toPrecision(15))) / p : Math.round(v / p) * p
  return s * r
}

// ---------- criteria (COUNTIF, SUMIF, …, database functions) ----------

const wildcardCache = new Map<string, RegExp>()
/** Excel wildcards: * any run, ? one character, ~ escapes. Case-insensitive, whole-string. */
export function wildcardRegex(pattern: string): RegExp {
  let re = wildcardCache.get(pattern)
  if (re) return re
  let src = ''
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) {
      src += escapeRe(pattern[++i])
    } else if (ch === '*') src += '[\\s\\S]*'
    else if (ch === '?') src += '[\\s\\S]'
    else src += escapeRe(ch)
  }
  re = new RegExp(`^${src}$`, 'i')
  if (wildcardCache.size > 500) wildcardCache.clear()
  wildcardCache.set(pattern, re)
  return re
}
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')
export const hasWildcard = (s: string) => /(^|[^~])[*?]/.test(s)

/** A test for one criteria value: 5, ">5", "<>apple", "a*", "=", "", TRUE, or an error value. */
export function criteria(c: Scalar): (v: Scalar) => boolean {
  if (c === null) c = ''
  if (typeof c === 'number') return (v) => (typeof v === 'number' ? round15(v) === round15(c as number) : typeof v === 'string' && parseNumber(v, parseDateTime) === c)
  if (typeof c === 'boolean') return (v) => v === c
  if (c instanceof CellError) return (v) => v === c
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!
  const op = m[1] ?? ''
  const rhs = m[2]
  // Numbers (and dates typed as text) in the criteria compare numerically.
  const n = rhs === '' ? null : parseNumber(rhs, parseDateTime)
  const upper = rhs.toUpperCase()
  const bool = upper === 'TRUE' ? true : upper === 'FALSE' ? false : null
  const err = /^#/.test(rhs) ? (CellError.all.get(rhs.toUpperCase() as never) ?? null) : null
  if (op === '' || op === '=') {
    if (rhs === '') return op === '=' ? (v) => v === null : (v) => v === null || v === ''
    if (n !== null) return (v) => typeof v === 'number' && round15(v) === round15(n)
    if (bool !== null) return (v) => v === bool
    if (err) return (v) => v === err
    if (hasWildcard(rhs) || rhs.includes('~')) {
      const re = wildcardRegex(rhs)
      return (v) => typeof v === 'string' && re.test(v)
    }
    const low = rhs.toLowerCase()
    return (v) => typeof v === 'string' && v.toLowerCase() === low
  }
  if (op === '<>') {
    if (rhs === '') return (v) => v !== null && v !== ''
    if (n !== null) return (v) => !(typeof v === 'number' && round15(v) === round15(n))
    if (bool !== null) return (v) => v !== bool
    if (err) return (v) => v !== err
    const re = wildcardRegex(rhs)
    return (v) => !(typeof v === 'string' && re.test(v))
  }
  const test = (cmp: number) => (op === '<' ? cmp < 0 : op === '>' ? cmp > 0 : op === '<=' ? cmp <= 0 : cmp >= 0)
  if (n !== null) return (v) => typeof v === 'number' && test(compareScalars(v, n))
  if (bool !== null) return (v) => typeof v === 'boolean' && test(compareScalars(v, bool))
  return (v) => typeof v === 'string' && test(compareScalars(v, rhs))
}

/** Values of a range argument (for COUNTIF's range, SUMIF's sum range…) as an array, keeping its shape. */
export function rangeValues(v: Value, ctx: EvalContext): ArrayValue | CellError {
  if (isRef(v)) {
    if (v.areas.length !== 1) return ERR.VALUE
    return areaArray(v.areas[0], ctx)
  }
  if (isArray(v)) return v
  return ERR.VALUE
}

/** Indexes of cells matching every (range, criteria) pair; ranges must have equal shapes. */
export function matchIndexes(pairs: [Value, Value][], ctx: EvalContext, shape?: { rows: number; cols: number }): { idx: number[]; rows: number; cols: number } | CellError {
  let rows = shape?.rows ?? -1
  let cols = shape?.cols ?? -1
  let idx: number[] | null = null
  for (const [rv, cv] of pairs) {
    const range = rangeValues(rv, ctx)
    if (isError(range)) return range
    if (rows < 0) {
      rows = range.rows
      cols = range.cols
    } else if (range.rows !== rows || range.cols !== cols) return ERR.VALUE
    const c = isRef(cv) || isArray(cv) ? toFirst(cv, ctx) : (cv as Scalar)
    const test = criteria(c)
    const next: number[] = []
    if (idx === null) {
      for (let i = 0; i < range.data.length; i++) if (test(range.data[i])) next.push(i)
    } else for (const i of idx) if (test(range.data[i])) next.push(i)
    idx = next
  }
  return { idx: idx ?? [], rows, cols }
}

function toFirst(v: Value, ctx: EvalContext): Scalar {
  const a = toArray(v, ctx)
  return isError(a) ? a : a.data[0]
}

/** Shape-checked numeric 2-D argument (matrix functions). */
export function matrix(v: Value, ctx: EvalContext): number[][] | CellError {
  const a = grid(v, ctx)
  if (isError(a)) return a
  const m: number[][] = []
  for (let r = 0; r < a.rows; r++) {
    const row: number[] = []
    for (let c = 0; c < a.cols; c++) {
      const x = a.data[r * a.cols + c]
      if (x instanceof CellError) return x
      if (typeof x !== 'number') return ERR.VALUE
      row.push(x)
    }
    m.push(row)
  }
  return m
}

export const fromMatrix = (m: number[][]): ArrayValue => makeArray(m.length, m[0].length, m.flat())

/** An array of values from a flat list, as one column. */
export const column = (xs: Scalar[]): ArrayValue | CellError => (xs.length ? makeArray(xs.length, 1, xs) : ERR.CALC)

/** Positive integer argument or #VALUE!/#NUM! as Excel validates them. */
export function int(v: number | undefined, dflt: number): number {
  return v === undefined ? dflt : Math.trunc(v)
}

export const checkNum = (x: number): number | CellError => (Number.isFinite(x) ? x : ERR.NUM)

export function sum(xs: number[]): number {
  // Kahan summation keeps long sums of decimals stable.
  let s = 0
  let comp = 0
  for (const x of xs) {
    const y = x - comp
    const t = s + y
    comp = t - s - y
    s = t
  }
  return s
}

export const mean = (xs: number[]) => sum(xs) / xs.length

export function variance(xs: number[], sample: boolean): number | CellError {
  const n = xs.length
  if (n < (sample ? 2 : 1)) return ERR.DIV0
  const m = mean(xs)
  let ss = 0
  for (const x of xs) ss += (x - m) * (x - m)
  return ss / (sample ? n - 1 : n)
}

export { ERR, isError, isArray, isRef, makeArray, toNumber, toText, CellError }
export type { Scalar, Value, ArrayValue, EvalContext }
