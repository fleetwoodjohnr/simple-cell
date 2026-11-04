// Analysis ToolPak (Data ▸ Data Analysis): Excel's statistical tools, written as values in Excel's layouts.
// Probabilities and critical values come from the worksheet functions (T.DIST.2T, F.INV.RT, LINEST…), so they agree
// with the cells' own results; those functions are tested against Microsoft's examples and LibreOffice.

import type { Doc, Tx } from './doc'
import { evaluate, finalValue } from './engine/evaluate'
import { CellError, numberToText, type Area, type Scalar } from './engine/values'
import { areaName, quoteSheet } from './refs'

export type ToolId =
  | 'anova1'
  | 'anova2r'
  | 'anova2'
  | 'correlation'
  | 'covariance'
  | 'expsmooth'
  | 'ftest'
  | 'fourier'
  | 'histogram'
  | 'movavg'
  | 'random'
  | 'rank'
  | 'regression'
  | 'sampling'
  | 'ttest-paired'
  | 'ttest-equal'
  | 'ttest-unequal'
  | 'ztest'

export interface ToolParams {
  input?: Area
  input2?: Area
  /** Histogram bins, Regression X range. */
  extra?: Area
  labels?: boolean
  byRows?: boolean
  alpha?: number
  /** Hypothesized mean difference (t, z), interval (moving average), damping (smoothing), rows per sample (ANOVA). */
  n?: number
  var1?: number
  var2?: number
  pareto?: boolean
  cumulative?: boolean
  chart?: boolean
  stdErrors?: boolean
  confidence?: number
  zeroConstant?: boolean
  residuals?: boolean
  inverse?: boolean
  /** Sampling: 'periodic' with period n, or 'random' with count n. */
  method?: 'periodic' | 'random'
  /** Random numbers: variables, count, distribution and its parameters, seed. */
  variables?: number
  count?: number
  distribution?: 'uniform' | 'normal' | 'bernoulli' | 'binomial' | 'poisson'
  p1?: number
  p2?: number
  seed?: number
  /** Where the output goes: a cell on a sheet, or a new sheet with this name. */
  out: { sheet: number; r: number; c: number } | { newSheet: string }
}

type Out = (Scalar | undefined)[][]

// ---------- inputs ----------

/** Numbers of an area, grouped by columns (or rows); the first cell of each group is its label when `labels`. */
function groups(doc: Doc, a: Area, byRows: boolean, labels: boolean, prefix = byRows ? 'Row' : 'Column'): { name: string; xs: number[] }[] {
  const sh = doc.sheet(a.sheet)
  const u = sh.used()
  const r1 = Math.min(a.r1, u.r)
  const c1 = Math.min(a.c1, u.c)
  const out: { name: string; xs: number[] }[] = []
  const n = byRows ? r1 - a.r0 + 1 : c1 - a.c0 + 1
  const len = byRows ? c1 - a.c0 + 1 : r1 - a.r0 + 1
  for (let g = 0; g < n; g++) {
    const at = (k: number) => (byRows ? sh.get(a.r0 + g, a.c0 + k) : sh.get(a.r0 + k, a.c0 + g))?.v ?? null
    const xs: number[] = []
    for (let k = labels ? 1 : 0; k < len; k++) {
      const v = at(k)
      if (typeof v === 'number') xs.push(v)
      else if (v !== null && v !== '') throw new Error(`The input range contains text or errors (${typeof v === 'string' ? `“${v}”` : String(v)}): it must contain only numbers${labels ? '' : ' (tick Labels if the first row or column is headings)'}.`)
    }
    const label = labels ? at(0) : null
    out.push({ name: label !== null && label !== '' ? String(label instanceof CellError ? label.code : label) : `${prefix} ${g + 1}`, xs })
  }
  return out
}

const one = (doc: Doc, a: Area | undefined, labels: boolean, what: string) => {
  if (!a) throw new Error(`Enter the ${what}.`)
  const g = groups(doc, { ...a, c1: a.c0 }, false, labels, 'Variable')[0]
  if (a.c1 > a.c0 && a.r1 > a.r0) throw new Error(`The ${what} must be a single row or column.`)
  return a.c1 > a.c0 ? groups(doc, { ...a, r1: a.r0 }, true, labels, 'Variable')[0] : g
}

const sum = (xs: number[]) => xs.reduce((s, x) => s + x, 0)
const mean = (xs: number[]) => sum(xs) / xs.length
const sq = (xs: number[], m = mean(xs)) => xs.reduce((s, x) => s + (x - m) ** 2, 0)
const variance = (xs: number[]) => (xs.length > 1 ? sq(xs) / (xs.length - 1) : NaN)

// ---------- worksheet functions ----------

/** Evaluate a worksheet function with numeric arguments (full precision), e.g. fx(doc, 'T.DIST.2T', 2.1, 10). */
function fx(doc: Doc, name: string, ...args: (number | string)[]): number {
  const s = doc.wb.sheets[0].id
  const text = `=${name}(${args.map((a) => (typeof a === 'number' ? (Number.isFinite(a) ? String(a).toUpperCase() : 'NA()') : a)).join(',')})`
  const ast = doc.engine.parseAt(text, s, 0, 0)
  const ctx = doc.engine.contextAt(s, 0, 0)
  const v = finalValue(evaluate(ast, ctx), ctx, false)
  return typeof v === 'number' ? v : NaN
}

/** A worksheet array function over references (LINEST): its result rows. */
function fxArray(doc: Doc, text: string): Scalar[][] {
  const s = doc.wb.sheets[0].id
  const ast = doc.engine.parseAt('=' + text, s, 0, 0)
  const ctx = doc.engine.contextAt(s, 0, 0)
  const v = finalValue(evaluate(ast, ctx), ctx, true)
  if (typeof v !== 'object' || v === null || !('kind' in v) || v.kind !== 'array') throw new Error(`${text.split('(')[0]} failed: ${v instanceof CellError ? v.code : String(v)}`)
  return Array.from({ length: v.rows }, (_, r) => v.data.slice(r * v.cols, (r + 1) * v.cols))
}

const refText = (doc: Doc, a: Area) => `${quoteSheet(doc.sheetName(a.sheet) ?? '')}!${areaName(a)}`

// ---------- the tools ----------

function anova1(doc: Doc, p: ToolParams): Out {
  const gs = groups(doc, need(p.input, 'input range'), !!p.byRows, !!p.labels)
  const alpha = p.alpha ?? 0.05
  const all = gs.flatMap((g) => g.xs)
  const grand = mean(all)
  const ssb = sum(gs.map((g) => g.xs.length * (mean(g.xs) - grand) ** 2))
  const ssw = sum(gs.map((g) => sq(g.xs)))
  const dfb = gs.length - 1
  const dfw = all.length - gs.length
  const F = ssb / dfb / (ssw / dfw)
  return [
    ['Anova: Single Factor'],
    [],
    ['SUMMARY'],
    ['Groups', 'Count', 'Sum', 'Average', 'Variance'],
    ...gs.map((g) => [g.name, g.xs.length, sum(g.xs), mean(g.xs), variance(g.xs)]),
    [],
    [],
    ['ANOVA'],
    ['Source of Variation', 'SS', 'df', 'MS', 'F', 'P-value', 'F crit'],
    ['Between Groups', ssb, dfb, ssb / dfb, F, fx(doc, 'F.DIST.RT', F, dfb, dfw), fx(doc, 'F.INV.RT', alpha, dfb, dfw)],
    ['Within Groups', ssw, dfw, ssw / dfw],
    [],
    ['Total', ssb + ssw, all.length - 1]
  ]
}

/** Two-factor ANOVA without replication: rows are one factor, columns the other (one value per cell). */
function anova2(doc: Doc, p: ToolParams): Out {
  const a = need(p.input, 'input range')
  const labels = !!p.labels
  const cols = groups(doc, labels ? { ...a, r0: a.r0, c0: a.c0 + 1 } : a, false, labels)
  const rows = groups(doc, labels ? { ...a, r0: a.r0 + 1 } : a, true, labels)
  const alpha = p.alpha ?? 0.05
  const R = rows.length
  const C = cols.length
  if (rows.some((r) => r.xs.length !== C)) throw new Error('Every row needs a value in every column.')
  const all = rows.flatMap((r) => r.xs)
  const grand = mean(all)
  const ssr = C * sum(rows.map((r) => (mean(r.xs) - grand) ** 2))
  const ssc = R * sum(cols.map((c) => (mean(c.xs) - grand) ** 2))
  const sst = sq(all)
  const sse = sst - ssr - ssc
  const dfr = R - 1
  const dfc = C - 1
  const dfe = dfr * dfc
  const mse = sse / dfe
  const line = (name: string, ss: number, df: number) => [name, ss, df, ss / df, ss / df / mse, fx(doc, 'F.DIST.RT', ss / df / mse, df, dfe), fx(doc, 'F.INV.RT', alpha, df, dfe)]
  return [
    ['Anova: Two-Factor Without Replication'],
    [],
    ['SUMMARY', 'Count', 'Sum', 'Average', 'Variance'],
    ...rows.map((r) => [r.name, r.xs.length, sum(r.xs), mean(r.xs), variance(r.xs)]),
    [],
    ...cols.map((c) => [c.name, c.xs.length, sum(c.xs), mean(c.xs), variance(c.xs)]),
    [],
    [],
    ['ANOVA'],
    ['Source of Variation', 'SS', 'df', 'MS', 'F', 'P-value', 'F crit'],
    line('Rows', ssr, dfr),
    line('Columns', ssc, dfc),
    ['Error', sse, dfe, mse],
    [],
    ['Total', sst, R * C - 1]
  ]
}

/** Two-factor ANOVA with replication: labels in the first row and column; `n` rows per sample. */
function anova2r(doc: Doc, p: ToolParams): Out {
  const a = need(p.input, 'input range')
  const per = Math.round(p.n ?? 0)
  if (per < 2) throw new Error('Rows per sample must be at least 2.')
  const sh = doc.sheet(a.sheet)
  const cols = a.c1 - a.c0
  const body = a.r1 - a.r0
  if (cols < 1 || body < per || body % per) throw new Error(`The input range must have headings in its first row and column, and a multiple of ${per} rows of data.`)
  const samples = body / per
  const val = (r: number, c: number) => {
    const v = sh.get(r, c)?.v
    if (typeof v !== 'number') throw new Error('The input range must contain only numbers (apart from its headings).')
    return v
  }
  const label = (r: number, c: number, d: string) => {
    const v = sh.get(r, c)?.v
    return v === null || v === undefined || v === '' ? d : String(v)
  }
  const cell = (s: number, j: number) => Array.from({ length: per }, (_, k) => val(a.r0 + 1 + s * per + k, a.c0 + 1 + j))
  const colNames = Array.from({ length: cols }, (_, j) => label(a.r0, a.c0 + 1 + j, `Column ${j + 1}`))
  const all = Array.from({ length: samples }, (_, s) => Array.from({ length: cols }, (_, j) => cell(s, j))).flat(2)
  const grand = mean(all)
  const N = all.length
  const sampleXs = (s: number) => Array.from({ length: cols }, (_, j) => cell(s, j)).flat()
  const colXs = (j: number) => Array.from({ length: samples }, (_, s) => cell(s, j)).flat()
  const ssS = per * cols * sum(Array.from({ length: samples }, (_, s) => (mean(sampleXs(s)) - grand) ** 2))
  const ssC = per * samples * sum(Array.from({ length: cols }, (_, j) => (mean(colXs(j)) - grand) ** 2))
  const ssW = sum(Array.from({ length: samples }, (_, s) => sum(Array.from({ length: cols }, (_, j) => sq(cell(s, j))))))
  const ssT = sq(all)
  const ssI = ssT - ssS - ssC - ssW
  const dfS = samples - 1
  const dfC = cols - 1
  const dfI = dfS * dfC
  const dfW = N - samples * cols
  const msW = ssW / dfW
  const alpha = p.alpha ?? 0.05
  const line = (name: string, ss: number, df: number) => [name, ss, df, ss / df, ss / df / msW, fx(doc, 'F.DIST.RT', ss / df / msW, df, dfW), fx(doc, 'F.INV.RT', alpha, df, dfW)]
  const block = (name: string, xsOf: (j: number) => number[], withTotal: boolean): Out => {
    const xs = Array.from({ length: cols }, (_, j) => xsOf(j))
    const tot = xs.flat()
    const row = (lab: string, f: (v: number[]) => number) => [lab, ...xs.map(f), ...(withTotal ? [f(tot)] : [])]
    return [[name], row('Count', (v) => v.length), row('Sum', sum), row('Average', mean), row('Variance', variance), []]
  }
  return [
    ['Anova: Two-Factor With Replication'],
    [],
    ['SUMMARY', ...colNames, 'Total'],
    ...Array.from({ length: samples }, (_, s) => block(label(a.r0 + 1 + s * per, a.c0, `Sample ${s + 1}`), (j) => cell(s, j), true)).flat(),
    ...block('Total', colXs, false),
    [],
    ['ANOVA'],
    ['Source of Variation', 'SS', 'df', 'MS', 'F', 'P-value', 'F crit'],
    line('Sample', ssS, dfS),
    line('Columns', ssC, dfC),
    line('Interaction', ssI, dfI),
    ['Within', ssW, dfW, msW],
    [],
    ['Total', ssT, N - 1]
  ]
}

function matrix(doc: Doc, p: ToolParams, f: (a: number[], b: number[]) => number): Out {
  const gs = groups(doc, need(p.input, 'input range'), !!p.byRows, !!p.labels)
  if (gs.some((g) => g.xs.length !== gs[0].xs.length)) throw new Error('Every column must have the same number of values.')
  return [['', ...gs.map((g) => g.name)], ...gs.map((g, i) => [g.name, ...gs.map((h, j) => (j <= i ? f(g.xs, h.xs) : undefined))])]
}
const correl = (a: number[], b: number[]) => {
  const ma = mean(a)
  const mb = mean(b)
  return sum(a.map((x, i) => (x - ma) * (b[i] - mb))) / Math.sqrt(sq(a) * sq(b))
}
/** The ToolPak's covariance divides by n (COVARIANCE.P). */
const covar = (a: number[], b: number[]) => {
  const ma = mean(a)
  const mb = mean(b)
  return sum(a.map((x, i) => (x - ma) * (b[i] - mb))) / a.length
}

function twoSamples(doc: Doc, p: ToolParams) {
  const labels = !!p.labels
  return [one(doc, p.input, labels, 'Variable 1 range'), one(doc, p.input2, labels, 'Variable 2 range')] as const
}

function tTail(doc: Doc, t: number, df: number, alpha: number): Out {
  return [
    ['t Stat', t],
    ['P(T<=t) one-tail', fx(doc, 'T.DIST.RT', Math.abs(t), df)],
    ['t Critical one-tail', fx(doc, 'T.INV', 1 - alpha, df)],
    ['P(T<=t) two-tail', fx(doc, 'T.DIST.2T', Math.abs(t), df)],
    ['t Critical two-tail', fx(doc, 'T.INV.2T', alpha, df)]
  ]
}

function ttest(doc: Doc, p: ToolParams, kind: 'paired' | 'equal' | 'unequal'): Out {
  const [a, b] = twoSamples(doc, p)
  const alpha = p.alpha ?? 0.05
  const d0 = p.n ?? 0
  const head: Out = [[kind === 'paired' ? 't-Test: Paired Two Sample for Means' : kind === 'equal' ? 't-Test: Two-Sample Assuming Equal Variances' : 't-Test: Two-Sample Assuming Unequal Variances'], [], ['', a.name, b.name], ['Mean', mean(a.xs), mean(b.xs)], ['Variance', variance(a.xs), variance(b.xs)], ['Observations', a.xs.length, b.xs.length]]
  if (kind === 'paired') {
    if (a.xs.length !== b.xs.length) throw new Error('A paired test needs the same number of values in both ranges.')
    const d = a.xs.map((x, i) => x - b.xs[i])
    const n = d.length
    const t = (mean(d) - d0) / Math.sqrt(variance(d) / n)
    return [...head, ['Pearson Correlation', correl(a.xs, b.xs)], ['Hypothesized Mean Difference', d0], ['df', n - 1], ...tTail(doc, t, n - 1, alpha)]
  }
  const n1 = a.xs.length
  const n2 = b.xs.length
  const v1 = variance(a.xs)
  const v2 = variance(b.xs)
  if (kind === 'equal') {
    const pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2)
    const t = (mean(a.xs) - mean(b.xs) - d0) / Math.sqrt(pooled * (1 / n1 + 1 / n2))
    return [...head, ['Pooled Variance', pooled], ['Hypothesized Mean Difference', d0], ['df', n1 + n2 - 2], ...tTail(doc, t, n1 + n2 - 2, alpha)]
  }
  const se2 = v1 / n1 + v2 / n2
  // Welch–Satterthwaite, rounded to a whole number as the ToolPak does.
  const df = Math.round(se2 ** 2 / ((v1 / n1) ** 2 / (n1 - 1) + (v2 / n2) ** 2 / (n2 - 1)))
  const t = (mean(a.xs) - mean(b.xs) - d0) / Math.sqrt(se2)
  return [...head, ['Hypothesized Mean Difference', d0], ['df', df], ...tTail(doc, t, df, alpha)]
}

function ztest(doc: Doc, p: ToolParams): Out {
  const [a, b] = twoSamples(doc, p)
  if (!(p.var1! > 0) || !(p.var2! > 0)) throw new Error('Enter the known variances of both variables.')
  const alpha = p.alpha ?? 0.05
  const d0 = p.n ?? 0
  const z = (mean(a.xs) - mean(b.xs) - d0) / Math.sqrt(p.var1! / a.xs.length + p.var2! / b.xs.length)
  return [
    ['z-Test: Two Sample for Means'],
    [],
    ['', a.name, b.name],
    ['Mean', mean(a.xs), mean(b.xs)],
    ['Known Variance', p.var1, p.var2],
    ['Observations', a.xs.length, b.xs.length],
    ['Hypothesized Mean Difference', d0],
    ['z', z],
    ['P(Z<=z) one-tail', 1 - fx(doc, 'NORM.S.DIST', Math.abs(z), 'TRUE')],
    ['z Critical one-tail', fx(doc, 'NORM.S.INV', 1 - alpha)],
    ['P(Z<=z) two-tail', 2 * (1 - fx(doc, 'NORM.S.DIST', Math.abs(z), 'TRUE'))],
    ['z Critical two-tail', fx(doc, 'NORM.S.INV', 1 - alpha / 2)]
  ]
}

function ftest(doc: Doc, p: ToolParams): Out {
  const [a, b] = twoSamples(doc, p)
  const alpha = p.alpha ?? 0.05
  const df1 = a.xs.length - 1
  const df2 = b.xs.length - 1
  const F = variance(a.xs) / variance(b.xs)
  // As Excel: the tail on the side F falls.
  const right = F >= 1
  return [
    ['F-Test Two-Sample for Variances'],
    [],
    ['', a.name, b.name],
    ['Mean', mean(a.xs), mean(b.xs)],
    ['Variance', variance(a.xs), variance(b.xs)],
    ['Observations', a.xs.length, b.xs.length],
    ['df', df1, df2],
    ['F', F],
    ['P(F<=f) one-tail', right ? fx(doc, 'F.DIST.RT', F, df1, df2) : fx(doc, 'F.DIST', F, df1, df2, 'TRUE')],
    ['F Critical one-tail', right ? fx(doc, 'F.INV.RT', alpha, df1, df2) : fx(doc, 'F.INV', alpha, df1, df2)]
  ]
}

function histogram(doc: Doc, p: ToolParams): Out {
  const xs = groups(doc, need(p.input, 'input range'), false, !!p.labels).flatMap((g) => g.xs)
  let bins: number[]
  if (p.extra) bins = groups(doc, p.extra, false, !!p.labels).flatMap((g) => g.xs)
  else {
    // ponytail: evenly spaced bins, √n of them (rounded up), from the minimum to the maximum; Excel's own choice may differ.
    const k = Math.max(1, Math.ceil(Math.sqrt(xs.length)))
    const lo = Math.min(...xs)
    const hi = Math.max(...xs)
    bins = Array.from({ length: k }, (_, i) => lo + ((hi - lo) * i) / Math.max(1, k - 1))
  }
  bins = [...bins].sort((x, y) => x - y)
  const counts = bins.map((b, i) => xs.filter((x) => x <= b && (i === 0 || x > bins[i - 1])).length)
  const more = xs.filter((x) => x > bins[bins.length - 1]).length
  const rows: [string | number, number][] = [...bins.map((b, i): [number, number] => [b, counts[i]]), ['More', more]]
  const n = xs.length
  const withCum = (list: [string | number, number][]) => {
    let acc = 0
    return list.map(([b, f]) => {
      acc += f
      return p.cumulative ? [b, f, `${(Math.round((acc / n) * 10000) / 100).toFixed(2)}%`] : [b, f]
    })
  }
  const head = p.cumulative ? ['Bin', 'Frequency', 'Cumulative %'] : ['Bin', 'Frequency']
  const out: Out = [head, ...withCum(rows)]
  if (p.pareto) {
    const sorted = [...rows].sort((x, y) => y[1] - x[1])
    const right = [p.cumulative ? ['Bin', 'Frequency', 'Cumulative %'] : ['Bin', 'Frequency'], ...withCum(sorted)]
    return out.map((row, i) => [...row, undefined, ...right[i]])
  }
  return out
}

function movingAverage(doc: Doc, p: ToolParams): Out {
  const xs = one(doc, p.input, !!p.labels, 'input range').xs
  const k = Math.round(p.n ?? 3)
  if (k < 2 || k > xs.length) throw new Error(`The interval must be from 2 to the number of values (${xs.length}).`)
  const avg = xs.map((_, i) => (i + 1 < k ? null : mean(xs.slice(i + 1 - k, i + 1))))
  return xs.map((_, i) => {
    const row: Scalar[] = [avg[i] ?? CellError.of('#N/A')]
    if (p.stdErrors) {
      // Excel: √(Σ (actual − forecast)² / interval) over the last `interval` forecasts.
      if (i + 1 < 2 * k - 1) row.push(CellError.of('#N/A'))
      else row.push(Math.sqrt(sum(Array.from({ length: k }, (_, j) => (xs[i - j] - avg[i - j]!) ** 2)) / k))
    }
    return row
  })
}

function expSmoothing(doc: Doc, p: ToolParams): Out {
  const xs = one(doc, p.input, !!p.labels, 'input range').xs
  const damping = p.n ?? 0.3
  if (!(damping >= 0 && damping <= 1)) throw new Error('The damping factor must be between 0 and 1.')
  const f: (number | null)[] = [null]
  for (let i = 1; i < xs.length; i++) f.push(i === 1 ? xs[0] : (1 - damping) * xs[i - 1] + damping * f[i - 1]!)
  return xs.map((_, i) => {
    const row: Scalar[] = [f[i] ?? CellError.of('#N/A')]
    if (p.stdErrors) row.push(i < 4 ? CellError.of('#N/A') : Math.sqrt(sum([1, 2, 3].map((j) => (xs[i - j] - f[i - j]!) ** 2)) / 3))
    return row
  })
}

function rankPercentile(doc: Doc, p: ToolParams): Out {
  const gs = groups(doc, need(p.input, 'input range'), !!p.byRows, !!p.labels)
  const blocks = gs.map((g) => {
    const order = g.xs.map((x, i) => ({ x, i })).sort((a, b) => b.x - a.x)
    const n = g.xs.length
    return [
      ['Point', g.name, 'Rank', 'Percent'],
      ...order.map(({ x, i }) => {
        const rank = 1 + g.xs.filter((y) => y > x).length
        const below = g.xs.filter((y) => y < x).length
        // PERCENTRANK.INC to three decimals, shown as a percentage.
        const pct = n > 1 ? Math.floor((below / (n - 1)) * 1000) / 1000 : 1
        return [i + 1, x, rank, `${(pct * 100).toFixed(2)}%`]
      })
    ]
  })
  const height = Math.max(...blocks.map((b) => b.length))
  return Array.from({ length: height }, (_, r) => blocks.flatMap((b, k) => [...(b[r] ?? [undefined, undefined, undefined, undefined]), ...(k < blocks.length - 1 ? [undefined] : [])]))
}

function regression(doc: Doc, p: ToolParams): Out {
  const y = need(p.input, 'Input Y range')
  const x = need(p.extra, 'Input X range')
  const labels = !!p.labels
  const body = (a: Area): Area => (labels ? { ...a, r0: a.r0 + 1 } : a)
  const yb = body(y)
  const xb = body(x)
  const k = xb.c1 - xb.c0 + 1
  const n = yb.r1 - yb.r0 + 1
  if (y.c1 !== y.c0) throw new Error('The Y range must be a single column.')
  if (xb.r1 - xb.r0 + 1 !== n) throw new Error('The X and Y ranges must have the same number of rows.')
  const zero = !!p.zeroConstant
  const L = fxArray(doc, `LINEST(${refText(doc, yb)},${refText(doc, xb)},${zero ? 'FALSE' : 'TRUE'},TRUE)`) as number[][]
  // LINEST: row 0 coefficients and row 1 their standard errors, last first; then R², SEy; F, df; SSreg, SSresid.
  const coef = [...L[0]].reverse()
  const se = [...L[1]].reverse()
  const [r2, sey] = [L[2][0], L[2][1]]
  const [F, dfRes] = [L[3][0], L[3][1]]
  const [ssReg, ssRes] = [L[4][0], L[4][1]]
  const dfReg = k
  const level = p.confidence ?? 0.95
  const names = Array.from({ length: k }, (_, j) => {
    const v = labels ? doc.sheet(x.sheet).get(x.r0, x.c0 + j)?.v : null
    return v !== null && v !== undefined && v !== '' ? String(v) : `X Variable ${k === 1 ? '' : j + 1}`.trim()
  })
  const terms = zero ? names.map((nm, j) => ({ nm, b: coef[j + 1], s: se[j + 1] })) : [{ nm: 'Intercept', b: coef[0], s: se[0] }, ...names.map((nm, j) => ({ nm, b: coef[j + 1], s: se[j + 1] }))]
  const tc95 = fx(doc, 'T.INV.2T', 0.05, dfRes)
  const tc = fx(doc, 'T.INV.2T', 1 - level, dfRes)
  const pct = `${Math.round(level * 1000) / 10}%`
  const out: Out = [
    ['SUMMARY OUTPUT'],
    [],
    ['Regression Statistics'],
    ['Multiple R', Math.sqrt(r2)],
    ['R Square', r2],
    ['Adjusted R Square', 1 - ((1 - r2) * (n - (zero ? 0 : 1))) / dfRes],
    ['Standard Error', sey],
    ['Observations', n],
    [],
    ['ANOVA'],
    ['', 'df', 'SS', 'MS', 'F', 'Significance F'],
    ['Regression', dfReg, ssReg, ssReg / dfReg, F, fx(doc, 'F.DIST.RT', F, dfReg, dfRes)],
    ['Residual', dfRes, ssRes, ssRes / dfRes],
    ['Total', dfReg + dfRes, ssReg + ssRes],
    [],
    ['', 'Coefficients', 'Standard Error', 't Stat', 'P-value', 'Lower 95%', 'Upper 95%', `Lower ${pct}`, `Upper ${pct}`],
    ...terms.map(({ nm, b, s }) => (s ? [nm, b, s, b / s, fx(doc, 'T.DIST.2T', Math.abs(b / s), dfRes), b - tc95 * s, b + tc95 * s, b - tc * s, b + tc * s] : [nm, b, CellError.of('#N/A'), CellError.of('#N/A'), CellError.of('#N/A')]))
  ]
  if (p.residuals) {
    const sh = doc.sheet(y.sheet)
    const num = (r: number, c: number) => Number(sh.get(r, c)?.v ?? 0)
    out.push([], [], ['RESIDUAL OUTPUT'], [], ['Observation', `Predicted ${labels ? String(sh.get(y.r0, y.c0)?.v ?? 'Y') : 'Y'}`, 'Residuals', 'Standard Residuals'])
    const res = Array.from({ length: n }, (_, i) => {
      const pred = coef[0] * (zero ? 0 : 1) + sum(Array.from({ length: k }, (_, j) => coef[j + 1] * num(xb.r0 + i, xb.c0 + j)))
      return { pred, e: num(yb.r0 + i, yb.c0) - pred }
    })
    // Excel standardizes residuals by their standard deviation (n − 1).
    const sdE = Math.sqrt(sum(res.map((r) => r.e ** 2)) / (n - 1))
    res.forEach((r, i) => out.push([i + 1, r.pred, r.e, r.e / sdE]))
  }
  return out
}

function fourier(doc: Doc, p: ToolParams): Out {
  const a = need(p.input, 'input range')
  const sh = doc.sheet(a.sheet)
  const vals: { re: number; im: number }[] = []
  for (let r = a.r0 + (p.labels ? 1 : 0); r <= a.r1; r++)
    for (let c = a.c0; c <= a.c1; c++) {
      const v = sh.get(r, c)?.v
      if (v === null || v === undefined || v === '') continue
      const re = fx(doc, 'IMREAL', typeof v === 'number' ? v : `"${String(v).replace(/"/g, '""')}"`)
      const im = fx(doc, 'IMAGINARY', typeof v === 'number' ? v : `"${String(v).replace(/"/g, '""')}"`)
      if (!Number.isFinite(re) || !Number.isFinite(im)) throw new Error('The input range must contain numbers or complex numbers like 3+4i.')
      vals.push({ re, im })
    }
  const n = vals.length
  if (n < 2 || n > 4096 || (n & (n - 1)) !== 0) throw new Error(`Fourier Analysis needs 2, 4, 8, … up to 4096 values; this range has ${n}.`)
  const out = fft(vals, !!p.inverse).map((z) => (p.inverse ? { re: z.re / n, im: z.im / n } : z))
  return [...(p.labels ? [[String(sh.get(a.r0, a.c0)?.v ?? '')]] : []), ...out.map((z) => [complexText(z.re, z.im)])]
}

/** Radix-2 FFT (the caller divides the inverse by n, as Excel does). */
function fft(x: { re: number; im: number }[], inverse: boolean): { re: number; im: number }[] {
  const n = x.length
  if (n === 1) return [x[0]]
  const even = fft(x.filter((_, i) => i % 2 === 0), inverse)
  const odd = fft(x.filter((_, i) => i % 2 === 1), inverse)
  const out = new Array<{ re: number; im: number }>(n)
  const sign = inverse ? 1 : -1
  for (let k = 0; k < n / 2; k++) {
    const t = (sign * 2 * Math.PI * k) / n
    const w = { re: Math.cos(t), im: Math.sin(t) }
    const o = { re: w.re * odd[k].re - w.im * odd[k].im, im: w.re * odd[k].im + w.im * odd[k].re }
    out[k] = { re: even[k].re + o.re, im: even[k].im + o.im }
    out[k + n / 2] = { re: even[k].re - o.re, im: even[k].im - o.im }
  }
  return out
}

/** Complex text as Excel's COMPLEX writes it, with tiny rounding noise removed. */
function complexText(re: number, im: number): string {
  const tidy = (v: number) => (Math.abs(v) < 1e-12 ? 0 : Number(v.toPrecision(15)))
  const a = tidy(re)
  const b = tidy(im)
  if (b === 0) return numberToText(a)
  const bt = b === 1 ? '' : b === -1 ? '-' : numberToText(b)
  return a === 0 ? `${bt}i` : `${numberToText(a)}${b > 0 ? '+' : ''}${bt}i`
}

/** A small seeded generator (mulberry32), so a seed gives the same numbers every time. */
function rng(seed: number | undefined): () => number {
  let s = (seed ?? Math.floor(Math.random() * 2 ** 32)) >>> 0
  return () => {
    s = (s + 0x6d2b79f5) >>> 0
    let t = s
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

function random(p: ToolParams): Out {
  const vars = Math.max(1, Math.round(p.variables ?? 1))
  const count = Math.round(p.count ?? 0)
  if (count < 1 || count * vars > 1_000_000) throw new Error('Enter how many random numbers to make (up to a million in all).')
  const u = rng(p.seed)
  const a = p.p1 ?? 0
  const b = p.p2 ?? 1
  const draw = (): number => {
    switch (p.distribution ?? 'uniform') {
      case 'uniform':
        return a + (b - a) * u()
      case 'normal': {
        // Box–Muller: mean a, standard deviation b.
        const r = Math.sqrt(-2 * Math.log(1 - u()))
        return a + b * r * Math.cos(2 * Math.PI * u())
      }
      case 'bernoulli':
        return u() < a ? 1 : 0
      case 'binomial': {
        let k = 0
        for (let i = 0; i < b; i++) if (u() < a) k++
        return k
      }
      case 'poisson': {
        // Knuth's method (λ = a).
        const L = Math.exp(-a)
        let k = 0
        let q = 1
        do {
          k++
          q *= u()
        } while (q > L && k < 1e6)
        return k - 1
      }
    }
  }
  return Array.from({ length: count }, () => Array.from({ length: vars }, draw))
}

function sampling(doc: Doc, p: ToolParams): Out {
  const xs = groups(doc, need(p.input, 'input range'), false, !!p.labels).flatMap((g) => g.xs)
  const n = Math.round(p.n ?? 0)
  if (n < 1) throw new Error(p.method === 'random' ? 'Enter the number of samples.' : 'Enter the period.')
  if (p.method === 'random') {
    const u = rng(p.seed)
    return Array.from({ length: n }, () => [xs[Math.floor(u() * xs.length)]])
  }
  if (n > xs.length) throw new Error(`The period must be at most the number of values (${xs.length}).`)
  const out: Out = []
  for (let i = n - 1; i < xs.length; i += n) out.push([xs[i]])
  return out
}

function need(a: Area | undefined, what: string): Area {
  if (!a) throw new Error(`Enter the ${what}.`)
  return a
}

const TOOLS: Record<ToolId, (doc: Doc, p: ToolParams) => Out> = {
  anova1,
  anova2r,
  anova2,
  correlation: (d, p) => matrix(d, p, correl),
  covariance: (d, p) => matrix(d, p, covar),
  expsmooth: expSmoothing,
  ftest,
  fourier,
  histogram,
  movavg: movingAverage,
  random: (_d, p) => random(p),
  rank: rankPercentile,
  regression,
  sampling,
  'ttest-paired': (d, p) => ttest(d, p, 'paired'),
  'ttest-equal': (d, p) => ttest(d, p, 'equal'),
  'ttest-unequal': (d, p) => ttest(d, p, 'unequal'),
  ztest
}

/** The tool's output table (exported for tests). */
export function toolOutput(doc: Doc, id: ToolId, p: ToolParams): Out {
  return TOOLS[id](doc, p)
}

/** Run a tool and write its output (one undoable step). Returns the output's sheet. */
export function runTool(doc: Doc, id: ToolId, p: ToolParams): { sheet: number } {
  const out = toolOutput(doc, id, p)
  let sheet = 0
  doc.transact('Data analysis', (tx: Tx) => {
    let at: { sheet: number; r: number; c: number }
    if ('newSheet' in p.out) {
      const name = p.out.newSheet.trim() || 'Analysis'
      let unique = name
      for (let k = 2; doc.wb.sheetByName(unique); k++) unique = `${name} ${k}`
      const sh = doc.wb.createSheet(unique)
      const idx = doc.wb.sheets.length
      tx.custom(
        () => doc.wb.sheets.splice(idx, 0, sh),
        () => doc.wb.sheets.splice(doc.wb.sheetIndex(sh.id), 1)
      )
      tx.structural = true
      at = { sheet: sh.id, r: 0, c: 0 }
    } else at = p.out
    sheet = at.sheet
    out.forEach((row, i) =>
      row.forEach((v, j) => {
        if (v === undefined) return
        // Numbers at full precision (String gives the shortest text that reads back as the same double).
        const text = v === null ? '' : typeof v === 'number' ? (Number.isFinite(v) ? String(v).toUpperCase() : '#NUM!') : v instanceof CellError ? v.code : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : /^-?[\d.]+%$/.test(v) ? v : "'" + v
        doc.inputTx(tx, at.sheet, at.r + i, at.c + j, text)
      })
    )
  })
  return { sheet }
}
