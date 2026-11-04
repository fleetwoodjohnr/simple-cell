import { describe, expect, it } from 'vitest'
import { Doc } from '@/core/doc'
import { runTool, toolOutput, type ToolId, type ToolParams } from '@/core/toolpak'
import type { Scalar } from '@/core/engine/values'

/** A sheet with columns of numbers from A1 (headings in row 1), and a helper to evaluate worksheet formulas. */
function book(cols: Record<string, number[]>) {
  const doc = new Doc()
  const s = doc.wb.sheets[0].id
  Object.entries(cols).forEach(([name, xs], c) => {
    doc.input(s, 0, c, name)
    xs.forEach((x, r) => doc.input(s, r + 1, c, String(x)))
  })
  const f = (formula: string): Scalar => {
    doc.input(s, 900, 0, formula)
    return doc.value(s, 900, 0)
  }
  const area = (c0: number, c1: number, rows: number) => ({ sheet: s, r0: 0, c0, r1: rows, c1 })
  return { doc, s, f, area }
}

/** The last row with this label (ANOVA tables come after the summaries, which can use the same words). */
const row = (out: (Scalar | undefined)[][], label: string) => out.findLast((r) => r[0] === label)!
const run = (doc: Doc, id: ToolId, p: Omit<ToolParams, 'out'>) => toolOutput(doc, id, { ...p, out: { newSheet: 'x' } })

const A = [6, 8, 4, 5, 3, 4]
const B = [8, 12, 9, 11, 6, 8]
const C = [13, 9, 11, 8, 7, 12]

describe('Analysis ToolPak', () => {
  it('Anova: Single Factor — sums of squares add up; F and its p-value agree with F.DIST.RT', () => {
    const { doc, f, area } = book({ A, B, C })
    const out = run(doc, 'anova1', { input: area(0, 2, 6), labels: true, alpha: 0.05 })
    const between = row(out, 'Between Groups') as number[]
    const within = row(out, 'Within Groups') as number[]
    const total = row(out, 'Total') as number[]
    expect(between[1] + within[1]).toBeCloseTo(total[1], 10)
    expect(total[1]).toBeCloseTo(f('=DEVSQ(A2:C7)') as number, 10)
    expect([between[2], within[2], total[2]]).toEqual([2, 15, 17])
    expect(between[4]).toBeCloseTo(between[3] / within[3], 12)
    expect(between[5]).toBeCloseTo(f(`=F.DIST.RT(${between[4]},2,15)`) as number, 12)
    expect(between[6]).toBeCloseTo(f('=F.INV.RT(0.05,2,15)') as number, 12)
    expect(row(out, 'A')).toEqual(['A', 6, 30, 5, f('=VAR.S(A2:A7)')])
  })

  it('Anova: Two-Factor without and with replication — the parts add up to the total', () => {
    const { doc, f, area } = book({ A, B, C })
    // No headings: rows and columns are named Row 1…, Column 1….
    const two = run(doc, 'anova2', { input: { ...area(0, 2, 6), r0: 1 }, labels: false })
    expect(two.map((x) => x[0])).toContain('Row 6')
    const ss = (l: string) => (row(two, l) as number[])[1]
    expect(ss('Rows') + ss('Columns') + ss('Error')).toBeCloseTo(ss('Total'), 10)
    expect(ss('Total')).toBeCloseTo(f('=DEVSQ(A2:C7)') as number, 10)
    // With replication: labels in column A of the range, 3 rows per sample (2 samples).
    const r = book({ S: [0, 0, 0, 0, 0, 0], A, B, C })
    r.doc.input(r.s, 1, 0, 'Low')
    r.doc.input(r.s, 4, 0, 'High')
    const rep = run(r.doc, 'anova2r', { input: r.area(0, 3, 6), n: 3 })
    const s2 = (l: string) => (row(rep, l) as number[])[1]
    expect(s2('Sample') + s2('Columns') + s2('Interaction') + s2('Within')).toBeCloseTo(s2('Total'), 10)
    expect(s2('Total')).toBeCloseTo(f('=DEVSQ(A2:C7)') as number, 10)
    expect((row(rep, 'Within') as number[])[2]).toBe(12)
    expect(rep.map((x) => x[0])).toContain('High')
  })

  it('t-tests: paired and equal-variance p-values equal T.TEST; unequal rounds df like Excel; z-test', () => {
    const { doc, f, area } = book({ A, B })
    const p = { input: area(0, 0, 6), input2: area(1, 1, 6), labels: true }
    expect((row(run(doc, 'ttest-paired', p), 'P(T<=t) two-tail') as number[])[1]).toBeCloseTo(f('=T.TEST(A2:A7,B2:B7,2,1)') as number, 12)
    expect((row(run(doc, 'ttest-equal', p), 'P(T<=t) two-tail') as number[])[1]).toBeCloseTo(f('=T.TEST(A2:A7,B2:B7,2,2)') as number, 12)
    const u = run(doc, 'ttest-unequal', p)
    expect(Number.isInteger((row(u, 'df') as number[])[1])).toBe(true)
    expect((row(u, 't Stat') as number[])[1]).toBeCloseTo((5 - 9) / Math.sqrt((f('=VAR.S(A2:A7)') as number) / 6 + (f('=VAR.S(B2:B7)') as number) / 6), 12)
    const z = run(doc, 'ztest', { ...p, var1: 3, var2: 4 })
    expect((row(z, 'z') as number[])[1]).toBeCloseTo(-4 / Math.sqrt(3 / 6 + 4 / 6), 12)
    expect((row(z, 'P(Z<=z) two-tail') as number[])[1]).toBeCloseTo(2 * (1 - (f(`=NORM.S.DIST(${4 / Math.sqrt(3 / 6 + 4 / 6)},TRUE)`) as number)), 12)
  })

  it('F-test: one tail is half of F.TEST', () => {
    const { doc, f, area } = book({ A, B })
    const out = run(doc, 'ftest', { input: area(0, 0, 6), input2: area(1, 1, 6), labels: true })
    expect((row(out, 'P(F<=f) one-tail') as number[])[1] * 2).toBeCloseTo(f('=F.TEST(A2:A7,B2:B7)') as number, 12)
  })

  it('correlation and covariance matrices equal CORREL and COVARIANCE.P', () => {
    const { doc, f, area } = book({ A, B, C })
    const cor = run(doc, 'correlation', { input: area(0, 2, 6), labels: true })
    expect(cor[0]).toEqual(['', 'A', 'B', 'C'])
    expect(cor[3][2]).toBeCloseTo(f('=CORREL(B2:B7,C2:C7)') as number, 12)
    expect(cor[1][1]).toBeCloseTo(1, 12)
    expect(cor[1][2]).toBeUndefined()
    const cov = run(doc, 'covariance', { input: area(0, 2, 6), labels: true })
    expect(cov[3][1]).toBeCloseTo(f('=COVARIANCE.P(A2:A7,C2:C7)') as number, 12)
  })

  it('regression: coefficients, statistics and p-values agree with LINEST, RSQ and T.DIST.2T; residuals', () => {
    const { doc, f, area } = book({ Y: [3, 5, 7, 10, 11, 15, 16], X1: [1, 2, 3, 4, 5, 6, 7], X2: [2, 1, 4, 3, 6, 5, 8] })
    const out = run(doc, 'regression', { input: area(0, 0, 7), extra: area(1, 2, 7), labels: true, residuals: true, confidence: 0.9 })
    expect((row(out, 'R Square') as number[])[1]).toBeCloseTo(f('=INDEX(LINEST(A2:A8,B2:C8,TRUE,TRUE),3,1)') as number, 12)
    const x1 = row(out, 'X1') as number[]
    expect(x1[1]).toBeCloseTo(f('=INDEX(LINEST(A2:A8,B2:C8),1,2)') as number, 12)
    expect(x1[4]).toBeCloseTo(f(`=T.DIST.2T(ABS(${x1[3]}),4)`) as number, 12)
    expect(out.find((r) => r[1] === 'Coefficients')![7]).toBe('Lower 90%')
    const res = out.slice(out.findIndex((r) => r[0] === 'Observation') + 1)
    expect(res).toHaveLength(7)
    expect(res.reduce((s, r) => s + (r[2] as number), 0)).toBeCloseTo(0, 9)
    // One variable: matches SLOPE/INTERCEPT/STEYX.
    const one = run(doc, 'regression', { input: area(0, 0, 7), extra: area(1, 1, 7), labels: true })
    expect((row(one, 'Intercept') as number[])[1]).toBeCloseTo(f('=INTERCEPT(A2:A8,B2:B8)') as number, 12)
    expect((row(one, 'Standard Error') as number[])[1]).toBeCloseTo(f('=STEYX(A2:A8,B2:B8)') as number, 12)
  })

  it('histogram equals FREQUENCY; Pareto and cumulative %', () => {
    const { doc, f, area } = book({ X: [1, 2, 2, 3, 5, 8, 9, 9, 10], Bin: [2, 5, 9] })
    const out = run(doc, 'histogram', { input: area(0, 0, 9), extra: { ...area(1, 1, 3) }, labels: true, cumulative: true, pareto: true })
    expect(out.slice(1).map((r) => r[1])).toEqual([0, 1, 2, 3].map((i) => f(`=INDEX(FREQUENCY(A2:A10,B2:B4),${i + 1})`)))
    // Pareto: sorted by frequency, ties in their original order.
    expect(out[1].slice(4)).toEqual([2, 3, '33.33%'])
    expect(out[2].slice(4)).toEqual([9, 3, '66.67%'])
    expect(out[4]).toEqual(['More', 1, '100.00%', undefined, 'More', 1, '100.00%'])
  })

  it('moving average and exponential smoothing with standard errors', () => {
    const { doc, area } = book({ X: [2, 4, 6, 8, 10, 12] })
    const ma = run(doc, 'movavg', { input: area(0, 0, 6), labels: true, n: 3, stdErrors: true })
    expect(ma.map((r) => (typeof r[0] === 'number' ? r[0] : 'NA'))).toEqual(['NA', 'NA', 4, 6, 8, 10])
    expect(ma[4][1]).toBeCloseTo(2, 12)
    const es = run(doc, 'expsmooth', { input: area(0, 0, 6), labels: true, n: 0.5 })
    expect(es.map((r) => (typeof r[0] === 'number' ? r[0] : 'NA'))).toEqual(['NA', 2, 3, 4.5, 6.25, 8.125])
  })

  it('rank and percentile equal RANK.EQ and PERCENTRANK.INC', () => {
    const { doc, f, area } = book({ X: [7, 3, 9, 3, 5] })
    const out = run(doc, 'rank', { input: area(0, 0, 5), labels: true })
    expect(out[0]).toEqual(['Point', 'X', 'Rank', 'Percent'])
    expect(out[1]).toEqual([3, 9, 1, '100.00%'])
    const three = out.filter((r) => r[1] === 3)
    expect(three.map((r) => r[2])).toEqual([f('=RANK.EQ(3,A2:A6)'), f('=RANK.EQ(3,A2:A6)')])
    expect(three[0][3]).toBe(`${((f('=PERCENTRANK.INC(A2:A6,3)') as number) * 100).toFixed(2)}%`)
  })

  it('Fourier analysis: the FFT of 1, 2, 3, 4 and back', () => {
    const { doc, area } = book({ X: [1, 2, 3, 4] })
    expect(run(doc, 'fourier', { input: area(0, 0, 4), labels: true }).slice(1).map((r) => r[0])).toEqual(['10', '-2+2i', '-2', '-2-2i'])
    const back = book({ X: [10, 0, -2, 0] })
    back.doc.input(back.s, 2, 0, "'-2+2i")
    back.doc.input(back.s, 4, 0, "'-2-2i")
    expect(run(back.doc, 'fourier', { input: back.area(0, 0, 4), labels: true, inverse: true }).slice(1).map((r) => r[0])).toEqual(['1', '2', '3', '4'])
    expect(() => run(doc, 'fourier', { input: area(0, 0, 3), labels: true })).toThrow(/2, 4, 8/)
  })

  it('random numbers (seeded, so repeatable) and sampling', () => {
    const doc = new Doc()
    const norm = toolOutput(doc, 'random', { variables: 2, count: 4000, distribution: 'normal', p1: 10, p2: 2, seed: 42, out: { newSheet: 'r' } })
    const xs = norm.map((r) => r[0] as number)
    expect(xs.reduce((a, b) => a + b, 0) / xs.length).toBeCloseTo(10, 0)
    expect(toolOutput(doc, 'random', { count: 5, seed: 7, out: { newSheet: 'r' } })).toEqual(toolOutput(doc, 'random', { count: 5, seed: 7, out: { newSheet: 'r' } }))
    const uni = toolOutput(doc, 'random', { count: 500, distribution: 'uniform', p1: 5, p2: 6, seed: 1, out: { newSheet: 'r' } }).map((r) => r[0] as number)
    expect(Math.min(...uni)).toBeGreaterThanOrEqual(5)
    expect(Math.max(...uni)).toBeLessThan(6)
    const { doc: d2, area } = book({ X: [1, 2, 3, 4, 5, 6, 7] })
    expect(run(d2, 'sampling', { input: area(0, 0, 7), labels: true, method: 'periodic', n: 3 }).map((r) => r[0])).toEqual([3, 6])
  })

  it('writes the output as values on a new sheet (one undoable step) or at a cell', () => {
    const { doc, s, area } = book({ A, B })
    const r = runTool(doc, 'correlation', { input: area(0, 1, 6), labels: true, out: { newSheet: 'Correlation' } })
    const out = doc.wb.sheetById(r.sheet)!
    expect(out.name).toBe('Correlation')
    expect(out.get(0, 1)?.v).toBe('A')
    expect(typeof out.get(2, 1)?.v).toBe('number')
    doc.undo()
    expect(doc.wb.sheetByName('Correlation')).toBeUndefined()
    runTool(doc, 'ttest-paired', { input: area(0, 0, 6), input2: area(1, 1, 6), labels: true, out: { sheet: s, r: 0, c: 5 } })
    expect(doc.value(s, 0, 5)).toBe('t-Test: Paired Two Sample for Means')
  })

  it('refuses text in the data unless it is a heading', () => {
    const { doc, area } = book({ A, B })
    expect(() => run(doc, 'anova1', { input: area(0, 1, 6), labels: false })).toThrow(/tick Labels/)
  })
})
