// Performance benchmarks (opt-in: npm run bench). Results are printed as a table for docs/PERFORMANCE_BENCHMARKS.md.
import { describe, expect, it } from 'vitest'
import { Host } from '@/core/host'
import { Doc } from '@/core/doc'

const results: [string, string, string][] = []
const time = async <T>(f: () => T | Promise<T>): Promise<[number, T]> => {
  const t = performance.now()
  const r = await f()
  return [performance.now() - t, r]
}
const ms = (x: number) => `${x < 10 ? x.toFixed(2) : Math.round(x)} ms`

describe('benchmarks', () => {
  it('open and save a 100,000-row XLSX', async () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0]
    for (let r = 0; r < 100_000; r++) {
      s.set(r, 0, { v: r })
      s.set(r, 1, { v: `Item ${r % 977}` })
      s.set(r, 2, { v: 45000 + (r % 365) })
      s.set(r, 3, { v: (r * 7919) % 1000 })
    }
    doc.engine.rebuild()
    const [tSave, out] = await time(() => new Host(doc).save('xlsx'))
    const [tOpen, h] = await time(() => Host.open(out.data, 'xlsx'))
    expect(h.doc.value(h.doc.wb.sheets[0].id, 99_999, 0)).toBe(99_999)
    results.push(['Open a 100,000-row × 4-column XLSX', ms(tOpen), '< 3,000 ms'])
    results.push(['Save it', ms(tSave), '—'])
  }, 120_000)

  it('recalculate a 100,000-cell dependency chain', () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    doc.input(s, 0, 0, '1')
    const sh = doc.wb.sheets[0]
    const ast = doc.parse('=A1+1', s, 1, 0)
    for (let r = 1; r < 100_000; r++) sh.set(r, 0, { v: 0, f: { ast } })
    doc.engine.rebuild()
    const t0 = performance.now()
    doc.engine.recalcAll()
    const tFull = performance.now() - t0
    expect(doc.value(s, 99_999, 0)).toBe(100_000)
    const t1 = performance.now()
    doc.input(s, 0, 0, '2')
    const tEdit = performance.now() - t1
    expect(doc.value(s, 99_999, 0)).toBe(100_001)
    results.push(['Full recalculation of a 100,000-cell chain', ms(tFull), '< 500 ms'])
    results.push(['Edit the head of that chain (100,000 dependents recalculate)', ms(tEdit), '< 500 ms'])
  }, 120_000)

  it('edit a cell in a 100,000-row sheet with totals: core time to the screen patch', () => {
    const h = new Host()
    const s = h.doc.wb.sheets[0]
    for (let r = 0; r < 100_000; r++) s.set(r, 0, { v: r })
    h.doc.engine.rebuild()
    h.doc.input(s.id, 0, 2, '=SUM(A:A)')
    h.doc.input(s.id, 1, 2, '=AVERAGE(A1:A100000)')
    h.changes()
    const samples: number[] = []
    for (let i = 0; i < 50; i++) {
      const t = performance.now()
      h.input(s.id, 500 + i, 0, String(i))
      h.changes()
      samples.push(performance.now() - t)
    }
    samples.sort((a, b) => a - b)
    results.push(['Type into a cell that two whole-column totals depend on (median of 50)', ms(samples[25]), '< 50 ms'])
  }, 120_000)

  it('memory for 1,000,000 numeric cells', () => {
    globalThis.gc?.()
    const before = process.memoryUsage().heapUsed
    const doc = new Doc()
    const s = doc.wb.sheets[0]
    for (let r = 0; r < 100_000; r++) for (let c = 0; c < 10; c++) s.set(r, c, { v: r * 10 + c })
    globalThis.gc?.()
    const after = process.memoryUsage().heapUsed
    expect(s.get(99_999, 9)?.v).toBe(999_999)
    results.push(['Heap for 1,000,000 numeric cells (100,000 × 10)', `${Math.round((after - before) / 1024 / 1024)} MB`, '< 500 MB'])
  }, 120_000)

  it('report', async () => {
    const { writeFileSync, mkdirSync } = await import('node:fs')
    const table = '| Measure | Result | Target |\n|---|---|---|\n' + results.map((r) => `| ${r.join(' | ')} |`).join('\n') + '\n'
    mkdirSync('test-results', { recursive: true })
    writeFileSync('test-results/bench.md', table)
    console.log(table)
  })
})
