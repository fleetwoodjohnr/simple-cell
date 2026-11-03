import { describe, expect, it } from 'vitest'
import { pivot, emptyPivot, type PivotConfig } from '@/core/pivot'
import { dateSerial } from '@/core/numfmt/dates'

const header = ['Region', 'Product', 'Sales', 'Units', 'Date']
const d = (y: number, m: number, day: number) => dateSerial(y, m, day, false)
const recs = [
  ['East', 'Pens', 10, 1, d(2024, 1, 5)],
  ['East', 'Ink', 5, 2, d(2024, 4, 9)],
  ['West', 'Pens', 20, 3, d(2024, 2, 1)],
  ['West', 'Pens', 7, 1, d(2025, 1, 2)],
  ['North', 'Ink', 3, 4, d(2025, 6, 30)]
]
const grid = (cfg: Partial<PivotConfig>) => pivot(header, recs, { ...emptyPivot(), ...cfg }).cells.map((r) => r.map((c) => c.v))

describe('pivot engine', () => {
  it('rows and a sum, with a grand total', () => {
    expect(grid({ rows: ['Region'], values: [{ field: 'Sales', agg: 'sum' }] })).toEqual([
      ['Region', 'Sum of Sales'],
      ['East', 15],
      ['North', 3],
      ['West', 27],
      ['Grand Total', 45]
    ])
  })

  it('rows × columns with grand totals both ways; empty intersections stay empty', () => {
    expect(grid({ rows: ['Region'], cols: ['Product'], values: [{ field: 'Sales', agg: 'sum' }] })).toEqual([
      ['Sum of Sales', 'Product', '', ''],
      ['Region', 'Ink', 'Pens', 'Grand Total'],
      ['East', 5, 10, 15],
      ['North', 3, null, 3],
      ['West', null, 27, 27],
      ['Grand Total', 8, 37, 45]
    ])
  })

  it('two row fields with subtotals; several values across; counts and averages', () => {
    expect(grid({ rows: ['Region', 'Product'], values: [{ field: 'Sales', agg: 'sum' }, { field: 'Units', agg: 'average' }] })).toEqual([
      ['Region', 'Product', 'Sum of Sales', 'Average of Units'],
      ['East', 'Ink', 5, 2],
      ['', 'Pens', 10, 1],
      ['East Total', '', 15, 1.5],
      ['North', 'Ink', 3, 4],
      ['North Total', '', 3, 4],
      ['West', 'Pens', 27, 2],
      ['West Total', '', 27, 2],
      ['Grand Total', '', 45, 2.2]
    ])
    expect(grid({ rows: ['Product'], values: [{ field: 'Region', agg: 'count' }] })).toEqual([
      ['Product', 'Count of Region'],
      ['Ink', 2],
      ['Pens', 3],
      ['Grand Total', 5]
    ])
  })

  it('percent of grand total; filters; date grouping by year', () => {
    const p = pivot(header, recs, { ...emptyPivot(), rows: ['Region'], values: [{ field: 'Sales', agg: 'sum', showAs: 'percentOfTotal' }] })
    expect(p.cells[1][1]).toEqual({ v: 15 / 45, kind: 'value', fmt: '0.00%' })
    expect(p.cells[4][1].v).toBe(1)
    expect(grid({ rows: ['Product'], filters: ['Region'], include: { Region: ['West'] }, values: [{ field: 'Sales', agg: 'sum' }] })).toEqual([['Region', 'West'], [], ['Product', 'Sum of Sales'], ['Pens', 27], ['Grand Total', 27]])
    expect(grid({ rows: ['Date'], dateGroups: { Date: 'years' }, values: [{ field: 'Sales', agg: 'sum' }] })).toEqual([['Date', 'Sum of Sales'], ['2024', 35], ['2025', 10], ['Grand Total', 45]])
    expect(pivot(header, recs, emptyPivot()).items.Region).toEqual(['East', 'North', 'West'])
  })
})

describe('PivotTables in a workbook', async () => {
  const { Host } = await import('@/core/host')
  function data() {
    const h = new Host()
    const s = h.doc.wb.sheets[0].id
    ;[['Region', 'Product', 'Sales'], ['East', 'Pens', '$10.00'], ['West', 'Ink', '$5.50'], ['East', 'Ink', '$2.00']].forEach((row, r) => row.forEach((v, c) => h.doc.input(s, r, c, v)))
    return { h, s }
  }
  it('created on a new sheet; layout writes formatted cells; refresh; undo', () => {
    const { h, s } = data()
    const r = h.createPivot(s, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 2 }, null)
    expect(r.ok).toBe(true)
    const { sheet: ps, id } = r as { sheet: number; id: number }
    expect(h.doc.wb.sheets.map((x) => x.name)).toEqual(['Sheet2', 'Sheet1'])
    h.setPivotConfig(ps, id, { ...emptyPivot(), rows: ['Region'], values: [{ field: 'Sales', agg: 'sum' }] })
    const v = (r: number, c: number) => h.doc.value(ps, r, c)
    expect([v(2, 0), v(2, 1), v(3, 0), v(3, 1), v(5, 0), v(5, 1)]).toEqual(['Region', 'Sum of Sales', 'East', 12, 'Grand Total', 17.5])
    expect(h.doc.display(ps, 3, 1)?.t).toBe('$12.00 ') // the source's number format (with its _) padding)
    expect(h.doc.styleAt(ps, 2, 0).bold).toBe(true)
    expect(h.pivotAt(ps, 3, 1)).toMatchObject({ id, fields: ['Region', 'Product', 'Sales'], numeric: [false, false, true] })
    // The source changes: refresh picks it up.
    h.doc.input(s, 1, 2, '20')
    expect(v(3, 1)).toBe(12)
    h.refreshPivots()
    expect(v(3, 1)).toBe(22)
    // A smaller layout clears what the larger one wrote.
    h.setPivotConfig(ps, id, { ...emptyPivot(), values: [{ field: 'Sales', agg: 'count' }] })
    expect(v(5, 0)).toBeNull()
    h.doc.undo()
    expect(v(5, 0)).toBe('Grand Total')
  })

  it('refuses to overwrite data; deleting clears its cells; the source follows inserted rows', () => {
    const { h, s } = data()
    h.doc.input(s, 5, 6, 'keep')
    const r = h.createPivot(s, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 2 }, { sheet: s, r: 2, c: 5 }) as { sheet: number; id: number }
    const bad = h.setPivotConfig(s, r.id, { ...emptyPivot(), rows: ['Region', 'Product'], values: [{ field: 'Sales', agg: 'sum' }] })
    expect(bad.ok).toBe(false)
    expect(h.doc.value(s, 5, 6)).toBe('keep')
    h.doc.input(s, 5, 6, '')
    h.setPivotConfig(s, r.id, { ...emptyPivot(), rows: ['Product'], values: [{ field: 'Sales', agg: 'sum' }] })
    h.doc.insertRows(s, 2, 1)
    const p = h.doc.wb.sheets[0].pivots[0]
    expect(p.source).toBe('Sheet1!$A$1:$C$5')
    expect(p.at).toEqual({ r: 3, c: 5 })
    h.deletePivot(s, r.id)
    expect(h.doc.value(s, 3, 5)).toBeNull()
    expect(h.doc.wb.sheets[0].pivots).toEqual([])
    expect(h.createPivot(s, { sheet: s, r0: 10, c0: 10, r1: 12, c1: 11 }, null).ok).toBe(false) // no headings
  })
})
