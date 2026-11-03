import { describe, expect, it } from 'vitest'
import { Host } from '@/core/host'

function sales() {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  ;[['Month', 'North', 'South'], ['Jan', '10', '7'], ['Feb', '12', 'n/a'], ['Mar', '15', '9']].forEach((row, r) => row.forEach((v, c) => h.doc.input(s, r, c, v)))
  return { h, s }
}

describe('charts', () => {
  it('a chart of a range: series by column, names from the header, categories from the first column', () => {
    const { h, s } = sales()
    const r = h.addChartFromRange(s, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 2 })
    expect(r.ok).toBe(true)
    const chart = h.doc.wb.sheets[0].charts[0]
    expect(chart.series).toEqual([
      { name: 'Sheet1!$B$1', cat: 'Sheet1!$A$2:$A$4', val: 'Sheet1!$B$2:$B$4' },
      { name: 'Sheet1!$C$1', cat: 'Sheet1!$A$2:$A$4', val: 'Sheet1!$C$2:$C$4' }
    ])
    expect(chart.anchor).toMatchObject({ r0: 0, c0: 4 })
    const data = h.chartData(s, chart.id)!
    expect(data.categories).toEqual(['Jan', 'Feb', 'Mar'])
    expect(data.series).toEqual([
      { name: 'North', values: [10, 12, 15] },
      { name: 'South', values: [7, null, 9] } // text is a gap
    ])
    // The data follows edits; hidden rows aren't plotted.
    h.doc.input(s, 1, 1, '11')
    expect(h.chartData(s, chart.id)!.series[0].values[0]).toBe(11)
    h.doc.hide(s, 'row', [2], true)
    expect(h.chartData(s, chart.id)!.categories).toEqual(['Jan', 'Mar'])
  })

  it('references and position follow inserted rows and columns; edit, delete, undo', () => {
    const { h, s } = sales()
    const { id } = h.addChartFromRange(s, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 1 }) as { id: number }
    h.doc.insertRows(s, 0, 2)
    let chart = h.doc.wb.sheets[0].charts[0]
    expect(chart.series[0].val).toBe('Sheet1!$B$4:$B$6')
    expect(chart.anchor.r0).toBe(2)
    h.doc.insertCols(s, 0, 1)
    chart = h.doc.wb.sheets[0].charts[0]
    expect(chart.series[0].cat).toBe('Sheet1!$B$4:$B$6')
    expect(chart.anchor.c0).toBe(4)
    expect(h.updateChart(s, id, { type: 'line', title: 'Trend', legend: 'bottom' }).ok).toBe(true)
    expect(h.doc.wb.sheets[0].charts[0]).toMatchObject({ type: 'line', title: 'Trend', legend: 'bottom' })
    expect(h.updateChart(s, id, { series: [{ val: 'Sheet1!$C$4:' }] }).ok).toBe(false) // a broken reference is refused
    h.deleteChart(s, id)
    expect(h.doc.wb.sheets[0].charts).toEqual([])
    h.doc.undo()
    expect(h.doc.wb.sheets[0].charts.length).toBe(1)
    expect(h.addChartFromRange(s, { sheet: s, r0: 20, c0: 20, r1: 20, c1: 20 }).ok).toBe(false)
  })

  it('rows of series when the range is wider than tall; scatter uses x values', () => {
    const h = new Host()
    const s = h.doc.wb.sheets[0].id
    ;[['', 'Q1', 'Q2', 'Q3'], ['Ann', '1', '2', '3']].forEach((row, r) => row.forEach((v, c) => v && h.doc.input(s, r, c, v)))
    h.addChartFromRange(s, { sheet: s, r0: 0, c0: 0, r1: 1, c1: 3 }, 'scatter')
    const chart = h.doc.wb.sheets[0].charts[0]
    expect(chart.series).toEqual([{ name: 'Sheet1!$A$2', cat: 'Sheet1!$B$1:$D$1', val: 'Sheet1!$B$2:$D$2' }])
    const d = h.chartData(s, chart.id)!
    expect(d.series[0]).toMatchObject({ name: 'Ann', values: [1, 2, 3], x: [null, null, null] })
  })
})

describe('chart drawing', () => {
  it('the SVG used for printing is well-formed XML for every chart type, with text from cells escaped', async () => {
    const { chartSvg } = await import('@/renderer/app/Charts')
    const { DOMParser } = await import('@xmldom/xmldom')
    for (const type of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'] as const) {
      const svg = chartSvg({ id: 1, type, grouping: 'clustered', title: 'A <b>&</b> "q"', series: [{ val: 'x' }], legend: 'right', anchor: { r0: 0, c0: 0, dx0: 0, dy0: 0, r1: 1, c1: 1, dx1: 0, dy1: 0 } }, { categories: ['<x>', 'b & c'], series: [{ name: 'N "1"', values: [1, 2], x: [1, 2] }] }, 300, 200)
      const errors: string[] = []
      new DOMParser({ onError: (_l, m) => errors.push(m) }).parseFromString(svg, 'image/svg+xml')
      expect(errors, type).toEqual([])
    }
  })
})
