import { describe, expect, it } from 'vitest'
import { Host } from '@/core/host'
import { parseArea } from '@/core/refs'

function sales() {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  const rows = [
    ['Region', 'Item', 'Units'],
    ['East', 'Pens', '10'],
    ['West', 'Ink', '5'],
    ['East', 'Paper', '25'],
    ['North', 'Pens', '7'],
    ['', 'Clips', '3']
  ]
  rows.forEach((row, r) => row.forEach((v, c) => v && h.doc.input(s, r, c, v)))
  const hidden = () => [...h.doc.wb.sheets[0].rowInfo].filter(([, i]) => i.hidden).map(([r]) => r).sort((a, b) => a - b)
  return { h, s, hidden }
}

describe('tables', () => {
  it('Ctrl+T: named table, structured references, header renames, total row', () => {
    const { h, s } = sales()
    expect(h.createTable(s, parseArea('A1:C6', s)!, true).ok).toBe(true)
    const t = h.doc.wb.tables[0]
    expect([t.name, t.columns.map((c) => c.name)]).toEqual(['Table1', ['Region', 'Item', 'Units']])
    h.doc.input(s, 0, 4, '=SUM(Table1[Units])')
    expect(h.doc.value(s, 0, 4)).toBe(50)
    h.input(s, 0, 2, 'Qty')
    expect(t.columns[2].name).toBe('Qty')
    h.doc.input(s, 1, 4, '=SUM(Table1[Qty])')
    expect(h.doc.value(s, 1, 4)).toBe(50)
    // Total row: label and SUBTOTAL(109) in the last column.
    h.setTableOptions('Table1', { totalsRow: true })
    expect(t.area.r1).toBe(6)
    expect(h.doc.value(s, 6, 0)).toBe('Total')
    expect(h.doc.formulaText(s, 6, 2)).toBe('=SUBTOTAL(109,Table1[Qty])')
    expect(h.doc.value(s, 6, 2)).toBe(50)
    h.doc.undo()
    expect(t.area.r1).toBe(5)
    expect(h.doc.value(s, 6, 0)).toBeNull()
  })

  it('table style is drawn beneath the cells’ own formatting', () => {
    const { h, s } = sales()
    h.createTable(s, parseArea('A1:C6', s)!, true)
    const cells = h.viewport(s, 0, 6, 0, 3)
    const at = (r: number, c: number) => cells.find((x) => x.r === r && x.c === c)
    expect(at(0, 0)?.x).toMatchObject({ bold: true, fill: '#4472C4', color: '#FFFFFF' })
    expect(at(1, 0)?.x?.fill).toBe('#DAE3F3') // first data row: stripe
    expect(at(2, 0)?.x?.fill).toBeUndefined()
    expect(at(5, 0)?.x?.borderBottom).toBeTruthy() // the empty cell A6 is still in the table
    h.doc.format([{ sheet: s, r0: 1, c0: 0, r1: 1, c1: 0 }], { fill: '#FF0000' })
    expect(h.viewport(s, 1, 1, 0, 0)[0].x?.fill).toBe('#FF0000')
    h.setTableOptions('Table1', { style: 'TableStyleLight9' })
    expect(h.viewport(s, 0, 0, 0, 0)[0].x?.fill).toBe('#4472C4')
  })

  it('without headers a header row is inserted; overlapping tables are refused', () => {
    const { h, s } = sales()
    expect(h.createTable(s, parseArea('A2:C6', s)!, false).ok).toBe(true)
    expect(h.doc.wb.tables[0].columns.map((c) => c.name)).toEqual(['Column1', 'Column2', 'Column3'])
    expect(h.doc.value(s, 1, 0)).toBe('Column1')
    expect(h.doc.value(s, 2, 0)).toBe('East')
    expect(h.createTable(s, parseArea('B3:D4', s)!, true).ok).toBe(false)
  })

  it('convert to range keeps the look and turns structured references into ranges', () => {
    const { h, s } = sales()
    h.createTable(s, parseArea('A1:C6', s)!, true)
    h.doc.input(s, 0, 4, '=SUM(Table1[Units])')
    h.convertTableToRange('Table1')
    expect(h.doc.wb.tables).toEqual([])
    expect(h.doc.formulaText(s, 0, 4)).toBe('=SUM($C$2:$C$6)')
    expect(h.doc.value(s, 0, 4)).toBe(50)
    expect(h.doc.styleAt(s, 0, 0)).toMatchObject({ bold: true, fill: '#4472C4' })
  })
})

describe('filters', () => {
  it('value checklist, blanks, custom conditions, top 10; SUBTOTAL ignores filtered rows', () => {
    const { h, s, hidden } = sales()
    h.toggleAutoFilter(s, parseArea('A1:C6', s)!)
    expect(h.filterValues(s, 'sheet', 0)).toEqual({ values: ['East', 'North', 'West'], blanks: true, current: null })
    h.setFilter(s, 'sheet', 0, { values: ['east'] })
    expect(hidden()).toEqual([2, 4, 5])
    h.doc.input(s, 7, 2, '=SUBTOTAL(109,C2:C6)')
    expect(h.doc.value(s, 7, 2)).toBe(35)
    h.setFilter(s, 'sheet', 0, { values: ['West'], blanks: true })
    expect(hidden()).toEqual([1, 3, 4])
    h.setFilter(s, 'sheet', 0, null)
    h.setFilter(s, 'sheet', 2, { custom: { conds: [{ op: 'greaterThanOrEqual', val: '7' }, { op: 'lessThan', val: '20' }], and: true } })
    expect(hidden()).toEqual([2, 3, 5])
    h.setFilter(s, 'sheet', 1, { custom: { conds: [{ op: 'equal', val: 'P*' }] } }) // begins with P
    expect(hidden()).toEqual([2, 3, 5]) // both filters: Pens(10), Pens(7)
    h.clearFilters(s, 'sheet')
    expect(hidden()).toEqual([])
    h.setFilter(s, 'sheet', 2, { top10: { top: true, percent: false, val: 2 } })
    expect(hidden()).toEqual([2, 4, 5])
    expect(h.meta().sheets[0].filters).toEqual([{ key: 'sheet', area: { sheet: s, r0: 0, c0: 0, r1: 5, c1: 2 }, active: [2] }])
    h.toggleAutoFilter(s, null)
    expect(hidden()).toEqual([])
    expect(h.doc.wb.sheets[0].autoFilter).toBeNull()
  })

  it('tables filter through their own header buttons; undo restores rows', () => {
    const { h, s, hidden } = sales()
    h.createTable(s, parseArea('A1:C6', s)!, true)
    h.setFilter(s, 'Table1', 1, { values: ['Pens'] })
    expect(hidden()).toEqual([2, 3, 5])
    h.doc.undo()
    expect(hidden()).toEqual([])
    h.setTableOptions('Table1', { autoFilter: false })
    expect(h.meta().sheets[0].filters).toEqual([])
  })
})
