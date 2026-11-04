import { describe, expect, it } from 'vitest'
import { Host } from '@/core/host'
import { parseArea } from '@/core/refs'

// Region | Item | Units, plus a formula column Double = Units*2, as Table1 on A1:D5.
function sales() {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  const rows = [
    ['Region', 'Item', 'Units', 'Double'],
    ['East', 'Pens', '10', '=C2*2'],
    ['West', 'Ink', '5', '=C3*2'],
    ['East', 'Paper', '25', '=C4*2'],
    ['North', 'Pens', '7', '=C5*2']
  ]
  rows.forEach((row, r) => row.forEach((v, c) => h.doc.input(s, r, c, v)))
  expect(h.createTable(s, parseArea('A1:D5', s)!, true).ok).toBe(true)
  const t = h.doc.wb.tables[0]
  const col = (c: number, r0 = 0, r1 = 5) => Array.from({ length: r1 - r0 + 1 }, (_, i) => h.doc.value(s, r0 + i, c))
  return { h, s, t, col }
}

describe('SQL on tables', () => {
  it('a SELECT goes to a new sheet as ordinary cells; undo removes it', async () => {
    const { h, s, col } = sales()
    const r = await h.runSql("SELECT Item, Units FROM Table1 WHERE Region = 'East' ORDER BY Units DESC", { kind: 'sheet' })
    expect(r).toMatchObject({ ok: true, rows: 2 })
    const q = h.doc.wb.sheetByName('Query')!
    expect([0, 1, 2].map((r) => [h.doc.value(q.id, r, 0), h.doc.value(q.id, r, 1)])).toEqual([['Item', 'Units'], ['Paper', 25], ['Pens', 10]])
    expect(h.doc.wb.tables).toHaveLength(1)
    expect(col(0)).toEqual(['Region', 'East', 'West', 'East', 'North', null])
    h.doc.undo()
    expect(h.doc.wb.sheetByName('Query')).toBeUndefined()
    expect(h.doc.wb.tableByName('Table2')).toBeUndefined()
    void s
  })

  it('UPDATE edits only the changed cells, keeping formulas; undo restores', async () => {
    const { h, s, col } = sales()
    expect(await h.runSql("UPDATE Table1 SET Units = Units + 1 WHERE Item = 'Pens'", { kind: 'sheet' })).toMatchObject({ ok: true, changed: ['Table1'], rows: null })
    expect(col(2, 1, 4)).toEqual([11, 5, 25, 8])
    // The formula column follows the new units and is still a formula.
    expect(col(3, 1, 4)).toEqual([22, 10, 50, 16])
    expect(h.doc.formulaText(s, 1, 3)).toBe('=C2*2')
    expect(h.doc.wb.sheetByName('Query')).toBeUndefined()
    h.doc.undo()
    expect(col(2, 1, 4)).toEqual([10, 5, 25, 7])
  })

  it('DELETE shrinks the table and clears the rows it no longer covers', async () => {
    const { h, s, t, col } = sales()
    expect((await h.runSql("DELETE FROM Table1 WHERE Region = 'East'", { kind: 'sheet' })).ok).toBe(true)
    expect(t.area.r1).toBe(2)
    expect(col(1)).toEqual(['Item', 'Ink', 'Pens', null, null, null])
    expect(h.doc.value(s, 3, 3)).toBeNull()
    h.doc.undo()
    expect(t.area.r1).toBe(4)
    expect(col(1)).toEqual(['Item', 'Pens', 'Ink', 'Paper', 'Pens', null])
  })

  it('INSERT grows the table, but not over data below it', async () => {
    const { h, s, t } = sales()
    expect((await h.runSql("INSERT INTO Table1 (Region, Item, Units) VALUES ('South', 'Tape', 3)", { kind: 'sheet' })).ok).toBe(true)
    expect(t.area.r1).toBe(5)
    expect([h.doc.value(s, 5, 0), h.doc.value(s, 5, 2)]).toEqual(['South', 3])
    h.doc.undo()
    h.doc.input(s, 6, 0, 'note')
    const r = await h.runSql("INSERT INTO Table1 VALUES ('a','b',1,2), ('c','d',3,4)", { kind: 'sheet' })
    expect(r).toMatchObject({ ok: false })
    expect(r.ok === false && r.error).toContain('A7')
    expect(t.area.r1).toBe(4)
  })

  it('queries and edits a plain range without making a table', async () => {
    const h = new Host(), s = h.doc.wb.sheets[0].id
    h.doc.input(s, 0, 0, 'Units'); h.doc.input(s, 1, 0, '3'); h.doc.input(s, 2, 0, '7')
    const source = { area: parseArea('A1:A3', s)!, headers: true }
    expect(await h.sqlPreview('SELECT SUM(Units) AS Total FROM Selection', source)).toMatchObject({ ok: true, rows: [[10]] })
    expect(await h.runSql('UPDATE Selection SET Units = Units * 2', { kind: 'sheet' }, source)).toMatchObject({ ok: true })
    expect(h.doc.value(s, 1, 0)).toBe(6)
    h.doc.undo(); expect(h.doc.value(s, 1, 0)).toBe(3)
    const out = await h.runSql('SELECT * FROM Selection', { kind: 'sheet' }, source)
    expect(out.ok).toBe(true); expect(h.doc.wb.tables).toHaveLength(0)
  })

  it('rolls source edits back when the output collides, and rejects schema edits', async () => {
    const { h, s, col } = sales()
    expect(await h.runSql('UPDATE Table1 SET Units=99; SELECT * FROM Table1', { kind: 'cell', sheet: s, r: 0, c: 0 })).toMatchObject({ ok: false })
    expect(col(2, 1, 4)).toEqual([10, 5, 25, 7])
    expect(await h.runSql('DROP TABLE Table1', { kind: 'sheet' })).toMatchObject({ ok: false })
  })

  it('a SELECT at a cell refuses to cover data', async () => {
    const { h, s } = sales()
    h.doc.input(s, 1, 7, 'busy')
    const r = await h.runSql('SELECT Item FROM Table1', { kind: 'cell', sheet: s, r: 0, c: 7 })
    expect(r.ok === false && r.error).toContain('H2')
    expect((await h.runSql("SELECT Item FROM Table1 WHERE Units > 20", { kind: 'cell', sheet: s, r: 0, c: 9 })).ok).toBe(true)
    expect([h.doc.value(s, 0, 9), h.doc.value(s, 1, 9)]).toEqual(['Item', 'Paper'])
  })

  it('preview changes nothing; SQL errors come back as messages', async () => {
    const { h, col } = sales()
    const p = await h.sqlPreview('DELETE FROM Table1; SELECT COUNT(*) AS n FROM Table1')
    expect(p).toMatchObject({ ok: true, columns: ['n'], rows: [[0]], rowCount: 1, changed: [{ table: 'Table1', rows: 0 }] })
    expect(col(0)).toEqual(['Region', 'East', 'West', 'East', 'North', null])
    const bad = await h.sqlPreview('SELEC * FROM Table1')
    expect(bad.ok === false && bad.error).toMatch(/syntax error/)
    expect(h.sqlTables()).toEqual([{ name: 'Table1', sheet: expect.any(Number), columns: ['Region', 'Item', 'Units', 'Double'] }])
  })
})


describe('SQL range boundaries', () => {
  it('supports headerless datasets, deletes, redo and generated duplicate headers', async () => {
    const h = new Host(), sheet = h.doc.wb.sheets[0].id
    h.doc.input(sheet, 0, 0, '3'); h.doc.input(sheet, 1, 0, '7')
    const source = { area: parseArea('A1:A2', sheet)!, headers: false }
    expect(await h.sqlPreview('SELECT SUM(Column1) AS Total FROM Selection', source)).toMatchObject({ ok: true, rows: [[10]] })
    expect(await h.runSql('DELETE FROM Selection WHERE Column1=3', { kind: 'sheet' }, source)).toMatchObject({ ok: true })
    expect(h.doc.value(sheet, 0, 0)).toBe(7); expect(h.doc.value(sheet, 1, 0)).toBeNull()
    h.doc.undo(); expect(h.doc.value(sheet, 0, 0)).toBe(3)
    h.doc.redo(); expect(h.doc.value(sheet, 0, 0)).toBe(7)
    h.doc.input(sheet, 0, 0, 'Name'); h.doc.input(sheet, 0, 1, 'Name')
    expect(await h.sqlRange({ area: parseArea('A1:C2', sheet)!, headers: true })).toMatchObject({ columns: ['Name', 'Name2', 'Column3'] })
  })

  it('rejects overlapping table writes and invalid output bounds; preview is harmless', async () => {
    const { h, s, col } = sales()
    const source = { area: parseArea('A1:D5', s)!, headers: true }
    expect(await h.sqlPreview('UPDATE Selection SET Units=100', source)).toMatchObject({ ok: false })
    expect(col(2, 1, 4)).toEqual([10, 5, 25, 7])
    expect(await h.runSql('SELECT 1', { kind: 'cell', sheet: s, r: -1, c: 0 })).toMatchObject({ ok: false })
    expect(await h.runSql('SELECT 1, 2', { kind: 'cell', sheet: s, r: 0, c: 16383 })).toMatchObject({ ok: false })
  })

  it('keeps zero-result headers and handles quoted column names and joins', async () => {
    const { h, s } = sales()
    h.doc.input(s, 0, 7, 'Unit Price'); h.doc.input(s, 1, 7, '2')
    const source = { area: parseArea('H1:H2', s)!, headers: true }
    expect(await h.sqlPreview('SELECT SUM(Units * "Unit Price") AS Total FROM Table1 CROSS JOIN Selection', source)).toMatchObject({ ok: true, rows: [[94]] })
    expect(await h.runSql('SELECT Item FROM Table1 WHERE 0', { kind: 'sheet' })).toMatchObject({ ok: true, rows: 0 })
    expect(h.doc.value(h.doc.wb.sheetByName('Query')!.id, 0, 0)).toBe('Item')
  })
})


it('rejects oversized sources before copying whole worksheet columns', async () => {
  const h = new Host(), sheet = h.doc.wb.sheets[0].id
  const source = { area: parseArea('A1:A1048576', sheet)!, headers: false }
  const result = await h.sqlPreview('SELECT * FROM Selection LIMIT 1', source)
  expect(result).toMatchObject({ ok: false })
  expect(!result.ok && result.error).toContain('one million cells')
})
