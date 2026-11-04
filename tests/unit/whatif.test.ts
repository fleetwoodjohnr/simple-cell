import { describe, expect, it } from 'vitest'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { unzipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'
import { createDataTable, scenarioCells, scenarioSummary, setScenarios, showScenario } from '@/core/whatif'

const soffice = ['/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'].find((p) => existsSync(p))

function book(cells: [string, string][]) {
  const h = new Host()
  const doc = h.doc
  const s = doc.wb.sheets[0].id
  for (const [ref, v] of cells) {
    const m = /^([A-Z])(\d+)$/.exec(ref)!
    doc.input(s, Number(m[2]) - 1, m[1].charCodeAt(0) - 65, v)
  }
  const at = (ref: string) => {
    const m = /^([A-Z])(\d+)$/.exec(ref)!
    return doc.value(s, Number(m[2]) - 1, m[1].charCodeAt(0) - 65)
  }
  return { h, doc, s, at }
}

/** B1 rate, B2 = 1000 × rate; rates down A5:A7 and the formula in B4: a one-variable table on B5:B7 (column input B1). */
const oneVar = (): [string, string][] => [['B1', '0.05'], ['B2', '=1000*B1'], ['B4', '=B2'], ['C4', '=B2*2'], ['A5', '0.01'], ['A6', '0.02'], ['A7', '0.03']]
/** Two variables: corner A4 = 10·B1 + B2; B1 values across B4:D4 (row input), B2 values down A5:A6 (column input). */
const twoVar = (): [string, string][] => [['B1', '1'], ['B2', '1'], ['A4', '=B1*10+B2'], ['B4', '1'], ['C4', '2'], ['D4', '3'], ['A5', '5'], ['A6', '7']]

describe('data tables', () => {
  it('one variable down a column, several formulas across', () => {
    const { doc, s, at } = book(oneVar())
    createDataTable(doc, s, { r0: 3, c0: 0, r1: 6, c1: 2 }, null, { r: 0, c: 1 })
    expect([at('B5'), at('B6'), at('B7')]).toEqual([10, 20, 30])
    expect([at('C5'), at('C6'), at('C7')]).toEqual([20, 40, 60])
    // The input cell is put back.
    expect(at('B1')).toBe(0.05)
    expect(at('B2')).toBe(50)
    // Results follow edits to the formulas and the input values.
    doc.input(s, 1, 1, '=2000*B1')
    expect(at('B5')).toBe(20)
    doc.input(s, 4, 0, '0.5')
    expect(at('B5')).toBe(1000)
  })

  it('one variable across a row (row input)', () => {
    const { doc, s, at } = book([['B1', '3'], ['A2', '=B1^2'], ['B1', '3'], ['B1', '3'], ['C1', '1'], ['D1', '2'], ['E1', '4'], ['C2', ''], ['A3', '=B1+1']])
    // Values across C1:E1 go into B1… but B1 must be outside the table: use row 1 values for input F1.
    doc.input(s, 0, 1, '')
    doc.input(s, 0, 5, '3')
    doc.input(s, 1, 0, '=F1^2')
    doc.input(s, 2, 0, '=F1+1')
    // Selection B1:E3: values in C1:E1, formulas in B2:B3 → results C2:E3, row input F1.
    doc.input(s, 1, 1, '=A2')
    doc.input(s, 2, 1, '=A3')
    createDataTable(doc, s, { r0: 0, c0: 1, r1: 2, c1: 4 }, { r: 0, c: 5 }, null)
    expect([at('C2'), at('D2'), at('E2')]).toEqual([1, 4, 16])
    expect([at('C3'), at('D3'), at('E3')]).toEqual([2, 3, 5])
  })

  it('two variables', () => {
    const { doc, s, at } = book(twoVar())
    createDataTable(doc, s, { r0: 3, c0: 0, r1: 5, c1: 3 }, { r: 0, c: 1 }, { r: 1, c: 1 })
    expect([at('B5'), at('C5'), at('D5')]).toEqual([15, 25, 35])
    expect([at('B6'), at('C6'), at('D6')]).toEqual([17, 27, 37])
    expect(at('A4')).toBe(11)
  })

  it('result cells can’t be edited one by one; the whole table can be cleared; undo brings it back', () => {
    const { doc, s, at } = book(oneVar())
    createDataTable(doc, s, { r0: 3, c0: 0, r1: 6, c1: 2 }, null, { r: 0, c: 1 })
    expect(() => doc.input(s, 4, 1, '5')).toThrow(/part of a data table/)
    expect(() => doc.clear([{ sheet: s, r0: 4, c0: 1, r1: 4, c1: 1 }])).toThrow(/part of a data table/)
    expect(at('B5')).toBe(10)
    doc.clear([{ sheet: s, r0: 4, c0: 1, r1: 6, c1: 2 }])
    expect(doc.wb.sheets[0].dataTables).toEqual([])
    expect(at('B5')).toBe(null)
    doc.undo()
    expect(doc.wb.sheets[0].dataTables).toHaveLength(1)
    expect(at('B5')).toBe(10)
  })

  it('“Automatic except data tables” waits for F9', () => {
    const { h, doc, s, at } = book(oneVar())
    createDataTable(doc, s, { r0: 3, c0: 0, r1: 6, c1: 2 }, null, { r: 0, c: 1 })
    h.setCalc('autoNoTable')
    doc.input(s, 1, 1, '=3000*B1')
    expect(at('B2')).toBe(150)
    expect(at('B5')).toBe(10)
    h.calculate()
    expect(at('B5')).toBe(30)
  })

  it('moves with inserted rows and columns', () => {
    const { doc, s, at } = book(oneVar())
    createDataTable(doc, s, { r0: 3, c0: 0, r1: 6, c1: 2 }, null, { r: 0, c: 1 })
    doc.insertRows(s, 0, 2)
    expect(doc.wb.sheets[0].dataTables[0]).toEqual({ area: { r0: 6, c0: 1, r1: 8, c1: 2 }, rowInput: null, colInput: { r: 2, c: 1 } })
    doc.input(s, 7, 0, '0.04')
    expect(at('B8')).toBe(40)
  })

  it('round trip through .xlsx (dataTable formulas, calcMode) and LibreOffice computes the same results', async () => {
    const { h, doc, s } = book([...twoVar(), ['G1', '0.05'], ['G2', '=1000*G1'], ['H4', '=G2'], ['G5', '0.01'], ['G6', '0.02']])
    createDataTable(doc, s, { r0: 3, c0: 0, r1: 5, c1: 3 }, { r: 0, c: 1 }, { r: 1, c: 1 })
    createDataTable(doc, s, { r0: 3, c0: 6, r1: 5, c1: 7 }, null, { r: 0, c: 6 })
    const saved = (await h.save('xlsx')).data
    const xml = strFromU8(unzipSync(saved)['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('<c r="B5"><f t="dataTable" ref="B5:D6" dt2D="1" dtr="1" r1="B1" r2="B2"/><v>15</v></c>')
    expect(xml).toContain('<f t="dataTable" ref="H5:H6" dt2D="0" dtr="0" r1="G1"/>')
    const back = await Host.open(saved, 'xlsx')
    expect(back.doc.wb.sheets[0].dataTables).toEqual(doc.wb.sheets[0].dataTables)
    expect(back.doc.value(s, 5, 3)).toBe(37)
    back.doc.input(s, 4, 0, '9')
    expect(back.doc.value(s, 4, 1)).toBe(19)
    if (!soffice) return
    // LibreOffice turns Excel data tables into MULTIPLE.OPERATIONS and calculates them: cached values we zero out must come back.
    for (const sh of doc.wb.sheets) for (const t of sh.dataTables) for (let r = t.area.r0; r <= t.area.r1; r++) for (let c = t.area.c0; c <= t.area.c1; c++) sh.set(r, c, { v: 0 })
    const zeroed = (await h.save('xlsx')).data
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-dt-'))
    writeFileSync(path.join(dir, 't.xlsx'), zeroed)
    // A throwaway profile that recalculates Excel files on load (LibreOffice's default keeps the cached values).
    mkdirSync(path.join(dir, 'profile', 'user'), { recursive: true })
    writeFileSync(
      path.join(dir, 'profile', 'user', 'registrymodifications.xcu'),
      '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>'
    )
    execFileSync(soffice, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'profile')).href}`, '--convert-to', 'csv', '--outdir', dir, path.join(dir, 't.xlsx')], { timeout: 120_000 })
    const rows = readFileSync(path.join(dir, 't.csv'), 'utf8').trim().split(/\r?\n/).map((l) => l.split(','))
    expect(rows[4].slice(1, 4)).toEqual(['15', '25', '35'])
    expect(rows[5].slice(1, 4)).toEqual(['17', '27', '37'])
    expect([rows[4][7], rows[5][7]]).toEqual(['10', '20'])
  }, 180_000)
})

describe('scenarios', () => {
  const setup = () => {
    const b = book([['B1', '10'], ['B2', '5'], ['B3', '=B1*B2']])
    const cells = scenarioCells('$B$1:$B$2')
    setScenarios(b.doc, b.s, [
      { name: 'Best', cells, values: ['20', '6'], comment: 'Optimistic', user: 'Pat' },
      { name: 'Worst', cells, values: ['4', '3'] }
    ])
    return b
  }

  it('show puts the values in (one undoable step)', () => {
    const { doc, s, at } = setup()
    showScenario(doc, s, 0)
    expect([at('B1'), at('B2'), at('B3')]).toEqual([20, 6, 120])
    expect(doc.wb.sheets[0].scenarios!.show).toBe(0)
    doc.undo()
    expect([at('B1'), at('B2'), at('B3')]).toEqual([10, 5, 50])
  })

  it('names must be unique; at most 32 changing cells', () => {
    const { doc, s } = setup()
    expect(() => setScenarios(doc, s, [{ name: 'A', cells: [{ r: 0, c: 0 }], values: ['1'] }, { name: 'a', cells: [{ r: 0, c: 0 }], values: ['2'] }])).toThrow(/already a scenario/)
    expect(() => scenarioCells('A1:A33')).toThrow(/at most 32/)
  })

  it('summary report: changing and result cells for current values and each scenario, the sheet unchanged', () => {
    const { doc, s, at } = setup()
    const id = scenarioSummary(doc, s, 'B3')
    const rep = doc.wb.sheetById(id)!
    expect(rep.name).toBe('Scenario Summary')
    const row = (r: number) => [2, 3, 4, 5].map((c) => rep.get(r, c)?.v ?? null)
    expect(row(1)).toEqual([null, 'Current Values:', 'Best', 'Worst'])
    expect(row(3)).toEqual(['$B$1', 10, 20, 4])
    expect(row(4)).toEqual(['$B$2', 5, 6, 3])
    expect(row(6)).toEqual(['$B$3', 50, 120, 12])
    expect([at('B1'), at('B3')]).toEqual([10, 50])
  })

  it('round trip through .xlsx', async () => {
    const { h, doc, s } = setup()
    scenarioSummary(doc, s, 'B3')
    const saved = (await h.save('xlsx')).data
    const xml = strFromU8(unzipSync(saved)['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('<scenarios current="0" show="0" sqref="B3"><scenario name="Best" count="2" user="Pat" comment="Optimistic"><inputCells r="B1" val="20"/><inputCells r="B2" val="6"/></scenario>')
    // Schema order: after sheetData, before mergeCells/conditional formats.
    expect(xml.indexOf('<scenarios')).toBeGreaterThan(xml.indexOf('</sheetData>'))
    const back = await Host.open(saved, 'xlsx')
    expect(back.doc.wb.sheets[0].scenarios).toEqual({ ...doc.wb.sheets[0].scenarios, resultCells: 'B3' })
  })
})
