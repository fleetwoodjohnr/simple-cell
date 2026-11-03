import { describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, readFileSync, writeFileSync, readdirSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { unzipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'
import { emptyPivot } from '@/core/pivot'

const soffice = ['/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'].find((p) => existsSync(p))

function book() {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  ;[['Region', 'Product', 'Sales'], ['East', 'Pens', '10'], ['West', 'Ink', '5'], ['East', 'Ink', '2']].forEach((row, r) => row.forEach((v, c) => h.doc.input(s, r, c, v)))
  const p = h.createPivot(s, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 2 }, { sheet: s, r: 0, c: 5 }) as { sheet: number; id: number }
  h.setPivotConfig(s, p.id, { ...emptyPivot(), rows: ['Region'], cols: ['Product'], values: [{ field: 'Sales', agg: 'sum' }] })
  return { h, s, id: p.id }
}

describe('PivotTables in XLSX', () => {
  it('writes cache and table definitions Excel refreshes on open, and reads them back', async () => {
    const { h, s } = book()
    const files = unzipSync((await h.save('xlsx')).data)
    const cache = strFromU8(files['xl/pivotCache/pivotCacheDefinition1.xml'])
    expect(cache).toContain('refreshOnLoad="1" saveData="0"')
    expect(cache).toContain('<worksheetSource ref="A1:C4" sheet="Sheet1"/>')
    expect(cache).toContain('<cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField>')
    const table = strFromU8(files['xl/pivotTables/pivotTable1.xml'])
    expect(table).toContain('<rowFields count="1"><field x="0"/></rowFields><colFields count="1"><field x="1"/></colFields>')
    expect(table).toContain('<dataField name="Sum of Sales" fld="2" baseField="0" baseItem="0"/>')
    expect(strFromU8(files['xl/workbook.xml'])).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/)
    expect(strFromU8(files['[Content_Types].xml'])).toContain('/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"')
    const back = await Host.open((await h.save('xlsx')).data, 'xlsx')
    const bs = back.doc.wb.sheets[0]
    expect(bs.pivots.map((p) => ({ name: p.name, source: p.source, config: p.config }))).toEqual(h.doc.wb.sheets[0].pivots.map((p) => ({ name: p.name, source: p.source, config: p.config })))
    // Unchanged: the same parts again. Refreshing after a source change works on the reopened file.
    const again = unzipSync((await back.save('xlsx')).data)
    expect(strFromU8(again['xl/pivotTables/pivotTable1.xml'])).toBe(table)
    back.doc.input(bs.id, 1, 2, '30')
    back.refreshPivots()
    expect(back.doc.value(bs.id, 2, 7)).toBe(30) // East × Pens
    // Deleted: its parts go.
    back.deletePivot(bs.id, bs.pivots[0].id)
    const gone = unzipSync((await back.save('xlsx')).data)
    expect(Object.keys(gone).some((p) => /pivotTable/.test(p))).toBe(false)
    void s
  })

  it.skipIf(!soffice)('LibreOffice reads the PivotTables Simple Cell writes', async () => {
    const { h } = book()
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-pivot-'))
    writeFileSync(path.join(dir, 'p.xlsx'), (await h.save('xlsx')).data)
    execFileSync(soffice!, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'ods', '--outdir', dir, path.join(dir, 'p.xlsx')], { timeout: 120_000 })
    const content = strFromU8(unzipSync(new Uint8Array(readFileSync(path.join(dir, readdirSync(dir).find((f) => f.endsWith('.ods'))!))))['content.xml'])
    expect(content).toContain('table:data-pilot-table table:name="PivotTable1"')
    expect(content).toMatch(/table:source-field-name="Region"[^>]*table:orientation="row"/)
    expect(content).toMatch(/table:source-field-name="Product"[^>]*table:orientation="column"/)
    expect(content).toMatch(/table:source-field-name="Sales"[^>]*table:orientation="data"[^>]*table:function="sum"/)
  }, 150_000)
})
