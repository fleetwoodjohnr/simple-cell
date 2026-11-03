import { describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, readFileSync, writeFileSync, readdirSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'
import { Host } from '@/core/host'

const soffice = ['/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'].find((p) => existsSync(p))

function book() {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  ;[['Month', 'North', 'South'], ['Jan', '10', '7'], ['Feb', '12', '9'], ['Mar', '15', '11']].forEach((row, r) => row.forEach((v, c) => h.doc.input(s, r, c, v)))
  return { h, s }
}

describe('charts in XLSX', () => {
  it('a new chart writes a drawing and a chart part, and reads back the same', async () => {
    const { h, s } = book()
    const { id } = h.addChartFromRange(s, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 2 }) as { id: number }
    h.updateChart(s, id, { title: 'Sales', legend: 'bottom', grouping: 'stacked', dataLabels: true, yTitle: 'Units' })
    const files = unzipSync((await h.save('xlsx')).data)
    expect(Object.keys(files)).toEqual(expect.arrayContaining(['xl/drawings/drawing1.xml', 'xl/charts/chart1.xml', 'xl/drawings/_rels/drawing1.xml.rels']))
    const chart = strFromU8(files['xl/charts/chart1.xml'])
    expect(chart).toContain('<c:barDir val="col"/><c:grouping val="stacked"/>')
    expect(chart).toContain('<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>')
    expect(chart).toContain('<c:val><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f></c:numRef></c:val>')
    expect(strFromU8(files['xl/worksheets/sheet1.xml'])).toMatch(/<drawing r:id="rId\d+"\/>/)
    expect(strFromU8(files['[Content_Types].xml'])).toContain('/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"')
    const back = await Host.open((await h.save('xlsx')).data, 'xlsx')
    const [c] = back.doc.wb.sheets[0].charts
    const orig = h.doc.wb.sheets[0].charts[0]
    expect({ ...c, id: 0, carry: undefined }).toEqual({ ...orig, id: 0, carry: undefined })
    // Saved again unchanged: the chart part is byte for byte the same.
    const again = unzipSync((await back.save('xlsx')).data)
    expect(strFromU8(again['xl/charts/chart1.xml'])).toBe(chart)
    // Moving it rewrites only the anchor; deleting the only chart removes the drawing.
    back.updateChart(back.doc.wb.sheets[0].id, c.id, { anchor: { ...c.anchor, r0: c.anchor.r0 + 5, r1: c.anchor.r1 + 5 } })
    const moved = unzipSync((await back.save('xlsx')).data)
    expect(strFromU8(moved['xl/charts/chart1.xml'])).toBe(chart)
    expect(strFromU8(moved['xl/drawings/drawing1.xml'])).toContain(`<xdr:row>${c.anchor.r0 + 5}</xdr:row>`)
    back.deleteChart(back.doc.wb.sheets[0].id, c.id)
    const gone = unzipSync((await back.save('xlsx')).data)
    expect(Object.keys(gone).some((p) => /drawing|charts/.test(p))).toBe(false)
    expect(strFromU8(gone['xl/worksheets/sheet1.xml'])).not.toContain('<drawing')
  })

  it('pictures in the same drawing are kept when a chart is added', async () => {
    const { h, s } = book()
    const first = unzipSync((await h.save('xlsx')).data)
    // Add a picture anchor by hand, as another app would have.
    first['xl/drawings/drawing1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill/><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`)
    first['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
    first['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(first['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>'))
    first['[Content_Types].xml'] = strToU8(strFromU8(first['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'))
    const opened = await Host.open(zipSync(first), 'xlsx')
    expect(opened.warnings.some((w) => /Pictures, shapes/.test(w.message))).toBe(true)
    opened.addChartFromRange(opened.doc.wb.sheets[0].id, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 1 }, 'pie')
    const out = unzipSync((await opened.save('xlsx')).data)
    const drawing = strFromU8(out['xl/drawings/drawing1.xml'])
    expect(drawing).toContain('name="Logo"')
    expect(drawing).toMatch(/<c:chart [^>]*r:id="rId\d+"\/>/)
    expect(strFromU8(out[Object.keys(out).find((p) => /charts\/chart\d+\.xml$/.test(p))!])).toContain('<c:pieChart>')
  })

  it.skipIf(!soffice)('LibreOffice reads the charts Simple Cell writes', async () => {
    const { h, s } = book()
    h.addChartFromRange(s, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 2 }, 'line')
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-chart-'))
    writeFileSync(path.join(dir, 'c.xlsx'), (await h.save('xlsx')).data)
    execFileSync(soffice!, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'ods', '--outdir', dir, path.join(dir, 'c.xlsx')], { timeout: 120_000 })
    const files = unzipSync(new Uint8Array(readFileSync(path.join(dir, readdirSync(dir).find((f) => f.endsWith('.ods'))!))))
    const obj = Object.keys(files).find((p) => /^Object \d+\/content\.xml$/.test(p))
    expect(obj).toBeTruthy()
    const content = strFromU8(files[obj!])
    expect(content).toContain('chart:class="chart:line"')
    expect(content).toContain('chart:values-cell-range-address="Sheet1.B2:Sheet1.B4"')
    expect(content).toContain('chart:values-cell-range-address="Sheet1.C2:Sheet1.C4"')
    expect(content).toContain('table:cell-range-address="Sheet1.A2:Sheet1.A4"')
  }, 150_000)
})
