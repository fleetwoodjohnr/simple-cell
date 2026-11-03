import { describe, expect, it } from 'vitest'
import { readFileSync, writeFileSync, mkdtempSync, existsSync, readdirSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'
import { Doc } from '@/core/doc'
import { readXlsx } from '@/core/formats/xlsx/read'
import { writeXlsx } from '@/core/formats/xlsx/write'
import { CellError } from '@/core/engine/values'
import { parseCell } from '@/core/refs'

const fixture = (n: string) => new Uint8Array(readFileSync(path.join(import.meta.dirname, '../fixtures', n)))
const soffice = ['/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'].find((p) => existsSync(p))

function val(doc: Doc, ref: string) {
  const [s, a] = ref.includes('!') ? ref.split('!') : [doc.wb.sheets[0].name, ref]
  const sh = doc.wb.sheetByName(s)!
  const p = parseCell(a)!
  const v = sh.get(p.r, p.c)?.v ?? null
  return v instanceof CellError ? v.code : v
}
const fmla = (doc: Doc, ref: string) => {
  const [s, a] = ref.includes('!') ? ref.split('!') : [doc.wb.sheets[0].name, ref]
  const p = parseCell(a)!
  return doc.formulaText(doc.wb.sheetByName(s)!.id, p.r, p.c)
}
const style = (doc: Doc, ref: string) => {
  const p = parseCell(ref)!
  return doc.styleAt(doc.wb.sheets[0].id, p.r, p.c)
}

async function open(bytes: Uint8Array) {
  return Host.open(bytes, 'xlsx')
}

describe('XLSX', () => {
  it('reads a workbook written by LibreOffice: values, formulas, styles, merges, sizes, names', async () => {
    const h = await open(fixture('features.xlsx'))
    const d = h.doc
    expect(d.wb.sheets.map((s) => s.name)).toEqual(['Data', 'Summary'])
    expect(val(d, 'A2')).toBe('Widget')
    expect(val(d, 'A4')).toBe('Gizmo & Co "quoted"')
    expect(val(d, 'D2')).toBe(5002)
    expect(fmla(d, 'D2')).toBe('=B2*C2')
    expect(val(d, 'E4')).toBe(true)
    expect(val(d, 'D5')).toBeCloseTo(5203.4)
    expect(val(d, 'E5')).toBe('few')
    expect(val(d, 'Summary!B1')).toBeCloseTo(5203.4)
    expect(fmla(d, 'Summary!B1')).toBe('=SUM(Data!D2:D4)')
    expect(val(d, 'Summary!B2')).toBeCloseTo(199.9)
    expect(val(d, 'Summary!B3')).toBeCloseTo(423.66)
    expect(d.wb.names.find((n) => n.name === 'Prices')).toBeTruthy()
    expect(style(d, 'A1').bold).toBe(true)
    expect(style(d, 'A1').color?.toLowerCase()).toBe('#1f4e79')
    expect(style(d, 'C2').numFmt).toBe('#,##0.00')
    expect(d.display(d.wb.sheets[0].id, 1, 2)?.t).toBe('1,250.50')
    expect(style(d, 'D5').fill?.toLowerCase()).toBe('#ffff00')
    expect(style(d, 'D5').borderTop?.style).toBeTruthy()
    expect(style(d, 'A5').wrap).toBe(true)
    expect(style(d, 'A5').italic).toBe(true)
    expect(d.wb.sheets[0].merges).toEqual([{ sheet: d.wb.sheets[0].id, r0: 4, c0: 0, r1: 4, c1: 1 }])
    expect(d.wb.sheets[0].colWidth(0)).toBeGreaterThan(120)
    expect(d.wb.sheets[0].rowHeight(4)).toBeGreaterThan(30)
    expect(d.display(d.wb.sheets[0].id, 1, 4)?.t).toBe('2024-01-15')
    // Every formula result matches what LibreOffice saved.
    expect(h.warnings.filter((w) => /differ from the values saved/.test(w.message))).toEqual([])
  })

  it('round-trips: everything read back the same', async () => {
    const a = await open(fixture('features.xlsx'))
    a.doc.input(a.doc.wb.sheets[0].id, 1, 1, '6')
    a.doc.freeze(a.doc.wb.sheets[0].id, 1, 1)
    const out = (await a.save('xlsx')).data
    const b = await open(out)
    const d = b.doc
    expect(val(d, 'D2')).toBe(7503)
    expect(fmla(d, 'D2')).toBe('=B2*C2')
    expect(val(d, 'E5')).toBe('many')
    expect(val(d, 'Summary!B1')).toBeCloseTo(7704.4)
    expect(style(d, 'C2').numFmt).toBe('#,##0.00')
    expect(style(d, 'A1').bold).toBe(true)
    expect(style(d, 'D5').fill?.toLowerCase()).toBe('#ffff00')
    expect(d.wb.sheets[0].merges.length).toBe(1)
    expect(d.wb.sheets[0].freeze).toEqual({ rows: 1, cols: 1 })
    expect(d.wb.names.map((n) => n.name)).toContain('Prices')
    expect(b.warnings.filter((w) => /differ from the values saved/.test(w.message))).toEqual([])
  })

  it.skipIf(!soffice)('LibreOffice reads the files Simple Cell writes', async () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    doc.input(s, 0, 0, 'Name')
    doc.input(s, 0, 1, 'Value')
    doc.input(s, 1, 0, 'pi')
    doc.input(s, 1, 1, '=ROUND(PI(),4)')
    doc.input(s, 2, 0, 'date')
    doc.input(s, 2, 1, '1/15/2024')
    doc.input(s, 3, 1, '=SUM(B2:B3)')
    doc.input(s, 4, 1, '=TEXTJOIN("-",TRUE,A1:A3)')
    doc.format([{ sheet: s, r0: 0, c0: 0, r1: 0, c1: 1 }], { bold: true, fill: '#ffcc00' })
    const h = new Host(doc)
    const out = (await h.save('xlsx')).data
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-xlsx-'))
    writeFileSync(path.join(dir, 'ours.xlsx'), out)
    execFileSync(soffice!, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,1033,false,true,false,false,false,-1', '--outdir', dir, path.join(dir, 'ours.xlsx')], { timeout: 120_000 })
    const csv = readFileSync(path.join(dir, readdirSync(dir).find((f) => f.endsWith('.csv'))!), 'utf8')
    expect(csv).toContain('Name,Value')
    expect(csv).toContain('pi,3.1416')
    expect(csv).toContain('date,01/15/2024')
    expect(csv).toContain(',45309.1416')
    expect(csv).toContain(',Name-pi-date')
  }, 150_000)

  it('keeps parts and elements it does not understand: drawings, charts, customXml, extLst, VBA', async () => {
    // A hand-made workbook in Excel's layout with extras Simple Cell doesn't model.
    const ws = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData><conditionalFormatting sqref="A1:A2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting><drawing r:id="rId1"/><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings/></ext></extLst></worksheet>`
    const files: Record<string, Uint8Array> = {
      '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/customXml/item1.xml" ContentType="application/xml"/></Types>`),
      '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="7"/><workbookPr codeName="ThisWorkbook"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}"><x15:workbookPr xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" chartTrackingRefBase="1"/></ext></extLst></workbook>`),
      'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`),
      'xl/worksheets/sheet1.xml': strToU8(ws),
      'xl/worksheets/_rels/sheet1.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`),
      'xl/drawings/drawing1.xml': strToU8(`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`),
      'xl/drawings/_rels/drawing1.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`),
      'xl/charts/chart1.xml': strToU8(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Sheet1!$A$1:$A$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`),
      'customXml/item1.xml': strToU8('<data xmlns="urn:example">keep me</data>'),
      'xl/vbaProject.bin': new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4])
    }
    const bytes = zipSync(files)
    const h = await Host.open(bytes, 'xlsm')
    expect(h.warnings.some((w) => /VBA macros/.test(w.message))).toBe(true)
    expect(h.warnings.some((w) => /Charts, pictures and shapes are kept/.test(w.message))).toBe(true)
    h.doc.input(h.doc.wb.sheets[0].id, 1, 0, '5')
    const out = unzipSync((await h.save('xlsm')).data)
    const names = Object.keys(out)
    expect(names).toContain('xl/drawings/drawing1.xml')
    expect(names).toContain('xl/charts/chart1.xml')
    expect(names).toContain('customXml/item1.xml')
    expect(strFromU8(out['customXml/item1.xml'])).toContain('keep me')
    expect([...out['xl/vbaProject.bin']]).toEqual([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4])
    const sheet = strFromU8(out['xl/worksheets/sheet1.xml'])
    expect(sheet).toMatch(/<sheetData>.*<\/sheetData><conditionalFormatting sqref="A1:A2">.*<\/conditionalFormatting><drawing r:id="rId1"\/><extLst>/)
    expect(strFromU8(out['[Content_Types].xml'])).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    expect(strFromU8(out['xl/workbook.xml'])).toContain('codeName="ThisWorkbook"')
    expect(strFromU8(out['xl/workbook.xml'])).toContain('chartTrackingRefBase')
    // Inserting rows moves preserved content too: conditional format range, chart series, drawing anchor.
    h.doc.insertRows(h.doc.wb.sheets[0].id, 0, 2)
    const moved = unzipSync((await h.save('xlsm')).data)
    expect(strFromU8(moved['xl/worksheets/sheet1.xml'])).toContain('<conditionalFormatting sqref="A3:A4">')
    expect(strFromU8(moved['xl/charts/chart1.xml'])).toContain('<c:f>Sheet1!$A$3:$A$4</c:f>')
    expect(strFromU8(moved['xl/drawings/drawing1.xml'])).toContain('<xdr:row>4</xdr:row>')
    h.doc.undo()
    const undone = unzipSync((await h.save('xlsm')).data)
    expect(strFromU8(undone['xl/charts/chart1.xml'])).toContain('<c:f>Sheet1!$A$1:$A$2</c:f>')
    expect(strFromU8(undone['xl/drawings/drawing1.xml'])).toContain('<xdr:row>2</xdr:row>')
    // Saved as .xlsx, the macros go (with a warning).
    const plain = await h.save('xlsx')
    expect(Object.keys(unzipSync(plain.data))).not.toContain('xl/vbaProject.bin')
    expect(plain.warnings.some((w) => /VBA macros are removed/.test(w.message))).toBe(true)
  })

  it('dynamic arrays spill after a round trip and are marked for Excel', async () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    doc.input(s, 0, 0, '=SEQUENCE(3)')
    doc.input(s, 0, 1, '=SUM(A1#)')
    doc.input(s, 0, 2, '={1,2;3,4}', { array: { sheet: s, r0: 0, c0: 2, r1: 1, c1: 3 } })
    const out = writeXlsx(doc.wb).data
    const parts = unzipSync(out)
    const sheet = strFromU8(parts['xl/worksheets/sheet1.xml'])
    expect(sheet).toContain('<c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f><v>1</v></c>')
    expect(sheet).toContain('_xlfn.ANCHORARRAY(A1)')
    expect(Object.keys(parts)).toContain('xl/metadata.xml')
    const back = new Doc(readXlsx(out).workbook)
    expect(val(back, 'A3')).toBe(3)
    expect(val(back, 'B1')).toBe(6)
    expect(val(back, 'D2')).toBe(4)
  })

  it('conditional formatting: read, shown, edited and written back', async () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    ;[5, 15, 25, 35].forEach((v, i) => doc.input(s, i, 0, String(v)))
    const area = (r0: number, r1: number) => ({ sheet: s, r0, c0: 0, r1, c1: 0 })
    doc.addCondFormat(s, [area(0, 3)], { type: 'cellIs', operator: 'between', formulas: ['10', '30'], dxf: { fill: '#ffc7ce', color: '#9c0006', bold: true }, stopIfTrue: true })
    doc.addCondFormat(s, [area(0, 3)], { type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }], colors: ['#f8696b', '#ffeb84', '#63be7b'] } })
    doc.addCondFormat(s, [area(0, 3)], { type: 'iconSet', iconSet: { set: '3Arrows', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }] } })
    doc.addCondFormat(s, [area(0, 3), { sheet: s, r0: 0, c0: 2, r1: 3, c1: 2 }], { type: 'expression', formulas: ['=MOD(ROW(),2)=0'], dxf: { numFmt: '0.0' } })
    const out = await new Host(doc).save('xlsx')
    const xml = strFromU8(unzipSync(out.data)['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('<conditionalFormatting sqref="A1:A4 C1:C4"><cfRule type="expression" dxfId="1" priority="1"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>')
    expect(xml).toContain('<cfRule type="cellIs" dxfId="0" priority="4" stopIfTrue="1" operator="between"><formula>10</formula><formula>30</formula></cfRule>')
    const styles = strFromU8(unzipSync(out.data)['xl/styles.xml'])
    expect(styles).toMatch(/<dxfs count="2"><dxf><font><b\/><color rgb="FF9C0006"\/><\/font><fill><patternFill><bgColor rgb="FFFFC7CE"\/><\/patternFill><\/fill><\/dxf><dxf><numFmt numFmtId="\d+" formatCode="0.0"\/><\/dxf><\/dxfs>/)
    const back = await open(out.data)
    const list = back.doc.listCondFormats(back.doc.wb.sheets[0].id)
    expect(list.map((r) => [r.type, r.priority, r.ranges])).toEqual([
      ['expression', 1, 'A1:A4 C1:C4'],
      ['iconSet', 2, 'A1:A4'],
      ['colorScale', 3, 'A1:A4'],
      ['cellIs', 4, 'A1:A4']
    ])
    expect(list[3].dxf).toEqual({ fill: '#FFC7CE', color: '#9C0006', bold: true })
    const cells = back.viewport(back.doc.wb.sheets[0].id, 0, 3, 0, 0)
    const a2 = cells.find((c) => c.r === 1)!
    expect(a2.t).toBe('15.0')
    expect(a2.icon).toEqual({ set: '3Arrows', i: 1 })
    expect(a2.x?.bold).toBe(true)
    // An unchanged file keeps its dxf ids (tables and PivotTables refer to them too).
    const again = strFromU8(unzipSync((await back.save('xlsx')).data)['xl/styles.xml'])
    expect(again.match(/<dxfs[\s\S]*<\/dxfs>/)?.[0]).toBe(styles.match(/<dxfs[\s\S]*<\/dxfs>/)?.[0])
  })

  it.skipIf(!soffice)('LibreOffice reads the conditional formats Simple Cell writes', async () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    ;[5, 15, 25].forEach((v, i) => doc.input(s, i, 0, String(v)))
    const a = { sheet: s, r0: 0, c0: 0, r1: 2, c1: 0 }
    doc.addCondFormat(s, [a], { type: 'cellIs', operator: 'greaterThan', formulas: ['10'], dxf: { fill: '#ffc7ce' } })
    doc.addCondFormat(s, [a], { type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['#f8696b', '#63be7b'] } })
    doc.addCondFormat(s, [a], { type: 'iconSet', iconSet: { set: '3Arrows', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }] } })
    doc.addCondFormat(s, [a], { type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '#638ec6' } })
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-cf-'))
    writeFileSync(path.join(dir, 'cf.xlsx'), (await new Host(doc).save('xlsx')).data)
    execFileSync(soffice!, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'ods', '--outdir', dir, path.join(dir, 'cf.xlsx')], { timeout: 120_000 })
    const content = strFromU8(unzipSync(new Uint8Array(readFileSync(path.join(dir, 'cf.ods'))))['content.xml'])
    expect(content).toContain('calcext:condition calcext:apply-style-name')
    expect(content).toContain('calcext:value="&gt;10"')
    expect(content).toContain('<calcext:color-scale>')
    expect(content).toContain('calcext:icon-set-type="3Arrows"')
    expect(content).toContain('<calcext:data-bar')
  }, 150_000)

  it('conditional formatting from Excel 2010+: data bar extensions stay linked', async () => {
    const h = await open(fixture('features.xlsx'))
    const sh = h.doc.wb.sheets[0]
    const carry = h.doc.wb.carry as { sheets: Record<number, { elements: { name: string; xml: string }[] }> }
    const id = '{1A2B3C4D-0000-4000-8000-000000000001}'
    carry.sheets[sh.id].elements.push({ name: 'extLst', xml: `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100" negativeBarColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/></x14:dataBar></x14:cfRule><xm:sqref>B2:B4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>` })
    // Read a main-part data bar that links to it, as Excel writes them.
    const { readCondFormat } = await import('@/core/formats/xlsx/condfmt')
    const b = readCondFormat(`<conditionalFormatting sqref="B2:B4"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>${id}</x14:id></ext></extLst></cfRule></conditionalFormatting>`, sh.id, [], () => '#638EC6', () => null)!
    sh.condFormats.push(b)
    const xml = strFromU8(unzipSync((await h.save('xlsx')).data)['xl/worksheets/sheet1.xml'])
    expect(xml).toContain(`<x14:id>${id}</x14:id></ext></extLst></cfRule></conditionalFormatting>`)
    expect(xml).toContain(`<x14:cfRule type="dataBar" id="${id}">`)
  })

  it('data validation: read, enforced, written back; LibreOffice reads it', async () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    doc.input(s, 0, 3, 'Red')
    doc.input(s, 1, 3, 'Blue')
    doc.setValidation(s, [{ sheet: s, r0: 0, c0: 0, r1: 9, c1: 0 }], { type: 'list', formulas: ['=$D$1:$D$2'], showInput: true, prompt: 'Pick a colour', errorStyle: 'warning', error: 'Not a colour' })
    doc.setValidation(s, [{ sheet: s, r0: 0, c0: 1, r1: 9, c1: 1 }], { type: 'whole', operator: 'greaterThan', formulas: ['0'] })
    doc.setValidation(s, [{ sheet: s, r0: 0, c0: 2, r1: 0, c1: 2 }], { type: 'list', formulas: ['Yes, No'], noDropdown: true })
    const out = await new Host(doc).save('xlsx')
    const xml = strFromU8(unzipSync(out.data)['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('<dataValidations count="3"><dataValidation type="list" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" error="Not a colour" prompt="Pick a colour" sqref="A1:A10"><formula1>$D$1:$D$2</formula1></dataValidation>')
    expect(xml).toContain('<dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B1:B10"><formula1>0</formula1></dataValidation>')
    expect(xml).toContain('<dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" sqref="C1"><formula1>"Yes, No"</formula1></dataValidation>')
    const back = await open(out.data)
    const bs = back.doc.wb.sheets[0].id
    expect(back.validationInfo(bs, 3, 0)).toMatchObject({ type: 'list', items: ['Red', 'Blue'], prompt: 'Pick a colour', errorStyle: 'warning' })
    expect(back.input(bs, 0, 1, '-3').ok).toBe(false)
    expect(back.input(bs, 0, 2, 'no').ok).toBe(true)
    expect(back.validationInfo(bs, 0, 2)?.items).toBeUndefined() // dropdown hidden
    // Unchanged on a second save.
    expect(strFromU8(unzipSync((await back.save('xlsx')).data)['xl/worksheets/sheet1.xml'])).toContain(xml.match(/<dataValidations[\s\S]*<\/dataValidations>/)![0])
    if (!soffice) return
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-dv-'))
    writeFileSync(path.join(dir, 'dv.xlsx'), out.data)
    execFileSync(soffice, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'ods', '--outdir', dir, path.join(dir, 'dv.xlsx')], { timeout: 120_000 })
    const content = strFromU8(unzipSync(new Uint8Array(readFileSync(path.join(dir, 'dv.ods'))))['content.xml'])
    expect(content).toContain('table:condition="of:cell-content-is-in-list([.$D$1:.$D$2])"')
    expect(content).toContain('table:condition="of:cell-content-is-in-list(&quot;Yes&quot;;&quot;No&quot;)" table:allow-empty-cell="true" table:display-list="no"')
    expect(content).toContain('of:cell-content-is-whole-number() and cell-content()&gt;0')
  }, 150_000)

  it('refuses zip bombs and non-zip files', () => {
    expect(() => readXlsx(new Uint8Array([1, 2, 3, 4]))).toThrow(/not a valid workbook/)
    const big = new Uint8Array(30 * 1024 * 1024) // 30 MB of zeros compresses to ~30 KB
    const bomb = zipSync({ 'xl/workbook.xml': big }, { level: 9 })
    const prev = (globalThis as Record<string, unknown>).x
    void prev
    expect(() => readXlsx(bomb)).toThrow(/zip bomb|compressed suspiciously/)
  })

  it('reports formulas whose saved results differ from Simple Cell’s', async () => {
    const ws = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>5</v></c></row></sheetData></worksheet>`
    const files = {
      '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'),
      '_rels/.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': strToU8('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': strToU8(ws)
    }
    const h = await Host.open(zipSync(files), 'xlsx')
    expect(val(h.doc, 'B1')).toBe(4)
    expect(h.warnings.some((w) => /1 formula result\(s\) differ.*S!B1/.test(w.message))).toBe(true)
  })

  it('formulas using functions Simple Cell lacks keep the saved value', async () => {
    const ws = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>_xlfn.STOCKHISTORY("MSFT",1)</f><v>412.5</v></c></row></sheetData></worksheet>`
    const files = {
      '_rels/.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': strToU8('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': strToU8(ws)
    }
    const h = await Host.open(zipSync(files), 'xlsx')
    expect(val(h.doc, 'A1')).toBe(412.5)
    expect(h.doc.display(h.doc.wb.sheets[0].id, 0, 0)?.st).toBe(1)
    const back = strFromU8(unzipSync((await h.save('xlsx')).data)['xl/worksheets/sheet1.xml'])
    expect(back).toContain('<f>_xlfn.STOCKHISTORY("MSFT",1)</f><v>412.5</v>')
  })
})
