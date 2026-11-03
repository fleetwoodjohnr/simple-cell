import { describe, expect, it } from 'vitest'
import { readFileSync, writeFileSync, mkdtempSync, existsSync, readdirSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { Host } from '@/core/host'
import { Doc } from '@/core/doc'
import { CellError } from '@/core/engine/values'
import { parseCell } from '@/core/refs'
import { odfToExcel, excelToOdf } from '@/core/formats/ods/formula'

const fixture = (n: string) => new Uint8Array(readFileSync(path.join(import.meta.dirname, '../fixtures', n)))
const soffice = ['/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'].find((p) => existsSync(p))
function val(doc: Doc, ref: string) {
  const [s, a] = ref.includes('!') ? ref.split('!') : [doc.wb.sheets[0].name, ref]
  const p = parseCell(a)!
  const v = doc.wb.sheetByName(s)!.get(p.r, p.c)?.v ?? null
  return v instanceof CellError ? v.code : v
}
const fmla = (doc: Doc, ref: string) => {
  const [s, a] = ref.includes('!') ? ref.split('!') : [doc.wb.sheets[0].name, ref]
  const p = parseCell(a)!
  return doc.formulaText(doc.wb.sheetByName(s)!.id, p.r, p.c)
}

describe('ODS', () => {
  it('translates OpenFormula', () => {
    expect(odfToExcel('of:=SUM([.A1:.B2];[$Other.C3])')).toBe('=SUM(A1:B2,Other!C3)')
    expect(odfToExcel("of:=[$'My Sheet'.$A$1]*2")).toBe("='My Sheet'!$A$1*2")
    expect(odfToExcel('of:=IF([.A1]>0;"a;b";{1;2|3;4})')).toBe('=IF(A1>0,"a;b",{1,2;3,4})')
    expect(odfToExcel('of:=COM.MICROSOFT.XLOOKUP([.A1];[.B:.B];[.C:.C])')).toBe('=XLOOKUP(A1,B:B,C:C)')
    expect(odfToExcel('of:=LEGACY.NORMSINV(0.5)')).toBe('=NORMSINV(0.5)')
    expect(excelToOdf('=SUM(A1:B2,Other!C3)')).toBe('of:=SUM([.A1:.B2];[$Other.C3])')
    expect(excelToOdf("='My Sheet'!$A$1*2")).toBe("of:=[$'My Sheet'.$A$1]*2")
    expect(excelToOdf('=XLOOKUP(A1,B:B,C:C)')).toBe('of:=COM.MICROSOFT.XLOOKUP([.A1];[.B:.B];[.C:.C])')
    expect(excelToOdf('=IF(A1>0,"a,b",{1,2;3,4})')).toBe('of:=IF([.A1]>0;"a,b";{1;2|3;4})')
  })

  it('reads an ODS written by LibreOffice', async () => {
    const h = await Host.open(fixture('features.ods'), 'ods')
    const d = h.doc
    expect(d.wb.sheets.map((s) => s.name)).toEqual(['Data', 'Summary'])
    expect(val(d, 'A4')).toBe('Gizmo & Co "quoted"')
    expect(val(d, 'D2')).toBe(5002)
    expect(fmla(d, 'D2')).toBe('=B2*C2')
    expect(val(d, 'E4')).toBe(true)
    expect(val(d, 'Summary!B1')).toBeCloseTo(5203.4)
    expect(fmla(d, 'Summary!B1')).toBe('=SUM(Data!D2:D4)')
    expect(val(d, 'Summary!B2')).toBeCloseTo(199.9)
    expect(val(d, 'Summary!B3')).toBeCloseTo(423.66)
    const st = (ref: string) => d.styleAt(d.wb.sheets[0].id, parseCell(ref)!.r, parseCell(ref)!.c)
    expect(st('A1').bold).toBe(true)
    expect(st('C2').numFmt).toBe('#,##0.00')
    expect(d.display(d.wb.sheets[0].id, 1, 2)?.t).toBe('1,250.50')
    expect(st('D5').fill?.toLowerCase()).toBe('#ffff00')
    expect(st('A5').wrap).toBe(true)
    expect(d.display(d.wb.sheets[0].id, 1, 4)?.t).toBe('2024-01-15')
    expect(d.display(d.wb.sheets[0].id, 4, 2)?.t).toBe('96.1%')
    expect(d.wb.sheets[0].merges.length).toBe(1)
    expect(d.wb.sheets[0].colWidth(0)).toBe(144)
    expect(d.wb.names.map((n) => n.name)).toContain('Prices')
    expect(h.warnings.filter((w) => /differ from the values saved/.test(w.message))).toEqual([])
  })

  it('round-trips through ODS', async () => {
    const a = await Host.open(fixture('features.ods'), 'ods')
    a.doc.input(a.doc.wb.sheets[0].id, 1, 1, '6')
    a.doc.freeze(a.doc.wb.sheets[0].id, 1, 0)
    const b = await Host.open((await a.save('ods')).data, 'ods')
    const d = b.doc
    expect(val(d, 'D2')).toBe(7503)
    expect(val(d, 'E5')).toBe('many')
    expect(val(d, 'Summary!B1')).toBeCloseTo(7704.4)
    expect(d.display(d.wb.sheets[0].id, 1, 2)?.t).toBe('1,250.50')
    expect(d.display(d.wb.sheets[0].id, 1, 4)?.t).toBe('2024-01-15')
    expect(d.wb.sheets[0].merges.length).toBe(1)
    expect(d.wb.sheets[0].freeze).toEqual({ rows: 1, cols: 0 })
    expect(d.wb.names.map((n) => n.name)).toContain('Prices')
  })

  it('keeps shapes, pictures and styled empty cells', async () => {
    const { unzipSync, zipSync, strFromU8, strToU8 } = await import('fflate')
    const parts = unzipSync(fixture('features.ods'))
    let content = strFromU8(parts['content.xml'])
    const shapes = '<table:shapes><draw:frame draw:name="Logo" svg:width="1in" svg:height="1in" svg:x="0in" svg:y="0in"><draw:image xlink:href="Pictures/logo.png" xlink:type="simple"/></draw:frame></table:shapes>'
    content = content.replace(/(<table:table table:name="Data"[^>]*>)/, `$1${shapes}`)
    // A yellow empty cell in G2
    content = content.replace(/(<office:automatic-styles>)/, '$1<style:style style:name="ceY" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffff00"/></style:style>')
    content = content.replace(/(<table:table table:name="Data"[\s\S]*?<table:table-row[^>]*>[\s\S]*?<\/table:table-row>\s*<table:table-row[^>]*>)/, '$1<table:table-cell table:number-columns-repeated="6"/><table:table-cell table:style-name="ceY"/>')
    parts['content.xml'] = strToU8(content)
    parts['Pictures/logo.png'] = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const h = await Host.open(zipSync(parts), 'ods')
    const sh = h.doc.wb.sheets[0]
    expect(h.doc.styleAt(sh.id, 1, 6).fill?.toLowerCase()).toBe('#ffff00')
    expect(h.warnings.some((w) => /Charts, pictures and shapes are kept/.test(w.message))).toBe(true)
    const out = unzipSync((await h.save('ods')).data)
    expect(out['Pictures/logo.png']).toEqual(new Uint8Array([0x89, 0x50, 0x4e, 0x47]))
    expect(strFromU8(out['content.xml'])).toContain('draw:name="Logo"')
    expect(strFromU8(out['META-INF/manifest.xml'])).toContain('manifest:full-path="Pictures/logo.png"')
  })

  it.skipIf(!soffice)('LibreOffice reads the ODS Simple Cell writes, including newer Excel functions', async () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    doc.input(s, 0, 0, '1')
    doc.input(s, 1, 0, '2')
    doc.input(s, 0, 1, 'a')
    doc.input(s, 1, 1, 'b')
    doc.input(s, 2, 0, '=SUM(A1:A2)*2')
    doc.input(s, 3, 0, '=XLOOKUP(2,A1:A2,B1:B2)')
    doc.input(s, 4, 0, '=TEXTJOIN("+",TRUE,B1:B2)')
    doc.input(s, 5, 0, '12.5%')
    const out = (await new Host(doc).save('ods')).data
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-ods-'))
    writeFileSync(path.join(dir, 'ours.ods'), out)
    execFileSync(soffice!, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,1033,false,true,true,false,false,-1', '--outdir', dir, path.join(dir, 'ours.ods')], { timeout: 120_000 })
    const csv = readFileSync(path.join(dir, readdirSync(dir).find((f) => f.endsWith('.csv'))!), 'utf8').split('\n')
    expect(csv[2].split(',')[0]).toBe('6')
    expect(csv[3].split(',')[0]).toBe('b')
    expect(csv[4].split(',')[0]).toBe('a+b')
    expect(csv[5].split(',')[0]).toBe('12.5%')
  }, 150_000)
})
