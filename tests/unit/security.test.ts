import { describe, expect, it } from 'vitest'
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'
import { Host } from '@/core/host'
import { Doc } from '@/core/doc'
import { exportDelimited } from '@/core/formats/csv'
import { isSafeExternalUrl } from '@/shared/urls'
import { parseXml } from '@/core/formats/xml'

async function blank() {
  return unzipSync((await new Host().save('xlsx')).data)
}

describe('security', () => {
  it('XML: external entities are never resolved; entity expansion bombs do nothing', () => {
    const xxe = '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><x>&e;</x>'
    expect(parseXml(xxe).documentElement.textContent).not.toMatch(/root:/)
    const laughs = '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;"><!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">]><lolz>&lol3;</lolz>'
    const t = performance.now()
    expect((parseXml(laughs).documentElement.textContent ?? '').length).toBeLessThan(100)
    expect(performance.now() - t).toBeLessThan(1000)
  })

  it('XLSX: a DOCTYPE in a part is ignored; absurd row numbers and huge spans are clamped', async () => {
    const files = await blank()
    files['xl/sharedStrings.xml'] = strToU8('<?xml version="1.0"?><!DOCTYPE sst [<!ENTITY e SYSTEM "file:///etc/passwd">]><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>&e;safe</t></si></sst>')
    const ws = strFromU8(files['xl/worksheets/sheet1.xml'])
    files['xl/worksheets/sheet1.xml'] = strToU8(ws.replace(/<sheetData\/>|<sheetData>[\s\S]*<\/sheetData>/, '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="99999999"><c r="XFD99999999"><v>1</v></c></row></sheetData>').replace('</cols>', '</cols>').replace(/<cols>[\s\S]*?<\/cols>/, '') .replace('<sheetData>', '<cols><col min="1" max="99999999" width="5"/></cols><sheetData>'))
    files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'))
    const rels = strFromU8(files['xl/_rels/workbook.xml.rels'])
    if (!rels.includes('sharedStrings')) files['xl/_rels/workbook.xml.rels'] = strToU8(rels.replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'))
    const h = await Host.open(zipSync(files), 'xlsx')
    const v = h.doc.value(h.doc.wb.sheets[0].id, 0, 0)
    expect(String(v)).not.toMatch(/root:/)
    expect(h.doc.wb.sheets[0].used().r).toBeLessThan(1_048_576)
  })

  it('CSV export guards against formula injection by default', () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    doc.input(s, 0, 0, "'=HYPERLINK(\"http://evil\",\"x\")")
    doc.input(s, 1, 0, "'+1+1")
    doc.input(s, 2, 0, "'@SUM(1)")
    doc.input(s, 3, 0, '-5')
    const out = new TextDecoder().decode(exportDelimited(doc.wb, ',', { guard: true }).data)
    expect(out.split(/\r?\n/).slice(0, 4)).toEqual(['"\'=HYPERLINK(""http://evil"",""x"")"', "'+1+1", "'@SUM(1)", '-5'])
  })

  it('only web and mail links are opened outside the app', () => {
    for (const ok of ['https://example.com/a?b=1', 'http://x.org', 'mailto:a@b.com']) expect(isSafeExternalUrl(ok), ok).toBe(true)
    for (const bad of ['javascript:alert(1)', 'file:///etc/passwd', 'data:text/html,x', 'vbscript:x', 'ms-msdt:/id', 'smb://host/share', 'https://ex\u0000ample.com', '\u0001https://x', 42, null, 'not a url']) expect(isSafeExternalUrl(bad), String(bad)).toBe(false)
  })

  it('deeply nested formulas fail cleanly instead of crashing', () => {
    const doc = new Doc()
    const s = doc.wb.sheets[0].id
    expect(() => doc.input(s, 0, 0, '=' + '('.repeat(5000) + '1' + ')'.repeat(5000))).not.toThrow(RangeError)
    const deep = '=' + 'ABS('.repeat(3000) + '1' + ')'.repeat(3000)
    try {
      doc.input(s, 1, 0, deep)
    } catch (e) {
      expect(e).not.toBeInstanceOf(RangeError)
    }
  })
})
