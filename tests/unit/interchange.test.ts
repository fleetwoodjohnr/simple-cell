import { describe, expect, it } from 'vitest'
import { Doc } from '@/core/doc'
import { Host } from '@/core/host'
import { importWorkbook, exportWorkbook } from '@/core/formats/registry'
import { importDelimited } from '@/core/formats/csv'
import { r1c1Formula } from '@/core/formats/xml2003'

const enc = (s: string) => new TextEncoder().encode(s)
const dec = (b: Uint8Array) => new TextDecoder().decode(b)

function sample() {
  const doc = new Doc()
  const s = doc.wb.sheets[0].id
  doc.input(s, 0, 0, 'Name')
  doc.input(s, 0, 1, 'Amount')
  doc.input(s, 0, 2, 'When')
  doc.input(s, 1, 0, 'Ann')
  doc.input(s, 1, 1, '12.5')
  doc.input(s, 1, 2, '2024-03-01')
  doc.input(s, 2, 0, 'Bob <b>')
  doc.input(s, 2, 1, '=B2*2')
  doc.input(s, 2, 2, '3/2/2024')
  doc.format([{ sheet: s, r0: 0, c0: 0, r1: 0, c1: 2 }], { bold: true })
  return doc
}

describe('interchange formats', () => {
  it('JSON: array of objects out and back', async () => {
    const out = await exportWorkbook('json', sample().wb)
    const data = JSON.parse(dec(out.data))
    expect(data).toEqual([
      { Name: 'Ann', Amount: 12.5, When: '2024-03-01' },
      { Name: 'Bob <b>', Amount: 25, When: '2024-03-02' }
    ])
    const back = new Doc((await importWorkbook('json', out.data)).workbook)
    const s = back.wb.sheets[0].id
    expect(back.value(s, 2, 1)).toBe(25)
    expect(back.display(s, 1, 2)?.t).toBe('2024-03-01')
    const arrays = new Doc((await importWorkbook('json', enc('{"rows":[[1,"x"],[true,null]]}'))).workbook)
    expect(arrays.value(arrays.wb.sheets[0].id, 1, 0)).toBe(true)
  })

  it('HTML: tables with spans in, formatted table out', async () => {
    const r = await importWorkbook('html', enc('<html><body><table><tr><th>A</th><th colspan="2">B &amp; C</th></tr><tr><td>1,000</td><td>x</td><td>2024-01-05</td></tr></table></body></html>'))
    const d = new Doc(r.workbook)
    const s = d.wb.sheets[0].id
    expect(d.value(s, 0, 1)).toBe('B & C')
    expect(d.wb.sheets[0].merges).toEqual([{ sheet: s, r0: 0, c0: 1, r1: 0, c1: 2 }])
    expect(d.value(s, 1, 0)).toBe(1000)
    expect(d.styleAt(s, 0, 0).bold).toBe(true)
    const out = dec((await exportWorkbook('html', sample().wb)).data)
    expect(out).toContain('<td style="font-weight:bold">Name</td>')
    expect(out).toContain('Bob &lt;b&gt;')
    expect(out).toContain('>25</td>')
  })

  it('XML Spreadsheet 2003 with R1C1 formulas', async () => {
    expect(r1c1Formula('=R[-1]C*2+SUM(R1C1:R2C[1])', 2, 1)).toBe('=B2*2+SUM($A$1:C$2)')
    expect(r1c1Formula('="R1C1"&RC[1]', 0, 0)).toBe('="R1C1"&B1')
    const out = await exportWorkbook('xml', sample().wb)
    const x = dec(out.data)
    expect(x).toContain('ss:Formula="=R[-1]C*2"')
    const back = new Doc((await importWorkbook('xml', out.data)).workbook)
    const s = back.wb.sheets[0].id
    expect(back.formulaText(s, 2, 1)).toBe('=B2*2')
    expect(back.value(s, 2, 1)).toBe(25)
    expect(back.styleAt(s, 0, 0).bold).toBe(true)
    expect(back.display(s, 2, 2)?.t).toBe('3/2/2024')
    // Names and frozen panes
    const doc = sample()
    doc.freeze(doc.wb.sheets[0].id, 1, 0)
    doc.defineName('Amounts', '=Sheet1!$B$2:$B$3')
    const again = new Doc((await importWorkbook('xml', (await exportWorkbook('xml', doc.wb)).data)).workbook)
    expect(again.wb.sheets[0].freeze).toEqual({ rows: 1, cols: 0 })
    again.input(again.wb.sheets[0].id, 5, 0, '=SUM(Amounts)')
    expect(again.value(again.wb.sheets[0].id, 5, 0)).toBe(37.5)
  })

  it('CSV: encodings, delimiters and quoting', () => {
    const latin1 = new Uint8Array([0x63, 0x61, 0x66, 0xe9, 0x3b, 0x31, 0x0a]) // "café;1" in Windows-1252
    const wb = importDelimited(latin1).workbook
    expect(wb.sheets[0].get(0, 0)?.v).toBe('café')
    expect(wb.sheets[0].get(0, 1)?.v).toBe(1)
    const q = importDelimited(enc('a,"b, ""c""\nd",3\n')).workbook
    expect(q.sheets[0].get(0, 1)?.v).toBe('b, "c"\nd')
  })

  it('opens every format through the host', async () => {
    for (const f of ['json', 'html', 'xml', 'xlsx', 'ods', 'csv'] as const) {
      const out = await exportWorkbook(f, sample().wb)
      const h = await Host.open(out.data, f)
      expect(h.doc.value(h.doc.wb.sheets[0].id, 1, 0), f).toBe('Ann')
    }
  })
})
