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
  ;[['3', '5', '2', '8'], ['-1', '4', '-2', '6']].forEach((row, r) => row.forEach((v, c) => h.doc.input(s, r, c, v)))
  return { h, s }
}

describe('sparklines', () => {
  it('one per location cell, values in the viewport, follow inserted rows, clear', () => {
    const { h, s } = book()
    expect(h.addSparklines(s, 'line', { sheet: s, r0: 0, c0: 0, r1: 1, c1: 3 }, { sheet: s, r0: 0, c0: 4, r1: 1, c1: 4 }).ok).toBe(true)
    const g = h.doc.wb.sheets[0].sparklines[0]
    expect(g.items).toEqual([
      { r: 0, c: 4, ref: 'Sheet1!$A$1:$D$1' },
      { r: 1, c: 4, ref: 'Sheet1!$A$2:$D$2' }
    ])
    const v = h.viewport(s, 0, 1, 4, 4)
    expect(v.find((x) => x.r === 1)?.spark).toMatchObject({ type: 'line', values: [-1, 4, -2, 6] })
    h.doc.input(s, 1, 0, '9')
    expect(h.viewport(s, 1, 1, 4, 4)[0].spark?.values[0]).toBe(9)
    h.doc.insertRows(s, 0, 1)
    expect(h.doc.wb.sheets[0].sparklines[0].items[0]).toEqual({ r: 1, c: 4, ref: 'Sheet1!$A$2:$D$2' })
    expect(h.addSparklines(s, 'column', { sheet: s, r0: 1, c0: 0, r1: 2, c1: 3 }, { sheet: s, r0: 1, c0: 5, r1: 3, c1: 5 }).ok).toBe(false) // 3 cells, 2 rows
    h.clearSparklines(s, [{ sheet: s, r0: 1, c0: 4, r1: 1, c1: 4 }])
    expect(h.doc.wb.sheets[0].sparklines[0].items.length).toBe(1)
  })

  it('XLSX: written into the extension list next to other extensions, and read back', async () => {
    const { h, s } = book()
    h.addSparklines(s, 'stacked', { sheet: s, r0: 0, c0: 0, r1: 1, c1: 3 }, { sheet: s, r0: 0, c0: 4, r1: 1, c1: 4 })
    const files = unzipSync((await h.save('xlsx')).data)
    let xml = strFromU8(files['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('<x14:sparklineGroup type="stacked" negative="1">')
    expect(xml).toContain('<x14:sparkline><xm:f>Sheet1!$A$2:$D$2</xm:f><xm:sqref>E2</xm:sqref></x14:sparkline>')
    // Another extension (as Excel writes for data bars) must survive a round trip.
    xml = xml.replace('</extLst>', '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings/></ext></extLst>')
    files['xl/worksheets/sheet1.xml'] = strToU8(xml)
    const back = await Host.open(zipSync(files), 'xlsx')
    expect(back.doc.wb.sheets[0].sparklines).toEqual(h.doc.wb.sheets[0].sparklines)
    back.clearSparklines(back.doc.wb.sheets[0].id, [{ sheet: s, r0: 0, c0: 4, r1: 0, c1: 4 }])
    const again = strFromU8(unzipSync((await back.save('xlsx')).data)['xl/worksheets/sheet1.xml'])
    expect(again).toContain('{78C0D931-6437-407d-A8EE-F0AAD7539E65}')
    expect(again).not.toContain('<xm:sqref>E1</xm:sqref>')
    expect(again).toContain('<xm:sqref>E2</xm:sqref>')
    if (!soffice) return
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-spark-'))
    writeFileSync(path.join(dir, 's.xlsx'), (await h.save('xlsx')).data)
    execFileSync(soffice, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'ods', '--outdir', dir, path.join(dir, 's.xlsx')], { timeout: 120_000 })
    const content = strFromU8(unzipSync(new Uint8Array(readFileSync(path.join(dir, readdirSync(dir).find((f) => f.endsWith('.ods'))!))))['content.xml'])
    expect(content).toContain('calcext:sparkline-group')
    expect(content).toContain('calcext:type="stacked"')
  }, 150_000)
})
