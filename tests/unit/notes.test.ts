import { describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { unzipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'

const soffice = ['/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'].find((p) => existsSync(p))

function sample() {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  h.doc.input(s, 0, 0, 'Budget')
  h.setNote(s, 0, 0, 'Checked by finance\nQ3 figures', 'Ann')
  h.setNote(s, 2, 1, 'On an empty cell', 'Bob')
  h.setLink([{ sheet: s, r0: 1, c0: 0, r1: 1, c1: 0 }], { url: 'https://example.com/report?a=1&b=2', tooltip: 'Report' }, 'The report')
  h.setLink([{ sheet: s, r0: 3, c0: 0, r1: 3, c1: 0 }], { location: 'Sheet1!C10' })
  return { h, s }
}

describe('notes and hyperlinks', () => {
  it('edit, show in the viewport, follow structure edits, Clear All, undo', () => {
    const { h, s } = sample()
    expect(h.noteAt(s, 0, 0)).toEqual({ author: 'Ann', text: 'Checked by finance\nQ3 figures' })
    expect(h.doc.value(s, 1, 0)).toBe('The report')
    expect(h.doc.styleAt(s, 1, 0)).toMatchObject({ color: '#0563C1', underline: 'single' })
    expect(h.doc.value(s, 3, 0)).toBe('Sheet1!C10') // an empty cell shows the address
    const cells = h.viewport(s, 0, 5, 0, 3)
    expect(cells.find((c) => c.r === 2 && c.c === 1)).toMatchObject({ note: 1, t: '' })
    expect(cells.find((c) => c.r === 1 && c.c === 0)?.link).toBe(1)
    h.doc.insertRows(s, 0, 2)
    expect(h.noteAt(s, 2, 0)?.author).toBe('Ann')
    expect(h.linkAt(s, 3, 0)?.url).toBe('https://example.com/report?a=1&b=2')
    h.doc.deleteRows(s, 4, 1) // the row with the note on B3 (now B5)
    expect(h.noteAt(s, 4, 1)).toBeNull()
    h.doc.undo()
    expect(h.noteAt(s, 4, 1)?.text).toBe('On an empty cell')
    h.doc.clear([{ sheet: s, r0: 2, c0: 0, r1: 3, c1: 0 }], 'contents')
    expect(h.noteAt(s, 2, 0)).not.toBeNull() // contents only
    h.doc.clear([{ sheet: s, r0: 2, c0: 0, r1: 3, c1: 0 }], 'all')
    expect(h.noteAt(s, 2, 0)).toBeNull()
    expect(h.linkAt(s, 3, 0)).toBeNull()
    h.setNote(s, 4, 1, null)
    expect(h.noteAt(s, 4, 1)).toBeNull()
  })

  it('XLSX: comments part, VML shapes and hyperlinks are written and read back', async () => {
    const { h, s } = sample()
    const out = await h.save('xlsx')
    const files = unzipSync(out.data)
    const comments = strFromU8(files['xl/comments1.xml'])
    expect(comments).toContain('<authors><author>Ann</author><author>Bob</author></authors>')
    expect(comments).toContain('<comment ref="A1" authorId="0">')
    const vml = strFromU8(files['xl/drawings/vmlDrawing1.vml'])
    expect((vml.match(/ObjectType="Note"/g) ?? []).length).toBe(2)
    expect(vml).toContain('<x:Row>2</x:Row><x:Column>1</x:Column>')
    const sheet = strFromU8(files['xl/worksheets/sheet1.xml'])
    expect(sheet).toMatch(/<hyperlinks><hyperlink ref="A2" r:id="(rId\d+)" tooltip="Report"\/><hyperlink ref="A4" location="Sheet1!C10"\/><\/hyperlinks>.*<legacyDrawing r:id="rId\d+"\/>/)
    const rels = strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels'])
    expect(rels).toContain('Target="https://example.com/report?a=1&amp;b=2" TargetMode="External"')
    expect(strFromU8(files['[Content_Types].xml'])).toContain('<Default Extension="vml"')
    const back = await Host.open(out.data, 'xlsx')
    const bs = back.doc.wb.sheets[0].id
    expect(back.noteAt(bs, 0, 0)).toMatchObject({ author: 'Ann', text: 'Checked by finance\nQ3 figures' })
    expect(back.linkAt(bs, 1, 0)).toEqual({ url: 'https://example.com/report?a=1&b=2', tooltip: 'Report' })
    expect(back.linkAt(bs, 3, 0)).toEqual({ location: 'Sheet1!C10' })
    // Unchanged notes keep their parts byte for byte; deleting the last note removes them.
    const again = unzipSync((await back.save('xlsx')).data)
    expect(strFromU8(again['xl/comments1.xml'])).toBe(comments)
    back.setNote(bs, 0, 0, null)
    back.setNote(bs, 2, 1, null)
    const none = unzipSync((await back.save('xlsx')).data)
    expect(Object.keys(none).some((p) => /comments|vml/.test(p))).toBe(false)
    expect(strFromU8(none['xl/worksheets/sheet1.xml'])).not.toContain('legacyDrawing')
    if (!soffice) return
    const dir = mkdtempSync(path.join(tmpdir(), 'sc-notes-'))
    writeFileSync(path.join(dir, 'n.xlsx'), out.data)
    execFileSync(soffice, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'ods', '--outdir', dir, path.join(dir, 'n.xlsx')], { timeout: 120_000 })
    const content = strFromU8(unzipSync(new Uint8Array(readFileSync(path.join(dir, 'n.ods'))))['content.xml'])
    expect(content).toContain('<dc:creator>Ann</dc:creator>')
    expect(content).toContain('Checked by finance')
    expect(content).toContain('xlink:href="https://example.com/report?a=1&amp;b=2"')
  }, 150_000)

  it('ODS: annotations and links round-trip', async () => {
    const { h } = sample()
    const back = await Host.open((await h.save('ods')).data, 'ods')
    const bs = back.doc.wb.sheets[0].id
    expect(back.noteAt(bs, 0, 0)).toMatchObject({ author: 'Ann', text: 'Checked by finance\nQ3 figures' })
    expect(back.noteAt(bs, 2, 1)?.text).toBe('On an empty cell')
    expect(back.linkAt(bs, 1, 0)?.url).toBe('https://example.com/report?a=1&b=2')
    expect(back.linkAt(bs, 3, 0)?.location).toBe('Sheet1!C10')
    expect(back.doc.value(bs, 1, 0)).toBe('The report')
  })
})
