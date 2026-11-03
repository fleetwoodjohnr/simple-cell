import { describe, expect, it } from 'vitest'
import { unzipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'
import { headerFooterText } from '@/core/print'

function grid(rows: number, cols: number) {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) h.doc.input(s, r, c, String(r * cols + c))
  return { h, s }
}

describe('printing', () => {
  it('paginates down then over, with breaks, titles and fit to page', () => {
    const { h, s } = grid(100, 14) // 100 rows × 20 px, 14 cols × 64 px
    let job = h.printJob(s)!
    // Letter portrait: 9.5 in × 96 = 912 px tall → 45 rows; 7.1 in = 681.6 px wide → 10 columns.
    expect(job.pages.map((p) => [p.rows[0], p.rows.length, p.cols[0], p.cols.length])).toEqual([
      [0, 45, 0, 10],
      [45, 45, 0, 10],
      [90, 10, 0, 10],
      [0, 45, 10, 4],
      [45, 45, 10, 4],
      [90, 10, 10, 4]
    ])
    h.setPageSetup(s, { order: 'overThenDown' })
    expect(h.printJob(s)!.pages.slice(0, 2).map((p) => p.cols[0])).toEqual([0, 10])
    h.setPageSetup(s, { order: 'downThenOver', rowBreaks: [20] })
    expect(h.printJob(s)!.pages.slice(0, 2).map((p) => p.rows.length)).toEqual([20, 45])
    h.setPageSetup(s, { rowBreaks: [] })
    h.setPrintName(s, 'titles', '$1:$1')
    job = h.printJob(s)!
    expect(job.pages[1].titleRows).toEqual([0])
    expect(job.pages[1].rows.length).toBe(44) // room for the repeated row
    expect(job.pages[0].titleRows).toEqual([])
    h.setPageSetup(s, { fitToPage: true, fitWidth: 1, fitHeight: 1 })
    job = h.printJob(s)!
    expect(job.pages.length).toBe(1)
    expect(job.scale).toBeCloseTo(912 / 2000)
    h.setPageSetup(s, { fitToPage: false, orientation: 'landscape' })
    expect(h.printJob(s)!.pages[0].cols.length).toBe(14) // 9.6 in wide
    h.setPrintName(s, 'area', '$B$2:$C$3')
    expect(h.printJob(s)!.pages.map((p) => [p.rows, p.cols])).toEqual([[[1, 2], [1, 2]]])
  })

  it('header and footer codes', () => {
    const ctx = { page: 2, pages: 5, sheet: 'Sales', file: 'Q3.xlsx', date: new Date(2024, 0, 15) }
    expect(headerFooterText('&L&"Arial,Bold"&14&A&CPage &P of &N&R&F', ctx)).toEqual({ l: 'Sales', c: 'Page 2 of 5', r: 'Q3.xlsx' })
    expect(headerFooterText('Tom && Jerry', ctx).c).toBe('Tom & Jerry')
  })

  it('XLSX keeps page setup, breaks, titles and print area', async () => {
    const { h, s } = grid(5, 3)
    h.setPageSetup(s, { orientation: 'landscape', paper: 9, fitToPage: true, fitWidth: 1, fitHeight: 0, gridlines: true, header: '&C&A', footer: 'Page &P', rowBreaks: [3], margins: { left: 0.5 } as never })
    h.setPrintName(s, 'area', '$A$1:$C$5')
    h.setPrintName(s, 'titles', '$1:$1')
    const files = unzipSync((await h.save('xlsx')).data)
    const xml = strFromU8(files['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
    expect(xml).toContain('<printOptions gridLines="1"/><pageMargins left="0.5" right="0.7"')
    expect(xml).toContain('<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" orientation="landscape"/>')
    expect(xml).toContain('<headerFooter><oddHeader>&amp;C&amp;A</oddHeader><oddFooter>Page &amp;P</oddFooter></headerFooter><rowBreaks count="1" manualBreakCount="1"><brk id="3" max="16383" man="1"/></rowBreaks>')
    const wbx = strFromU8(files['xl/workbook.xml'])
    expect(wbx).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$C$5</definedName>')
    expect(wbx).toContain('<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName>')
    const back = await Host.open((await h.save('xlsx')).data, 'xlsx')
    const ps = back.pageSetup(back.doc.wb.sheets[0].id)
    expect(ps).toMatchObject({ orientation: 'landscape', paper: 9, fitToPage: true, fitWidth: 1, fitHeight: 0, gridlines: true, header: '&C&A', footer: 'Page &P', rowBreaks: [3], area: { r0: 0, c0: 0, r1: 4, c1: 2 }, titles: { rows: [0, 0] } })
    expect(ps.margins.left).toBe(0.5)
  })
})
