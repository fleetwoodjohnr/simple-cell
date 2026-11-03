import { describe, expect, it } from 'vitest'
import { unzipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'
import { outlineRuns } from '@/renderer/app/Outline'

describe('grouping and outline', () => {
  it('groups, collapses, expands, nested levels, level buttons; XLSX keeps it', async () => {
    const h = new Host()
    const s = h.doc.wb.sheets[0].id
    for (let r = 0; r < 10; r++) h.doc.input(s, r, 0, String(r + 1))
    h.doc.input(s, 10, 0, '=SUBTOTAL(109,A1:A10)')
    h.group(s, 'row', 1, 8, 1) // rows 2–9
    h.group(s, 'row', 2, 4, 1) // rows 3–5, level 2
    const sh = h.doc.wb.sheets[0]
    expect(outlineRuns(h.meta().sheets[0].outlineRows)).toEqual([
      { level: 1, s: 1, e: 8 },
      { level: 2, s: 2, e: 4 }
    ])
    const hidden = () => [...sh.rowInfo].filter(([, i]) => i.hidden).map(([r]) => r).sort((a, b) => a - b)
    h.showDetail(s, 'row', 2, 4, false) // collapse the inner group
    expect(hidden()).toEqual([2, 3, 4])
    expect(sh.rowInfo.get(5)?.collapsed).toBe(true)
    expect(h.doc.value(s, 10, 0)).toBe(55 - 3 - 4 - 5) // SUBTOTAL skips hidden rows
    h.showDetail(s, 'row', 1, 8, false) // collapse the outer one
    expect(hidden()).toEqual([1, 2, 3, 4, 5, 6, 7, 8])
    h.showDetail(s, 'row', 1, 8, true) // expanding it keeps the inner group collapsed
    expect(hidden()).toEqual([2, 3, 4])
    h.showOutlineLevel(s, 'row', 3)
    expect(hidden()).toEqual([])
    h.showOutlineLevel(s, 'row', 1)
    expect(hidden()).toEqual([1, 2, 3, 4, 5, 6, 7, 8])
    const xml = strFromU8(unzipSync((await h.save('xlsx')).data)['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('outlineLevelRow="2"')
    expect(xml).toMatch(/<row r="4"[^>]*hidden="1"[^>]*outlineLevel="2"/)
    const back = await Host.open((await h.save('xlsx')).data, 'xlsx')
    expect(outlineRuns(back.meta().sheets[0].outlineRows)).toEqual(outlineRuns(h.meta().sheets[0].outlineRows))
    h.showOutlineLevel(s, 'row', 3)
    h.group(s, 'row', 1, 8, -1)
    expect(outlineRuns(h.meta().sheets[0].outlineRows)).toEqual([{ level: 1, s: 2, e: 4 }])
    h.group(s, 'row', 2, 4, -1)
    expect(h.meta().sheets[0].outlineRows).toEqual([])
  })
})
