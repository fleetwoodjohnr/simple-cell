import { describe, expect, it } from 'vitest'
import { book } from './helpers/book'
import { parseArea } from '@/core/refs'
import type { CfSpec } from '@/core/doc'
import { Host } from '@/core/host'

const RED = { fill: '#ffc7ce', color: '#9c0006' }

function setup(values: (string | number)[], spec: CfSpec, range = `A1:A${values.length}`) {
  const b = book()
  values.forEach((v, i) => v !== '' && b.set(`A${i + 1}`, v))
  b.doc.addCondFormat(b.sheet, [parseArea(range, b.sheet)!], spec)
  const at = (ref: string) => {
    const a = parseArea(ref, b.sheet)!
    return b.doc.condFormatsIn(b.sheet, a).get(a.r0 * 16384 + a.c0)
  }
  const hit = (ref: string) => !!at(ref)?.style
  return { b, at, hit }
}

describe('conditional formatting', () => {
  it('cell value rules (blank cells count as 0, as in Excel)', () => {
    const { hit } = setup([3, 7, 5, ''], { type: 'cellIs', operator: 'greaterThan', formulas: ['5'], dxf: RED })
    expect([hit('A1'), hit('A2'), hit('A3'), hit('A4')]).toEqual([false, true, false, false])
    const lt = setup([3, 7, '', 'x'], { type: 'cellIs', operator: 'lessThan', formulas: ['5'], dxf: RED })
    expect([lt.hit('A1'), lt.hit('A2'), lt.hit('A3'), lt.hit('A4')]).toEqual([true, false, true, false])
    const bw = setup([1, 5, 10, 11], { type: 'cellIs', operator: 'between', formulas: ['10', '5'], dxf: RED })
    expect([bw.hit('A1'), bw.hit('A2'), bw.hit('A3'), bw.hit('A4')]).toEqual([false, true, true, false])
    const eq = setup(['Apple', 'pear'], { type: 'cellIs', operator: 'equal', formulas: ['"apple"'], dxf: RED })
    expect([eq.hit('A1'), eq.hit('A2')]).toEqual([true, false])
  })

  it('formula rules are relative to the first cell of the range', () => {
    const b = book({ A1: 1, B1: 'x', A2: 2, B2: 'y', A3: 3, B3: 'x' })
    b.doc.addCondFormat(b.sheet, [parseArea('A1:C3', b.sheet)!], { type: 'expression', formulas: ['=$B1="x"'], dxf: RED })
    const cf = b.doc.condFormatsIn(b.sheet, parseArea('A1:C3', b.sheet)!)
    // Whole rows 1 and 3 (including the empty C cells), not row 2.
    expect([...cf.keys()].sort((x, y) => x - y)).toEqual([0, 1, 2, 2 * 16384, 2 * 16384 + 1, 2 * 16384 + 2])
    expect(cf.get(0)?.style).toEqual(RED)
  })

  it('updates when the cells it depends on change', () => {
    const { b, hit } = setup([3, 7], { type: 'cellIs', operator: 'greaterThan', formulas: ['$C$1'], dxf: RED })
    b.set('C1', 5)
    expect(hit('A1')).toBe(false)
    b.set('C1', 1)
    expect(hit('A1')).toBe(true)
  })

  it('top/bottom, average, duplicates', () => {
    const nums = [10, 50, 20, 40, 30, 50]
    const top = setup(nums, { type: 'top10', rank: 2, dxf: RED })
    expect(nums.map((_, i) => top.hit(`A${i + 1}`))).toEqual([false, true, false, false, false, true]) // ties count
    const bottom = setup(nums, { type: 'top10', rank: 34, percent: true, bottom: true, dxf: RED })
    expect(nums.map((_, i) => bottom.hit(`A${i + 1}`))).toEqual([true, false, true, false, false, false]) // 34% of 6 → 2
    const avg = setup(nums, { type: 'aboveAverage', dxf: RED }) // average 33.33
    expect(nums.map((_, i) => avg.hit(`A${i + 1}`))).toEqual([false, true, false, true, false, true])
    const below = setup(nums, { type: 'aboveAverage', aboveAverage: false, dxf: RED })
    expect(nums.map((_, i) => below.hit(`A${i + 1}`))).toEqual([true, false, true, false, true, false])
    const dup = setup(['a', 'A', 'b', 3, 3, ''], { type: 'duplicateValues', dxf: RED })
    expect([1, 2, 3, 4, 5, 6].map((i) => dup.hit(`A${i}`))).toEqual([true, true, false, true, true, false])
    const uniq = setup(['a', 'A', 'b'], { type: 'uniqueValues', dxf: RED })
    expect([uniq.hit('A1'), uniq.hit('A3')]).toEqual([false, true])
  })

  it('text, blank and error rules', () => {
    const t = setup(['Hello world', 'say hello', 'bye'], { type: 'containsText', text: 'HELLO', formulas: ['NOT(ISERROR(SEARCH("HELLO",A1)))'], dxf: RED })
    expect([t.hit('A1'), t.hit('A2'), t.hit('A3')]).toEqual([true, true, false])
    const begins = setup(['Hello world', 'say hello'], { type: 'beginsWith', text: 'hel', dxf: RED })
    expect([begins.hit('A1'), begins.hit('A2')]).toEqual([true, false])
    const blanks = setup(['x', '', 'y'], { type: 'containsBlanks', dxf: RED })
    expect([blanks.hit('A1'), blanks.hit('A2')]).toEqual([false, true])
    const errs = setup(['=1/0', 2], { type: 'containsErrors', dxf: RED })
    expect([errs.hit('A1'), errs.hit('A2')]).toEqual([true, false])
  })

  it('priority and stop if true', () => {
    const { b, at } = setup([7], { type: 'cellIs', operator: 'greaterThan', formulas: ['1'], dxf: { fill: '#00ff00', bold: true } })
    b.doc.addCondFormat(b.sheet, [parseArea('A1', b.sheet)!], { type: 'cellIs', operator: 'greaterThan', formulas: ['5'], dxf: { fill: '#ff0000' } })
    // The newer rule has priority: its fill wins; bold from the other rule still applies.
    expect(at('A1')?.style).toEqual({ fill: '#ff0000', bold: true })
    const [first, second] = b.doc.listCondFormats(b.sheet)
    expect(first.formulas).toEqual(['=5'])
    b.doc.moveCondRule(b.sheet, first.id, 1)
    expect(at('A1')?.style).toEqual({ fill: '#00ff00', bold: true })
    b.doc.deleteCondRule(b.sheet, second.id)
    expect(at('A1')?.style).toEqual({ fill: '#ff0000' })
    b.doc.undo()
    expect(at('A1')?.style).toEqual({ fill: '#00ff00', bold: true })
  })

  it('colour scales, data bars, icon sets', () => {
    const vals = [0, 50, 100, 25]
    const cs = setup(vals, { type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['#ff0000', '#00ff00'] } })
    expect(cs.at('A1')?.style?.fill).toBe('#ff0000')
    expect(cs.at('A2')?.style?.fill).toBe('#808000')
    expect(cs.at('A3')?.style?.fill).toBe('#00ff00')
    const three = setup(vals, { type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }], colors: ['#f8696b', '#ffeb84', '#63be7b'] } })
    expect(three.at('A2')?.style?.fill).toBe('#e0e282') // 20% of the way from the midpoint (37.5) to the max
    const db = setup(vals, { type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '#638ec6' } })
    expect(db.at('A1')?.bar).toEqual({ w: 0.1, color: '#638ec6' })
    expect(db.at('A3')?.bar?.w).toBeCloseTo(0.9)
    expect(db.at('A2')?.bar?.w).toBeCloseTo(0.5)
    const ic = setup(vals, { type: 'iconSet', iconSet: { set: '3Arrows', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }] } })
    expect(vals.map((_, i) => ic.at(`A${i + 1}`)?.icon?.i)).toEqual([0, 1, 2, 0])
    const rev = setup(vals, { type: 'iconSet', iconSet: { set: '3Arrows', reverse: true, showValue: false, cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }] } })
    expect(rev.at('A1')?.icon?.i).toBe(2)
    expect(rev.at('A1')?.hide).toBe(true)
  })

  it('follows inserted and deleted rows; clearing from a selection splits ranges', () => {
    const { b } = setup([1, 2, 3, 4], { type: 'expression', formulas: ['=A1>$C$1'], dxf: RED })
    b.set('C1', 2)
    b.doc.insertRows(b.sheet, 0, 1)
    const [r] = b.doc.listCondFormats(b.sheet)
    expect(r.ranges).toBe('A2:A5')
    expect(r.formulas).toEqual(['=A2>$C$2'])
    expect(b.doc.condFormatsIn(b.sheet, parseArea('A1:A5', b.sheet)!).size).toBe(2) // 3 and 4
    b.doc.clearCondFormats(b.sheet, [parseArea('A3:A4', b.sheet)!])
    expect(b.doc.listCondFormats(b.sheet)[0].ranges).toBe('A2 A5')
    b.doc.clearCondFormats(b.sheet, null)
    expect(b.doc.listCondFormats(b.sheet)).toEqual([])
  })

  it('the viewport includes empty coloured cells and the rule format', () => {
    const h = new Host()
    const s = h.doc.wb.sheets[0].id
    h.doc.input(s, 0, 0, '9')
    h.addCondFormat(s, [parseArea('A1:B2', s)!], { type: 'expression', formulas: ['=TRUE'], dxf: { fill: '#ffff00', numFmt: '0.00' } })
    const cells = h.viewport(s, 0, 5, 0, 5)
    expect(cells.length).toBe(4)
    const a1 = cells.find((c) => c.r === 0 && c.c === 0)!
    expect(a1.t).toBe('9.00')
    expect(a1.x?.fill).toBe('#ffff00')
  })
})
