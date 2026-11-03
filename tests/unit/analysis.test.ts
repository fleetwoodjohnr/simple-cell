import { describe, expect, it } from 'vitest'
import { book } from './helpers/book'
import { descriptiveStats, goalSeek, removeDuplicates, textToColumns } from '@/core/analysis'

describe('data tools', () => {
  it('Remove Duplicates keeps the first of each, by the chosen columns, under the header', () => {
    const b = book({ A1: 'Name', B1: 'City', A2: 'Ann', B2: 'Oslo', A3: 'Bo', B3: 'Rome', A4: 'ann', B4: 'Oslo', A5: 'Bo', B5: 'Paris', C5: 'outside' })
    const s = b.sheet
    expect(removeDuplicates(b.doc, { sheet: s, r0: 0, c0: 0, r1: 4, c1: 1 }, [0, 1], true)).toEqual({ removed: 1, kept: 3 })
    expect([b.get('A2'), b.get('A3'), b.get('A4'), b.get('B4'), b.get('A5')]).toEqual(['Ann', 'Bo', 'Bo', 'Paris', null])
    expect(b.get('C5')).toBe('outside')
    expect(removeDuplicates(b.doc, { sheet: s, r0: 0, c0: 0, r1: 3, c1: 1 }, [0], true)).toEqual({ removed: 1, kept: 2 })
    b.doc.undo()
    expect(b.get('A4')).toBe('Bo')
  })

  it('Text to Columns splits by delimiters, with quotes and consecutive delimiters', () => {
    const b = book({ A1: 'Lee, Ann, 42', A2: '"Smith, Jr", Bo,, 7', A3: 5 })
    expect(textToColumns(b.doc, { sheet: b.sheet, r0: 0, c0: 0, r1: 2, c1: 0 }, [','], { consecutive: true })).toBe(3)
    expect([b.get('A1'), b.get('B1'), b.get('C1')]).toEqual(['Lee', ' Ann', 42])
    expect([b.get('A2'), b.get('B2'), b.get('C2')]).toEqual(['Smith, Jr', ' Bo', 7])
    expect(b.get('A3')).toBe(5)
  })

  it('Goal Seek finds the input that makes a formula reach a value', () => {
    const b = book({ A1: 1000, B1: '=A1*1.08-50' })
    const r = goalSeek(b.doc, { sheet: b.sheet, r: 0, c: 1 }, 2000, { sheet: b.sheet, r: 0, c: 0 })
    expect(r.ok).toBe(true)
    expect(b.get('A1')).toBeCloseTo(2050 / 1.08, 6)
    expect(b.get('B1')).toBeCloseTo(2000, 6)
    const q = book({ A1: 1, B1: '=A1^2' }) // nonlinear
    expect(goalSeek(q.doc, { sheet: q.sheet, r: 0, c: 1 }, 2, { sheet: q.sheet, r: 0, c: 0 }).value).toBeCloseTo(Math.SQRT2, 6)
    b.doc.undo()
    expect(b.get('A1')).toBe(1000)
    expect(() => goalSeek(b.doc, { sheet: b.sheet, r: 0, c: 0 }, 1, { sheet: b.sheet, r: 0, c: 1 })).toThrow(/formula/)
  })

  it('Descriptive Statistics matches the worksheet functions', () => {
    const b = book({ A1: 'Score', A2: 2, A3: 4, A4: 4, A5: 5, A6: 9 })
    descriptiveStats(b.doc, { sheet: b.sheet, r0: 0, c0: 0, r1: 5, c1: 0 }, { sheet: b.sheet, r: 0, c: 2 }, true)
    const at = (i: number) => b.get(`D${3 + i}`)
    expect(b.get('C1')).toBe('Score')
    expect(b.get('C3')).toBe('Mean')
    b.set('F1', '=AVERAGE(A2:A6)').set('F2', '=STDEV.S(A2:A6)').set('F3', '=KURT(A2:A6)').set('F4', '=SKEW(A2:A6)').set('F5', '=MODE.SNGL(A2:A6)')
    expect(at(0)).toBeCloseTo(b.get('F1') as number)
    expect(at(4)).toBeCloseTo(b.get('F2') as number)
    expect(at(6)).toBeCloseTo(b.get('F3') as number)
    expect(at(7)).toBeCloseTo(b.get('F4') as number)
    expect(at(3)).toBe(b.get('F5'))
    expect([at(2), at(8), at(11), at(12)]).toEqual([4, 7, 24, 5])
  })
})
