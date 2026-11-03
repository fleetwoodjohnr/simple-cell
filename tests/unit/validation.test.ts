import { describe, expect, it } from 'vitest'
import { Host } from '@/core/host'
import { parseArea } from '@/core/refs'
import type { DvSpec } from '@/core/doc'

function setup(range: string, spec: DvSpec) {
  const h = new Host()
  const s = h.doc.wb.sheets[0].id
  h.setValidation(s, [parseArea(range, s)!], spec)
  const type = (ref: string, text: string) => {
    const a = parseArea(ref, s)!
    return h.input(s, a.r0, a.c0, text)
  }
  const value = (ref: string) => {
    const a = parseArea(ref, s)!
    return h.doc.value(s, a.r0, a.c0)
  }
  return { h, s, type, value }
}

describe('data validation', () => {
  it('whole numbers between limits; a stop rule undoes the entry', () => {
    const { type, value, h } = setup('A1:A10', { type: 'whole', operator: 'between', formulas: ['1', '10'], showError: true })
    expect(type('A1', '5').ok).toBe(true)
    const bad = type('A1', '11')
    expect(bad.ok).toBe(false)
    expect(bad.alert).toEqual({ style: 'stop', title: 'Simple Cell', message: 'This value doesn’t match the data validation restrictions defined for this cell.' })
    expect(value('A1')).toBe(5)
    expect(type('A2', '2.5').ok).toBe(false)
    expect(type('A2', 'abc').ok).toBe(false)
    // Undo history holds only the accepted entry.
    expect(h.doc.undoStack.length).toBe(2) // the rule, then 5
    expect(h.doc.redoStack.length).toBe(0)
    expect(type('B1', '99').ok).toBe(true) // outside the range
  })

  it('warning and information rules keep the entry and report the alert', () => {
    const { type, value } = setup('A1', { type: 'decimal', operator: 'greaterThan', formulas: ['0'], errorStyle: 'warning', errorTitle: 'Check', error: 'Must be positive' })
    const r = type('A1', '-1')
    expect(r.ok).toBe(true)
    expect(r.alert).toEqual({ style: 'warning', title: 'Check', message: 'Must be positive' })
    expect(value('A1')).toBe(-1)
  })

  it('lists: literal items and ranges, case-insensitive, with the dropdown items', () => {
    const lit = setup('A1:A5', { type: 'list', formulas: ['Yes, No, Maybe'] })
    expect(lit.type('A1', 'no').ok).toBe(true)
    expect(lit.type('A2', 'Perhaps').ok).toBe(false)
    expect(lit.h.validationInfo(lit.s, 0, 0)?.items).toEqual(['Yes', 'No', 'Maybe'])
    const h = new Host()
    const s = h.doc.wb.sheets[0].id
    h.input(s, 0, 2, 'Red')
    h.input(s, 1, 2, 'Green')
    h.input(s, 2, 2, '2024-03-01')
    h.setValidation(s, [parseArea('A1:A3', s)!], { type: 'list', formulas: ['=$C$1:$C$5'] })
    expect(h.validationInfo(s, 1, 0)?.items).toEqual(['Red', 'Green', '3/1/2024'])
    expect(h.input(s, 0, 0, 'green').ok).toBe(true)
    expect(h.input(s, 1, 0, '2024-03-01').ok).toBe(true)
    expect(h.input(s, 2, 0, 'Blue').ok).toBe(false)
    expect(h.validationInfo(s, 0, 0)?.formulas).toEqual(['=$C$1:$C$5'])
  })

  it('dates, text length, custom formulas, blanks', () => {
    const d = setup('A1', { type: 'date', operator: 'greaterThanOrEqual', formulas: ['=DATE(2024,1,1)'] })
    expect(d.type('A1', '2023-12-31').ok).toBe(false)
    expect(d.type('A1', '2024-01-02').ok).toBe(true)
    const t = setup('A1', { type: 'textLength', operator: 'lessThanOrEqual', formulas: ['3'] })
    expect(t.type('A1', 'abcd').ok).toBe(false)
    expect(t.type('A1', 'abc').ok).toBe(true)
    const c = setup('A1:A3', { type: 'custom', formulas: ['=COUNTIF($A$1:$A$3,A1)=1'] })
    expect(c.type('A1', 'x').ok).toBe(true)
    expect(c.type('A2', 'x').ok).toBe(false) // a duplicate
    const blank = setup('A1', { type: 'whole', operator: 'greaterThan', formulas: ['0'], allowBlank: false })
    expect(blank.h.doc.checkValidation(blank.s, 0, 0)).not.toBeNull()
  })

  it('input messages, follows inserted rows, cleared by type none, invalid cells', () => {
    const { h, s, type } = setup('B2:B4', { type: 'whole', operator: 'between', formulas: ['=$D$1', '10'], showInput: true, promptTitle: 'Quantity', prompt: 'Between D1 and 10' })
    h.input(s, 0, 3, '1')
    expect(h.validationInfo(s, 1, 1)).toMatchObject({ prompt: 'Between D1 and 10', promptTitle: 'Quantity', ranges: 'B2:B4' })
    h.doc.insertRows(s, 0, 1)
    expect(h.validationInfo(s, 2, 1)).toMatchObject({ ranges: 'B3:B5', formulas: ['=$D$2', '=10'] })
    expect(type('B3', '0').ok).toBe(false)
    h.input(s, 0, 3, '9') // D1 is now an unrelated cell
    expect(type('B3', '5').ok).toBe(true)
    h.setValidation(s, [parseArea('B3:B5', s)!], { type: 'whole', operator: 'equal', formulas: ['1'] })
    expect(h.invalidCells(s)).toEqual([{ r: 2, c: 1 }])
    h.setValidation(s, [parseArea('B1:B10', s)!], { type: 'none' })
    expect(h.validationInfo(s, 2, 1)).toBeNull()
  })
})
