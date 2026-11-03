import { describe, expect, it } from 'vitest'
import { learn, run } from '@/core/flashfill'
import { book } from './helpers/book'

const ex = (pairs: [string | string[], string][]) => pairs.map(([i, o]) => ({ inputs: Array.isArray(i) ? i : [i], output: o }))

describe('flash fill', () => {
  it('learns common transformations from one or two examples', () => {
    const cases: { train: [string | string[], string][]; test: [string | string[], string][] }[] = [
      { train: [['John Smith', 'Smith']], test: [['Mary Ann Lee', 'Lee'], ['Bo Zhang', 'Zhang']] },
      { train: [['John Smith', 'John']], test: [['Mary Lee', 'Mary']] },
      { train: [['John Smith', 'Smith, John']], test: [['Mary Lee', 'Lee, Mary']] },
      { train: [['john smith', 'J. Smith']], test: [['mary lee', 'M. Lee']] },
      { train: [['John Ronald Tolkien', 'JRT']], test: [['George Raymond Martin', 'GRM']] },
      { train: [['ann.lee@example.com', 'ann.lee']], test: [['bo@corp.org', 'bo']] },
      { train: [[['Ann', 'Lee'], 'Lee, Ann']], test: [[['Bo', 'Zhang'], 'Zhang, Bo']] },
      { train: [['hello world', 'HELLO WORLD']], test: [['good day', 'GOOD DAY']] },
      { train: [['SKU-1042-B', '1042']], test: [['SKU-77-A', '77']] },
      { train: [['2024-03-15', '15/03/2024']], test: [['1999-12-01', '01/12/1999']] }
    ]
    for (const c of cases) {
      const p = learn(ex(c.train))
      expect(p, JSON.stringify(c.train)).not.toBeNull()
      for (const [i, o] of c.test) expect(run(p!, Array.isArray(i) ? i : [i]), JSON.stringify(i)).toBe(o)
    }
  })

  it('a second example fixes an ambiguous first one', () => {
    // From "Ann Lee" → "Lee" alone, "second word" and "last word" both fit; the second example decides.
    const p = learn(ex([['Ann Lee', 'Lee'], ['Mary Jo Ray', 'Ray']]))
    expect(run(p!, ['Al Bo Cy Di'])).toBe('Di')
    expect(learn(ex([['abc', 'xyz']]))).toBeNull() // nothing in common: no pattern
  })

  it('fills the column in the sheet (Ctrl+E), skipping a header row', () => {
    const b = book({ A1: 'Full name', B1: 'First', A2: 'Ann Lee', B2: 'Ann', A3: 'Bo Zhang', A4: 'Cy Ray', A5: 'Di Moss' })
    expect(b.doc.flashFill(b.sheet, 2, 1)).toBe(3)
    expect([b.get('B3'), b.get('B4'), b.get('B5')]).toEqual(['Bo', 'Cy', 'Di'])
    b.doc.undo()
    expect(b.get('B3')).toBeNull()
    expect(() => b.doc.flashFill(b.sheet, 0, 3)).toThrow(/data next to/)
  })
})
