import { describe, expect, it } from 'vitest'
import { book } from './helpers/book'
import { CASES, DATA, DEVIATIONS } from './functions/cases'

const b = book(DATA)

/** Numbers match to 9 significant digits, or to the digits the expected value is written with (doc examples are rounded). */
function close(actual: unknown, expected: unknown, rel = 0) {
  if (typeof expected === 'number' && typeof actual === 'number') {
    if (rel && Math.abs(actual - expected) <= Math.abs(expected) * rel) return true
    if (expected === 0) return Math.abs(actual) < 1e-9
    const digits = expected.toExponential().split('e')[0].replace(/[-.]/g, '').length
    const exp = Math.floor(Math.log10(Math.abs(expected)))
    const tol = Math.max(Math.abs(expected) * 5e-9, 0.5000001 * Math.pow(10, exp - Math.min(digits, 10) + 1))
    return Math.abs(actual - expected) <= tol
  }
  return actual === expected
}

describe('worksheet functions', () => {
  for (const [name, cases] of Object.entries(CASES)) {
    it(name, () => {
      const failures: string[] = []
      for (const [formula, want] of cases) {
        const got = b.calc(formula)
        if (!close(got, want, DEVIATIONS[name]?.tolerance)) failures.push(`=${formula} → ${JSON.stringify(got)} (expected ${JSON.stringify(want)})`)
      }
      expect(failures).toEqual([])
    })
  }
})
