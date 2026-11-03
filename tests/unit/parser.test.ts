import { describe, expect, it } from 'vitest'
import { parseFormula, printFormula, type Node } from '@/core/engine/parser'

const sheets = ['Sheet1', 'Data Sheet', 'Sheet3']
const ctx = { row: 4, col: 2, sheetId: (n: string) => { const i = sheets.findIndex((s) => s.toLowerCase() === n.toLowerCase()); return i < 0 ? null : i }, sheetName: (i: number) => sheets[i] ?? null }
const rt = (f: string, at = ctx) => printFormula(parseFormula(f, at), at)
const shape = (n: Node): unknown => {
  switch (n.t) {
    case 'bin': return [n.op, shape(n.a), shape(n.b)]
    case 'un': return [n.op, shape(n.a)]
    case 'pct': return ['%', shape(n.a)]
    case 'paren': return shape(n.a)
    case 'func': return [n.name, ...n.args.map(shape)]
    case 'num': return n.v
    case 'ref': return 'ref'
    default: return n.t
  }
}

describe('formula parser', () => {
  it('round-trips references, sheets and functions', () => {
    for (const f of ['A1', '$A$1', 'A$1:$B2', 'A:C', '$2:$5', "'Data Sheet'!B2:C9", 'Sheet1:Sheet3!A1', 'SUM(A1:A10,B2)', 'IF(A1>0,"yes","no")', 'IF(A1,,1)', '{1,2;3,4}', '-2^2', 'A1#', '@A1:A10', 'SUM(A1:A3 A2:C2)', 'SUM((A1,B2))', 'Table1[[#Headers],[Col 1]:[Col2]]', 'Table1[@Price]', 'LAMBDA(x,x+1)(3)', '"say ""hi"""', '50%', '1E+21', '#N/A', 'TRUE', '[Book2.xlsx]Sheet1!$A$1', 'LOG10(100)', 'A1&" "&B1', 'TEXT(A1,"0.00")'])
      expect(rt('=' + f)).toBe(f)
  })
  it('follows Excel operator precedence', () => {
    expect(shape(parseFormula('-2^2', ctx))).toEqual(['^', ['-', 2], 2])
    expect(shape(parseFormula('1+2*3^2', ctx))).toEqual(['+', 1, ['*', 2, ['^', 3, 2]]])
    expect(shape(parseFormula('2^3^2', ctx))).toEqual(['^', ['^', 2, 3], 2])
    expect(shape(parseFormula('1&2=12', ctx))).toEqual(['=', ['&', 1, 2], 12])
    expect(shape(parseFormula('5%^2', ctx))).toEqual(['^', ['%', 5], 2])
    expect(shape(parseFormula(' 1 + 2 ', ctx))).toEqual(['+', 1, 2])
  })
  it('stores relative references as offsets, so a copy moves them', () => {
    const n = parseFormula('=A1+$B$2', { ...ctx, row: 0, col: 0 })
    expect(printFormula(n, { ...ctx, row: 3, col: 1 })).toBe('B4+$B$2')
    expect(printFormula(n, { ...ctx, row: -1 + 0, col: 0 })).toBe('#REF!+$B$2')
  })
  it('strips file prefixes', () => {
    expect(rt('=_xlfn.XLOOKUP(A1,B:B,C:C)')).toBe('XLOOKUP(A1,B:B,C:C)')
    expect(rt('=_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(4)')).toBe('LAMBDA(x,x*2)(4)')
  })
  it('rejects broken formulas', () => {
    for (const f of ['=SUM(1', '=1+', '="abc', '={1,2;3}', '=)'] ) expect(() => parseFormula(f, ctx)).toThrow()
  })
  it('unknown sheets print as #REF!', () => {
    expect(rt('=Nope!A1')).toBe('#REF!')
  })
})
