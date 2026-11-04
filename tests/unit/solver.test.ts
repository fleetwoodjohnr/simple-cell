import { describe, expect, it } from 'vitest'
import { unzipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'
import { branchAndBound, emptySolverModel, finishSolver, nelderMead, readSolverModel, simplex, solve, type SolverModel } from '@/core/solver'

describe('linear programming', () => {
  it('solves a textbook LP (Hillier & Lieberman’s Wyndor Glass: 3x + 5y → 36 at (2, 6))', () => {
    const r = simplex({ c: [3, 5], rows: [{ a: [1, 0], rel: '<=', b: 4 }, { a: [0, 2], rel: '<=', b: 12 }, { a: [3, 2], rel: '<=', b: 18 }] })
    expect(r.status).toBe('optimal')
    if (r.status !== 'optimal') return
    expect(r.z).toBeCloseTo(36, 9)
    expect(r.x[0]).toBeCloseTo(2, 9)
    expect(r.x[1]).toBeCloseTo(6, 9)
  })

  it('handles ≥ and = rows (phase one), and reports infeasible and unbounded problems', () => {
    // minimize 2x + 3y with x + y ≥ 4, x + 3y ≥ 6 → (3, 1), 9
    const r = simplex({ c: [-2, -3], rows: [{ a: [1, 1], rel: '>=', b: 4 }, { a: [1, 3], rel: '>=', b: 6 }] })
    expect(r.status === 'optimal' && -r.z).toBeCloseTo(9, 9)
    const eq = simplex({ c: [1, 1], rows: [{ a: [1, 2], rel: '=', b: 4 }, { a: [1, 0], rel: '<=', b: 3 }] })
    expect(eq.status === 'optimal' && eq.z).toBeCloseTo(3.5, 9)
    expect(simplex({ c: [1, 1], rows: [{ a: [1, 1], rel: '<=', b: 1 }, { a: [1, 1], rel: '>=', b: 2 }] }).status).toBe('infeasible')
    expect(simplex({ c: [1, 0], rows: [{ a: [1, -1], rel: '<=', b: 1 }] }).status).toBe('unbounded')
  })

  it('branch and bound finds the integer optimum, not the rounded LP one', () => {
    // max y: −x + y ≤ 1, 3x + 2y ≤ 12, 2x + 3y ≤ 12 → LP (1.8, 2.8); integer optimum y = 2
    const lp = { c: [0, 1], rows: [{ a: [-1, 1], rel: '<=' as const, b: 1 }, { a: [3, 2], rel: '<=' as const, b: 12 }, { a: [2, 3], rel: '<=' as const, b: 12 }] }
    expect(simplex(lp).status === 'optimal' && (simplex(lp) as { z: number }).z).toBeCloseTo(2.8, 9)
    const r = branchAndBound(lp, [
      [1, 0],
      [0, 1]
    ])
    expect(r.status).toBe('optimal')
    if (r.status !== 'optimal') return
    expect(r.z).toBeCloseTo(2, 9)
    expect(Number.isInteger(Math.round(r.x[0] * 1e9) / 1e9)).toBe(true)
  })

  it('Nelder–Mead finds the Rosenbrock minimum', () => {
    const r = nelderMead(([x, y]) => 100 * (y - x * x) ** 2 + (1 - x) ** 2, [-1.2, 1], { maxEvals: 5000 })
    expect(r.x[0]).toBeCloseTo(1, 4)
    expect(r.x[1]).toBeCloseTo(1, 4)
  })
})

/** Wyndor on a sheet: B1:B2 variables, D1 the profit, C1:C3 the constraint rows, E1:E3 their limits. */
function wyndor() {
  const h = new Host()
  const doc = h.doc
  const s = doc.wb.sheets[0].id
  const put = (ref: [number, number], v: string) => doc.input(s, ref[0], ref[1], v)
  put([0, 0], 'Doors')
  put([1, 0], 'Windows')
  put([0, 1], '0')
  put([1, 1], '0')
  put([0, 2], '=B1')
  put([1, 2], '=2*B2')
  put([2, 2], '=3*B1+2*B2')
  put([0, 4], '4')
  put([1, 4], '12')
  put([2, 4], '18')
  put([0, 3], '=3*B1+5*B2')
  const model: SolverModel = { ...emptySolverModel(), objective: '$D$1', goal: 'max', changing: '$B$1:$B$2', constraints: [{ lhs: '$C$1:$C$3', rel: '<=', rhs: '$E$1:$E$3' }] }
  return { h, doc, s, model }
}

describe('Solver on a sheet', () => {
  it('Simplex LP: finds the optimum, leaves the sheet as it was, and Keep applies it as one undoable step', () => {
    const { doc, s, model } = wyndor()
    const r = solve(doc, s, model)
    expect(r.status).toBe('optimal')
    expect(r.values).toEqual([2, 6])
    expect(r.objective).toBe(36)
    expect(doc.value(s, 0, 1)).toBe(0)
    expect(doc.value(s, 0, 3)).toBe(0)
    finishSolver(doc, s, model, r, null)
    expect([doc.value(s, 0, 1), doc.value(s, 1, 1), doc.value(s, 0, 3)]).toEqual([2, 6, 36])
    doc.undo()
    expect([doc.value(s, 0, 1), doc.value(s, 1, 1), doc.value(s, 0, 3)]).toEqual([0, 0, 0])
  })

  it('minimize, Value Of, and variables that may go negative', () => {
    const { doc, s, model } = wyndor()
    expect(solve(doc, s, { ...model, goal: 'min' }).objective).toBe(0)
    const v = solve(doc, s, { ...model, goal: 'value', value: 21 })
    expect(v.status).toBe('optimal')
    expect(v.objective).toBeCloseTo(21, 9)
    // A lower bound below zero: "non-negative" leaves bounded variables alone (Excel's rule).
    const neg = solve(doc, s, { ...model, goal: 'min', constraints: [...model.constraints, { lhs: '$B$1', rel: '>=', rhs: '-2' }] })
    expect(neg.values).toEqual([-2, 0])
    expect(neg.objective).toBe(-6)
  })

  it('reports a model that is not linear, infeasible or unbounded', () => {
    const { doc, s, model } = wyndor()
    doc.input(s, 0, 3, '=B1*B2')
    expect(solve(doc, s, model).status).toBe('not-linear')
    doc.input(s, 0, 3, '=3*B1+5*B2')
    expect(solve(doc, s, { ...model, constraints: [...model.constraints, { lhs: '$B$1', rel: '>=', rhs: '5' }] }).status).toBe('infeasible')
    expect(solve(doc, s, { ...model, constraints: [] }).status).toBe('unbounded')
  })

  it('integer constraints (branch and bound) and binary ones', () => {
    const { doc, s, model } = wyndor()
    // Limits that make the LP optimum fractional: 3x + 2y ≤ 17 → LP (5/3, 6), integer (1, 6) = 33 or (2, 5) = 31…
    doc.input(s, 2, 4, '17')
    expect(solve(doc, s, model).values[0]).toBeCloseTo(5 / 3, 9)
    const r = solve(doc, s, { ...model, constraints: [...model.constraints, { lhs: '$B$1:$B$2', rel: 'int', rhs: '' }] })
    expect(r.status).toBe('optimal')
    expect(r.values).toEqual([1, 6])
    expect(r.objective).toBe(33)
    const b = solve(doc, s, { ...model, constraints: [...model.constraints, { lhs: '$B$1', rel: 'bin', rhs: '' }] })
    expect(b.values[0]).toBe(1)
  })

  it('nonlinear: maximize x·y with x + y ≤ 10 → (5, 5)', () => {
    const h = new Host()
    const doc = h.doc
    const s = doc.wb.sheets[0].id
    doc.input(s, 0, 0, '1')
    doc.input(s, 1, 0, '1')
    doc.input(s, 2, 0, '=A1*A2')
    doc.input(s, 3, 0, '=A1+A2')
    const r = solve(doc, s, { ...emptySolverModel(), objective: 'A3', changing: 'A1:A2', constraints: [{ lhs: 'A4', rel: '<=', rhs: '10' }], method: 'nonlinear' })
    expect(r.status).toBe('converged')
    expect(r.values[0]).toBeCloseTo(5, 3)
    expect(r.values[1]).toBeCloseTo(5, 3)
    expect(r.objective).toBeCloseTo(25, 3)
  })

  it('refuses formulas in variable cells, and cells on other sheets', () => {
    const { doc, s, model } = wyndor()
    expect(() => solve(doc, s, { ...model, changing: '$C$1' })).toThrow(/contains a formula/)
    doc.addSheet('Other')
    expect(() => solve(doc, s, { ...model, objective: 'Other!A1' })).toThrow(/sheet being solved/)
  })

  it('Answer Report: objective, variables and constraints with their status and slack', () => {
    const { doc, s, model } = wyndor()
    const r = solve(doc, s, model)
    finishSolver(doc, s, model, r, r)
    const rep = doc.wb.sheetByName('Answer Report 1')!
    const rows: string[] = []
    for (let i = 0; i < 25; i++) {
      const line = [1, 2, 3, 4, 5, 6].map((c) => rep.get(i, c)?.v ?? '').join('|')
      if (line.replace(/\|/g, '')) rows.push(line)
    }
    // Names come from the text to the left and above (Excel's rule): row 1 is "Doors".
    expect(rows).toContain('$D$1|Doors|0|36||')
    expect(rows).toContain('$B$1|Doors|0|2|Contin|')
    expect(rows).toContain('$C$1|Doors|2|$C$1<=$E$1|Not Binding|2')
    expect(rows).toContain('$C$2|Windows|12|$C$2<=$E$2|Binding|0')
  })

  it('the model is kept as Excel’s hidden sheet names and survives an .xlsx round trip', async () => {
    const { h, doc, s, model } = wyndor()
    const m = { ...model, constraints: [...model.constraints, { lhs: '$B$1:$B$2', rel: 'int' as const, rhs: '' }] }
    finishSolver(doc, s, m, null, null)
    const names = Object.fromEntries(doc.wb.names.filter((n) => n.scope === s).map((n) => [n.name, n.text]))
    expect(names).toMatchObject({ solver_opt: 'Sheet1!$D$1', solver_adj: 'Sheet1!$B$1:$B$2', solver_typ: '1', solver_num: '2', solver_lhs1: 'Sheet1!$C$1:$C$3', solver_rel1: '1', solver_rhs1: 'Sheet1!$E$1:$E$3', solver_rel2: '4', solver_rhs2: '"integer"', solver_eng: '2', solver_neg: '1' })
    const saved = (await h.save('xlsx')).data
    const xml = strFromU8(unzipSync(saved)['xl/workbook.xml'])
    expect(xml).toMatch(/<definedName name="solver_opt" localSheetId="0" hidden="1">Sheet1!\$D\$1<\/definedName>/)
    const back = await Host.open(saved, 'xlsx')
    const bs = back.doc.wb.sheets[0]
    expect(readSolverModel(back.doc.wb.names, bs.id, bs.name)).toEqual(m)
  })
})
