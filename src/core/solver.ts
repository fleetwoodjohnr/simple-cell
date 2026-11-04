// Solver (Excel's Data ▸ Solver): maximize, minimize or reach a value in an objective cell by changing variable cells,
// subject to constraints. Two methods:
// - Simplex LP: the model's linearity is checked from the sheet itself, then a two-phase simplex (Bland's rule) solves
//   it; integer and binary constraints are handled by branch and bound.
// - Nonlinear: Nelder–Mead on a penalized objective, with growing penalties. It finds a local optimum; it is not Excel's
//   GRG algorithm, and the results can differ on problems with several local optima.
// Every evaluation sets the variable cells and recalculates the workbook, as Goal Seek does; the cells are restored
// afterwards. The model is kept the way Excel keeps it, as hidden sheet-level names (solver_opt, solver_adj, …), so
// Excel's Solver reads models saved here and the other way round.

import type { Doc, Tx } from './doc'
import type { Cell, DefinedName } from './model/workbook'
import { areaName, cellName, parseArea, quoteSheet } from './refs'
import { numberToText } from './engine/values'

export type SolverRel = '<=' | '>=' | '=' | 'int' | 'bin'
export interface SolverConstraint {
  lhs: string
  rel: SolverRel
  /** A number or a reference (empty for int and bin). */
  rhs: string
}
export interface SolverModel {
  objective: string
  goal: 'max' | 'min' | 'value'
  value: number
  changing: string
  constraints: SolverConstraint[]
  /** Variables without a lower-bound constraint stay ≥ 0 (Excel's "Make Unconstrained Variables Non-Negative"). */
  nonNegative: boolean
  method: 'simplex' | 'nonlinear'
}

export const emptySolverModel = (): SolverModel => ({ objective: '', goal: 'max', value: 0, changing: '', constraints: [], nonNegative: true, method: 'simplex' })

export type SolverStatus = 'optimal' | 'converged' | 'infeasible' | 'unbounded' | 'not-linear' | 'limit'
export interface SolverResult {
  status: SolverStatus
  message: string
  /** Variable cells, in order (A1 names), with their original and solved values. */
  cells: string[]
  original: number[]
  values: number[]
  originalObjective: number
  objective: number
  /** For the Answer Report: labels, and each constraint cell's value at the solution. */
  report: {
    objective: { cell: string; name: string }
    vars: { cell: string; name: string; integer: boolean }[]
    constraints: { cell: string; name: string; value: number; formula: string; binding: boolean; slack: number }[]
  }
}

const MESSAGES: Record<SolverStatus, string> = {
  optimal: 'Solver found a solution. All constraints and optimality conditions are satisfied.',
  converged: 'Solver has converged to the current solution. All constraints are satisfied. (The nonlinear method finds a local optimum.)',
  infeasible: 'Solver could not find a feasible solution.',
  unbounded: 'The objective cell values do not converge: the problem is unbounded.',
  'not-linear': 'The linearity conditions required by the Simplex LP method are not satisfied. Try the nonlinear method.',
  limit: 'Solver stopped at its iteration limit before finding the optimum.'
}

const MAX_VARS = 200
const EPS = 1e-9

// ---------- linear programming ----------

export type LpRel = '<=' | '>=' | '='
export interface Lp {
  /** Maximize c·x with x ≥ 0. */
  c: number[]
  rows: { a: number[]; rel: LpRel; b: number }[]
}
export type LpResult = { status: 'optimal'; x: number[]; z: number } | { status: 'infeasible' | 'unbounded' | 'limit' }

/** Two-phase simplex on a dense tableau, with Bland's rule (no cycling). */
export function simplex(lp: Lp, maxIter = 20_000): LpResult {
  const n = lp.c.length
  const rows = lp.rows.map((r) => (r.b < 0 ? { a: r.a.map((v) => -v), rel: (r.rel === '<=' ? '>=' : r.rel === '>=' ? '<=' : '=') as LpRel, b: -r.b } : r))
  const m = rows.length
  let nS = 0
  let nA = 0
  const slack = rows.map((r) => (r.rel === '=' ? -1 : nS++))
  const art = rows.map((r) => (r.rel === '<=' ? -1 : nA++))
  const W = n + nS + nA
  const T = rows.map((r, i) => {
    const row = new Array<number>(W + 1).fill(0)
    for (let j = 0; j < n; j++) row[j] = r.a[j] ?? 0
    if (slack[i] >= 0) row[n + slack[i]] = r.rel === '<=' ? 1 : -1
    if (art[i] >= 0) row[n + nS + art[i]] = 1
    row[W] = r.b
    return row
  })
  const basis = rows.map((r, i) => (r.rel === '<=' ? n + slack[i] : n + nS + art[i]))
  let iterations = 0

  const pivot = (pr: number, pc: number) => {
    const p = T[pr][pc]
    const prow = T[pr]
    for (let j = 0; j <= W; j++) prow[j] /= p
    for (let i = 0; i < m; i++) {
      if (i === pr) continue
      const f = T[i][pc]
      if (Math.abs(f) > 0) for (let j = 0; j <= W; j++) T[i][j] -= f * prow[j]
    }
    basis[pr] = pc
  }
  // Maximize obj over the columns `allowed` lets enter.
  const run = (obj: number[], allowed: number): 'optimal' | 'unbounded' | 'limit' => {
    const inBasis = new Set(basis)
    while (true) {
      if (iterations++ > maxIter) return 'limit'
      let pc = -1
      for (let j = 0; j < allowed; j++) {
        if (inBasis.has(j)) continue
        let rj = obj[j]
        for (let i = 0; i < m; i++) rj -= obj[basis[i]] * T[i][j]
        if (rj > EPS) {
          pc = j
          break
        }
      }
      if (pc < 0) return 'optimal'
      let pr = -1
      let best = Infinity
      for (let i = 0; i < m; i++) {
        if (T[i][pc] <= EPS) continue
        const ratio = T[i][W] / T[i][pc]
        if (ratio < best - EPS || (ratio <= best + EPS && pr >= 0 && basis[i] < basis[pr])) {
          best = ratio
          pr = i
        }
      }
      if (pr < 0) return 'unbounded'
      inBasis.delete(basis[pr])
      pivot(pr, pc)
      inBasis.add(pc)
    }
  }

  if (nA) {
    const obj1 = new Array<number>(W).fill(0)
    for (let j = n + nS; j < W; j++) obj1[j] = -1
    const r1 = run(obj1, W)
    if (r1 === 'limit') return { status: 'limit' }
    const infeas = basis.reduce((s, b, i) => (b >= n + nS ? s + T[i][W] : s), 0)
    if (infeas > 1e-7 * Math.max(1, ...rows.map((r) => Math.abs(r.b)))) return { status: 'infeasible' }
    // Drive artificials still in the basis (at zero) out, where a real column can replace them.
    for (let i = 0; i < m; i++) {
      if (basis[i] < n + nS) continue
      for (let j = 0; j < n + nS; j++)
        if (Math.abs(T[i][j]) > EPS && !basis.includes(j)) {
          pivot(i, j)
          break
        }
    }
  }
  const obj2 = new Array<number>(W).fill(0)
  for (let j = 0; j < n; j++) obj2[j] = lp.c[j]
  const r2 = run(obj2, n + nS)
  if (r2 !== 'optimal') return { status: r2 }
  const x = new Array<number>(n).fill(0)
  basis.forEach((b, i) => {
    if (b < n) x[b] = T[i][W]
  })
  return { status: 'optimal', x, z: lp.c.reduce((s, cj, j) => s + cj * x[j], 0) }
}

/**
 * Branch and bound: each of `ints` is a linear expression of x (coefficients) that must be a whole number. Exact: no
 * optimality gap.
 */
export function branchAndBound(lp: Lp, ints: number[][], maxNodes = 20_000): LpResult {
  if (!ints.length) return simplex(lp)
  let best: { x: number[]; z: number } | null = null
  const stack: Lp['rows'][] = [[]]
  let nodes = 0
  let limited = false
  while (stack.length) {
    if (nodes++ > maxNodes) {
      limited = true
      break
    }
    const extra = stack.pop()!
    const r = simplex({ c: lp.c, rows: [...lp.rows, ...extra] })
    if (r.status === 'unbounded' && nodes === 1) return r
    if (r.status === 'limit') limited = true
    if (r.status !== 'optimal') continue
    if (best && r.z <= best.z + 1e-9 * Math.max(1, Math.abs(best.z))) continue
    const value = (a: number[]) => a.reduce((s, v, k) => s + v * r.x[k], 0)
    const j = ints.find((a) => Math.abs(value(a) - Math.round(value(a))) > 1e-6)
    if (j === undefined) {
      best = { x: r.x, z: r.z }
      continue
    }
    // Depth first; the branch nearer the LP value is explored first.
    const v = value(j)
    const down = { a: j, rel: '<=' as const, b: Math.floor(v) }
    const up = { a: j, rel: '>=' as const, b: Math.ceil(v) }
    const nearUp = v - Math.floor(v) > 0.5
    stack.push([...extra, nearUp ? down : up], [...extra, nearUp ? up : down])
  }
  if (best) return { status: 'optimal', ...best }
  return { status: limited ? 'limit' : 'infeasible' }
}

// ---------- nonlinear ----------

/** Nelder–Mead minimization of f from x0. */
export function nelderMead(f: (x: number[]) => number, x0: number[], opts: { maxEvals?: number; tol?: number } = {}): { x: number[]; fx: number } {
  const n = x0.length
  const maxEvals = opts.maxEvals ?? 400 * (n + 1)
  const tol = opts.tol ?? 1e-10
  let evals = 0
  const F = (x: number[]) => {
    evals++
    const v = f(x)
    return Number.isNaN(v) ? Infinity : v
  }
  const pts: { x: number[]; fx: number }[] = [{ x: [...x0], fx: F(x0) }]
  for (let i = 0; i < n; i++) {
    const x = [...x0]
    x[i] += Math.max(Math.abs(x[i]) * 0.1, 0.5)
    pts.push({ x, fx: F(x) })
  }
  while (evals < maxEvals) {
    pts.sort((a, b) => a.fx - b.fx)
    const spread = Math.abs(pts[n].fx - pts[0].fx)
    const size = Math.max(...pts.slice(1).map((p) => Math.max(...p.x.map((v, i) => Math.abs(v - pts[0].x[i])))))
    if (spread <= tol * (Math.abs(pts[0].fx) + tol) && size < 1e-9 * (1 + Math.max(...pts[0].x.map(Math.abs)))) break
    const centroid = new Array<number>(n).fill(0)
    for (let k = 0; k < n; k++) for (let i = 0; i < n; i++) centroid[i] += pts[k].x[i] / n
    const along = (t: number) => centroid.map((c, i) => c + t * (pts[n].x[i] - c))
    const xr = along(-1)
    const fr = F(xr)
    if (fr < pts[0].fx) {
      const xe = along(-2)
      const fe = F(xe)
      pts[n] = fe < fr ? { x: xe, fx: fe } : { x: xr, fx: fr }
    } else if (fr < pts[n - 1].fx) pts[n] = { x: xr, fx: fr }
    else {
      const xc = fr < pts[n].fx ? along(-0.5) : along(0.5)
      const fc = F(xc)
      if (fc < Math.min(fr, pts[n].fx)) pts[n] = { x: xc, fx: fc }
      else
        for (let k = 1; k <= n; k++) {
          const x = pts[k].x.map((v, i) => pts[0].x[i] + 0.5 * (v - pts[0].x[i]))
          pts[k] = { x, fx: F(x) }
        }
    }
  }
  pts.sort((a, b) => a.fx - b.fx)
  return pts[0]
}

// ---------- the model on the sheet ----------

type Ref = { r: number; c: number }

/** Cells of "B1:B3,D1" (optionally "Sheet1!…", which must be the solver's sheet). */
function cellsOf(doc: Doc, sheet: number, text: string): Ref[] {
  const out: Ref[] = []
  for (const raw of text.split(',')) {
    let t = raw.trim()
    if (!t) continue
    const bang = t.lastIndexOf('!')
    if (bang > 0) {
      const name = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
      if (doc.wb.sheetByName(name)?.id !== sheet) throw new Error(`${raw.trim()}: Solver’s cells must be on the sheet being solved.`)
      t = t.slice(bang + 1)
    }
    const a = parseArea(t, sheet)
    if (!a || (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1) > 10_000) throw new Error(`“${raw.trim()}” isn’t a cell or range.`)
    for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) out.push({ r, c })
  }
  return out
}

const same = (a: Ref, b: Ref) => a.r === b.r && a.c === b.c

/** The right-hand side of one cell of a constraint, as the report's formula shows it ("$C$1" or "10"). */
function rhsCellsText(rhs: string, i: number, n: number): string {
  const t = rhs.trim().replace(/^=/, '')
  if (t !== '' && Number.isFinite(Number(t))) return numberToText(Number(t))
  const parts = t.split(',')
  if (parts.length > 1 || t.includes(':')) {
    // A range: the i-th cell of it.
    const a = parseArea(t.replace(/^.*!/, ''))
    if (a) {
      const w = a.c1 - a.c0 + 1
      const k = n === 1 ? 0 : i
      return absolute(cellName(a.r0 + Math.floor(k / w), a.c0 + (k % w)))
    }
  }
  return absolute(t.replace(/^.*!/, ''))
}

interface Row {
  rel: '<=' | '>=' | '='
  /** lhs − rhs at the current variable values. */
  g: () => number
  /** A constraint from the model (not an implied bound): its cell and right-hand side, for the report. */
  cell?: Ref
  rhs?: () => number
  rhsText?: string
}

/** Binds a model to the sheet: sets variables, reads the objective and the constraints. */
class Bound {
  readonly vars: Ref[]
  private readonly saved: (Cell | undefined)[]
  readonly start: number[]
  readonly rows: Row[] = []
  readonly ints = new Set<number>()
  readonly bins = new Set<number>()
  /** Variables with a lower bound given by a constraint (exempt from "non-negative"). */
  readonly bounded = new Set<number>()
  readonly objective: Ref

  constructor(
    private doc: Doc,
    private sheet: number,
    m: SolverModel
  ) {
    const sh = doc.sheet(sheet)
    const obj = cellsOf(doc, sheet, m.objective)
    if (obj.length !== 1) throw new Error('The objective must be a single cell.')
    this.objective = obj[0]
    this.vars = cellsOf(doc, sheet, m.changing)
    if (!this.vars.length) throw new Error('Enter the variable cells Solver may change.')
    if (this.vars.length > MAX_VARS) throw new Error(`Solver can change at most ${MAX_VARS} cells; this model has ${this.vars.length}.`)
    this.saved = this.vars.map((v) => sh.get(v.r, v.c))
    this.saved.forEach((cell, i) => {
      if (cell?.f) throw new Error(`The variable cell ${cellName(this.vars[i].r, this.vars[i].c)} contains a formula. Variable cells must contain values.`)
    })
    this.start = this.saved.map((cell) => (typeof cell?.v === 'number' ? cell.v : 0))
    const varIndex = (r: Ref) => this.vars.findIndex((v) => same(v, r))
    for (const k of m.constraints) {
      const lhs = cellsOf(doc, sheet, k.lhs)
      if (!lhs.length) throw new Error('A constraint has no cell reference.')
      if (k.rel === 'int' || k.rel === 'bin') {
        for (const r of lhs) {
          const i = varIndex(r)
          if (i < 0) throw new Error(`${k.rel === 'int' ? 'Integer' : 'Binary'} constraints must refer to variable cells (${cellName(r.r, r.c)} isn’t one).`)
          this.ints.add(i)
          if (k.rel === 'bin') {
            this.bins.add(i)
            this.bounded.add(i)
          }
        }
        continue
      }
      const rhsText = k.rhs.trim().replace(/^=/, '')
      const num = Number(rhsText)
      let rhs: (i: number) => number
      if (rhsText !== '' && Number.isFinite(num)) rhs = () => num
      else {
        const cells = cellsOf(doc, sheet, rhsText)
        if (cells.length !== 1 && cells.length !== lhs.length) throw new Error(`The constraint ${k.lhs} ${k.rel} ${k.rhs} compares ${lhs.length} cells with ${cells.length}.`)
        rhs = (i) => this.read(cells.length === 1 ? cells[0] : cells[i])
      }
      lhs.forEach((r, i) => {
        if (k.rel !== '<=' && varIndex(r) >= 0) this.bounded.add(varIndex(r))
        const rhsText = rhsCellsText(k.rhs, i, lhs.length)
        this.rows.push({ rel: k.rel as Row['rel'], g: () => this.read(r) - rhs(i), cell: r, rhs: () => rhs(i), rhsText })
      })
    }
    if (m.nonNegative) this.vars.forEach((_, i) => !this.bounded.has(i) && this.rows.push({ rel: '>=', g: () => this.x[i] }))
    for (const i of this.bins) this.rows.push({ rel: '<=', g: () => this.x[i] - 1 }, { rel: '>=', g: () => this.x[i] })
  }

  private x: number[] = []

  set(x: number[]) {
    this.x = x
    const sh = this.doc.sheet(this.sheet)
    this.vars.forEach((v, i) => {
      sh.set(v.r, v.c, { ...(this.saved[i] ?? {}), v: x[i] })
      this.doc.engine.touch(this.sheet, v.r, v.c)
    })
    this.doc.engine.recalc()
  }

  read(r: Ref): number {
    const v = this.doc.sheet(this.sheet).get(r.r, r.c)?.v
    return typeof v === 'number' ? v : v === null || v === undefined ? 0 : NaN
  }

  restore() {
    const sh = this.doc.sheet(this.sheet)
    this.vars.forEach((v, i) => {
      sh.set(v.r, v.c, this.saved[i])
      this.doc.engine.touch(this.sheet, v.r, v.c)
    })
    this.doc.engine.recalc()
  }

  /** Largest constraint violation at the current values. */
  violation(): number {
    let worst = 0
    for (const row of this.rows) {
      const g = row.g()
      const v = row.rel === '<=' ? g : row.rel === '>=' ? -g : Math.abs(g)
      worst = Math.max(worst, Number.isNaN(g) ? Infinity : v)
    }
    return worst
  }
}

/** Run Solver. The sheet is left as it was; the result says what the variables would become. */
export function solve(doc: Doc, sheet: number, m: SolverModel): SolverResult {
  const b = new Bound(doc, sheet, m)
  const n = b.vars.length
  const objAt = () => b.read(b.objective)
  try {
    b.set(b.start)
    const originalObjective = objAt()
    if (Number.isNaN(originalObjective)) throw new Error('The objective cell must contain a formula that returns a number.')
    const sign = m.goal === 'min' ? -1 : 1
    let status: SolverStatus
    let x: number[] = b.start
    if (m.method === 'simplex') {
      // Coefficients from unit steps away from zero, then a check at two other points that the model is linear.
      const zero = new Array<number>(n).fill(0)
      const at = (p: number[]) => {
        b.set(p)
        return { obj: objAt(), g: b.rows.map((r) => r.g()) }
      }
      const base = at(zero)
      const cols = zero.map((_, j) => at(zero.map((_, k) => (k === j ? 1 : 0))))
      const c = cols.map((col) => col.obj - base.obj)
      const A = b.rows.map((_, i) => cols.map((col) => col.g[i] - base.g[i]))
      const linearAt = (p: number[]) => {
        const v = at(p)
        const close = (actual: number, predicted: number) => Math.abs(actual - predicted) <= 1e-7 * Math.max(1, Math.abs(actual), Math.abs(predicted))
        return close(v.obj, base.obj + c.reduce((s, cj, j) => s + cj * p[j], 0)) && v.g.every((g, i) => close(g, base.g[i] + A[i].reduce((s, a, j) => s + a * p[j], 0)))
      }
      const probe = zero.map((_, j) => 1.5 + j * 0.37)
      if (![base.obj, ...base.g, ...c, ...A.flat()].every(Number.isFinite) || !linearAt(probe) || !linearAt(b.start.map((v, j) => v + probe[j] * 2))) status = 'not-linear'
      else {
        // x = x⁺ − x⁻ for every variable (the non-negativity rows above keep x⁻ at zero where they apply).
        const split = (row: number[]) => [...row, ...row.map((v) => -v)]
        const rows: Lp['rows'] = b.rows.map((r, i) => ({ a: split(A[i]), rel: r.rel, b: -base.g[i] }))
        let cost = split(c.map((v) => sign * v))
        if (m.goal === 'value') {
          rows.push({ a: split(c), rel: '=', b: m.value - base.obj })
          cost = cost.map(() => 0)
        }
        // An integer variable is x⁺ − x⁻: branch on that expression.
        const r = branchAndBound({ c: cost, rows }, [...b.ints].map((i) => split(zero.map((_, k) => (k === i ? 1 : 0)))))
        if (r.status === 'optimal') {
          x = zero.map((_, j) => clean(r.x[j] - r.x[n + j]))
          b.set(x)
          status = b.violation() <= 1e-6 * Math.max(1, ...x.map(Math.abs)) ? 'optimal' : 'infeasible'
        } else status = r.status
      }
    } else {
      if (b.ints.size) throw new Error('Integer and binary constraints need the Simplex LP method in Simple Cell.')
      const f = (p: number[], mu: number) => {
        b.set(p)
        const o = objAt()
        const goal = m.goal === 'value' ? (o - m.value) ** 2 : -sign * o
        let pen = 0
        for (const row of b.rows) {
          const g = row.g()
          const v = row.rel === '<=' ? Math.max(0, g) : row.rel === '>=' ? Math.max(0, -g) : g
          pen += v * v
        }
        return goal + mu * pen
      }
      let best = { x: b.start, fx: Infinity }
      // Growing penalties, each starting from the last result; a restart per stage gets Nelder–Mead out of stalls.
      for (const mu of [1e2, 1e4, 1e6, 1e8, 1e10]) {
        for (let run = 0; run < 2; run++) best = nelderMead((p) => f(p, mu), best.x)
        if (!Number.isFinite(best.fx) || Math.abs(objAt()) > 1e30) break
      }
      x = best.x.map(clean)
      b.set(x)
      const o = objAt()
      status = Math.abs(o) > 1e30 ? 'unbounded' : b.violation() <= 1e-6 * Math.max(1, ...x.map(Math.abs)) ? 'converged' : 'infeasible'
    }
    const good = status === 'optimal' || status === 'converged'
    b.set(good ? x : b.start)
    const label = (p: Ref) => labelOf(doc, sheet, p.r, p.c)
    const report: SolverResult['report'] = {
      objective: { cell: absolute(cellName(b.objective.r, b.objective.c)), name: label(b.objective) },
      vars: b.vars.map((v, i) => ({ cell: absolute(cellName(v.r, v.c)), name: label(v), integer: b.ints.has(i) })),
      constraints: b.rows
        .filter((row) => row.cell)
        .map((row) => {
          const value = b.read(row.cell!)
          const rhs = row.rhs!()
          const slack = clean(Math.abs(value - rhs))
          const cell = absolute(cellName(row.cell!.r, row.cell!.c))
          return { cell, name: label(row.cell!), value, formula: `${cell}${row.rel}${row.rhsText}`, binding: slack <= 1e-6 * Math.max(1, Math.abs(rhs)), slack }
        })
    }
    return { status, message: MESSAGES[status], cells: b.vars.map((v) => cellName(v.r, v.c)), original: b.start, values: good ? x : b.start, originalObjective, objective: objAt(), report }
  } finally {
    b.restore()
  }
}

/** Values within 1e-9 of a whole number (rounding noise from the simplex) become that number; −0 becomes 0. */
const clean = (v: number) => (Math.abs(v - Math.round(v)) < 1e-9 * Math.max(1, Math.abs(v)) ? Math.round(v) : v) + 0

// ---------- Excel's storage: hidden names on the sheet ----------

const REL_CODE: Record<SolverRel, number> = { '<=': 1, '=': 2, '>=': 3, int: 4, bin: 5 }
const REL_OF: Record<string, SolverRel> = { 1: '<=', 2: '=', 3: '>=', 4: 'int', 5: 'bin' }

const absolute = (text: string) =>
  text
    .split(',')
    .map((p) => p.trim().replace(/\$/g, '').replace(/([A-Za-z]+)(\d+)/g, (_, c: string, r: string) => `$${c.toUpperCase()}$${r}`))
    .join(',')

/** The model's names (name → formula text without "="), as Excel's Solver writes them. */
export function solverNames(m: SolverModel, sheetName: string): [string, string][] {
  const q = quoteSheet(sheetName)
  const ref = (t: string) =>
    absolute(t)
      .split(',')
      .map((p) => `${q}!${p}`)
      .join(',')
  const out: [string, string][] = [
    ['solver_opt', ref(m.objective)],
    ['solver_typ', String(m.goal === 'max' ? 1 : m.goal === 'min' ? 2 : 3)],
    ['solver_val', numberToText(m.value)],
    ['solver_adj', ref(m.changing)],
    ['solver_num', String(m.constraints.length)],
    ['solver_eng', m.method === 'simplex' ? '2' : '1'],
    ['solver_neg', m.nonNegative ? '1' : '2'],
    ['solver_ver', '3']
  ]
  m.constraints.forEach((k, i) => {
    const rhs = k.rel === 'int' ? '"integer"' : k.rel === 'bin' ? '"binary"' : Number.isFinite(Number(k.rhs)) && k.rhs.trim() !== '' ? numberToText(Number(k.rhs)) : ref(k.rhs.replace(/^=/, ''))
    out.push([`solver_lhs${i + 1}`, ref(k.lhs)], [`solver_rel${i + 1}`, String(REL_CODE[k.rel])], [`solver_rhs${i + 1}`, rhs])
  })
  return out
}

/** The model saved on a sheet (Excel's or ours), or null. */
export function readSolverModel(names: DefinedName[], sheet: number, sheetName: string): SolverModel | null {
  const get = (n: string) => names.find((d) => d.scope === sheet && d.name.toLowerCase() === n)?.text
  const opt = get('solver_opt')
  const adj = get('solver_adj')
  if (opt === undefined && adj === undefined) return null
  // References lose this sheet's prefix (the dialog shows them as $B$1).
  const local = (t: string | undefined) =>
    (t ?? '')
      .replace(/^=/, '')
      .split(',')
      .map((p) => {
        const bang = p.lastIndexOf('!')
        const name = bang > 0 ? p.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : null
        return name === null || name.toLowerCase() === sheetName.toLowerCase() ? p.slice(bang + 1) : p
      })
      .join(',')
  const m = emptySolverModel()
  m.objective = local(opt)
  m.changing = local(adj)
  m.goal = get('solver_typ') === '2' ? 'min' : get('solver_typ') === '3' ? 'value' : 'max'
  m.value = Number(get('solver_val') ?? 0) || 0
  m.method = get('solver_eng') === '2' ? 'simplex' : get('solver_eng') === undefined ? 'simplex' : 'nonlinear'
  m.nonNegative = get('solver_neg') !== '2'
  const count = Number(get('solver_num') ?? 0)
  for (let i = 1; i <= count; i++) {
    const rel = REL_OF[get(`solver_rel${i}`) ?? '']
    if (!rel) continue
    const rhs = (get(`solver_rhs${i}`) ?? '').replace(/^=/, '')
    m.constraints.push({ lhs: local(get(`solver_lhs${i}`)), rel, rhs: rel === 'int' || rel === 'bin' ? '' : /^-?[\d.]+(E[-+]?\d+)?$/i.test(rhs) ? rhs : local(rhs) })
  }
  return m
}

/** Replace the sheet's solver_* names with the model's (hidden, as Excel writes them). */
export function writeSolverModelTx(doc: Doc, tx: Tx, sheet: number, m: SolverModel) {
  const sheetName = doc.sheetName(sheet) ?? ''
  const fresh: DefinedName[] = solverNames(m, sheetName).map(([name, text]) => ({ name, scope: sheet, text, ast: doc.parse('=' + text, sheet, 0, 0), hidden: true }))
  const old = doc.wb.names.filter((d) => d.scope === sheet && /^solver_/i.test(d.name))
  tx.prop(doc.wb, 'names', [...doc.wb.names.filter((d) => !old.includes(d)), ...fresh])
  for (const d of [...old, ...fresh]) tx.namesChanged.add(d.name.toLowerCase())
}

// ---------- Answer Report ----------

/** Text to the left of a cell and above it, as Excel names report rows ("Profit Total"). */
function labelOf(doc: Doc, sheet: number, r: number, c: number): string {
  const sh = doc.sheet(sheet)
  let left = ''
  for (let cc = c - 1; cc >= 0 && !left; cc--) {
    const v = sh.get(r, cc)?.v
    if (typeof v === 'string' && v.trim()) left = v.trim()
  }
  let up = ''
  for (let rr = r - 1; rr >= 0 && !up; rr--) {
    const v = sh.get(rr, c)?.v
    if (typeof v === 'string' && v.trim()) up = v.trim()
  }
  return [left, up].filter(Boolean).join(' ')
}

/** Write an Answer Report on a new sheet (Excel's layout, in short). Returns the sheet's id. */
export function answerReportTx(doc: Doc, tx: Tx, sheet: number, m: SolverModel, res: SolverResult): number {
  let k = 1
  while (doc.wb.sheetByName(`Answer Report ${k}`)) k++
  const out = doc.wb.createSheet(`Answer Report ${k}`)
  const at = doc.wb.sheetIndex(sheet)
  tx.custom(
    () => doc.wb.sheets.splice(at, 0, out),
    () => doc.wb.sheets.splice(doc.wb.sheetIndex(out.id), 1)
  )
  tx.structural = true
  let r = 0
  const line = (...vs: (string | number)[]) => {
    vs.forEach((v, c) => doc.inputTx(tx, out.id, r, c + 1, typeof v === 'number' ? numberToText(v) : "'" + v))
    r++
  }
  const rep = res.report
  line('Simple Cell Answer Report')
  line(`Worksheet: ${doc.sheetName(sheet) ?? ''}`)
  line(`Result: ${res.message}`)
  line(`Method: ${m.method === 'simplex' ? 'Simplex LP' : 'Nonlinear (Nelder–Mead)'}`)
  r++
  line(`Objective Cell (${m.goal === 'max' ? 'Max' : m.goal === 'min' ? 'Min' : `Value Of ${numberToText(m.value)}`})`)
  line('Cell', 'Name', 'Original Value', 'Final Value')
  line(rep.objective.cell, rep.objective.name, res.originalObjective, res.objective)
  r++
  line('Variable Cells')
  line('Cell', 'Name', 'Original Value', 'Final Value', 'Integer')
  rep.vars.forEach((v, i) => line(v.cell, v.name, res.original[i], res.values[i], v.integer ? 'Integer' : 'Contin'))
  r++
  line('Constraints')
  line('Cell', 'Name', 'Cell Value', 'Formula', 'Status', 'Slack')
  for (const c of rep.constraints) line(c.cell, c.name, c.value, c.formula, c.binding ? 'Binding' : 'Not Binding', c.slack)
  return out.id
}

/** After Solve: keep the solution (one undoable step with the model and the report) or only save the model. */
export function finishSolver(doc: Doc, sheet: number, m: SolverModel, keep: SolverResult | null, report: SolverResult | null) {
  return doc.transact('Solver', (tx) => {
    writeSolverModelTx(doc, tx, sheet, m)
    if (keep)
      keep.cells.forEach((cn, i) => {
        const p = cellsOf(doc, sheet, cn)[0]
        doc.inputTx(tx, sheet, p.r, p.c, numberToText(keep.values[i]))
      })
    if (report) answerReportTx(doc, tx, sheet, m, report)
  })
}
