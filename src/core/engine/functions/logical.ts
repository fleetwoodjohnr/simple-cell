// Logical functions, LET, LAMBDA and the LAMBDA helper functions.
import { defineFunctions, evaluate, invokeLambda, toArray, toScalar, type EvalContext } from '../evaluate'
import type { Node } from '../parser'
import { CellError, ERR, compareScalars, isArray, isError, isLambda, isRef, makeArray, toBool, type LambdaValue, type Scalar, type Value } from '../values'
import { allValues, fn, MANY } from './helpers'

/** Evaluate a lazily passed argument to one value (or array, for array-returning branches). */
const ev = (n: Node | undefined, ctx: EvalContext): Value => (n === undefined ? null : evaluate(n, ctx))

/** Scalar condition for IF-like functions, or an array when the test is an array (IF lifts over it). */
function condition(v: Value, ctx: EvalContext): Scalar | Scalar[] | { arr: ReturnType<typeof toArray> } {
  if (isRef(v) && (v.areas.length > 1 || v.areas[0].r0 !== v.areas[0].r1 || v.areas[0].c0 !== v.areas[0].c1) && !ctx.legacy) return { arr: toArray(v, ctx) }
  if (isArray(v) && v.rows * v.cols > 1) return { arr: v }
  return toScalar(v, ctx)
}

function logicalValues(args: Value[], ctx: EvalContext): boolean[] | CellError {
  const out: boolean[] = []
  for (const v of args) {
    if (isRef(v) || isArray(v)) {
      const vals = allValues([v], ctx)
      if (isError(vals)) return vals
      for (const x of vals) {
        if (x instanceof CellError) return x
        if (typeof x === 'number') out.push(x !== 0)
        else if (typeof x === 'boolean') out.push(x)
      }
    } else {
      const b = toBool(v as Scalar)
      if (isError(b)) return b
      out.push(b)
    }
  }
  return out.length ? out : ERR.VALUE
}

function lambdaOf(v: Value): LambdaValue | CellError {
  if (isError(v)) return v
  return isLambda(v) ? v : ERR.VALUE
}

function mapArrays(arrays: Value[], f: LambdaValue, ctx: EvalContext): Value {
  const arrs = arrays.map((a) => toArray(a, ctx))
  for (const a of arrs) if (isError(a)) return a
  const list = arrs as Exclude<(typeof arrs)[number], CellError>[]
  const rows = Math.max(...list.map((a) => a.rows))
  const cols = Math.max(...list.map((a) => a.cols))
  if (list.some((a) => (a.rows !== rows && a.rows !== 1) || (a.cols !== cols && a.cols !== 1))) return ERR.VALUE
  const data: Scalar[] = []
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      const args = list.map((a) => a.data[(a.rows === 1 ? 0 : r) * a.cols + (a.cols === 1 ? 0 : c)])
      data.push(single(invokeLambda(f, args, ctx), ctx))
    }
  return makeArray(rows, cols, data)
}

/** A LAMBDA helper's per-item result must be one value. */
function single(v: Value, ctx: EvalContext): Scalar {
  if (isArray(v)) return v.rows * v.cols === 1 ? v.data[0] : ERR.CALC
  if (isRef(v)) {
    const a = v.areas[0]
    if (v.areas.length === 1 && a.r0 === a.r1 && a.c0 === a.c1) return ctx.read(a.sheet, a.r0, a.c0) ?? 0
    return ERR.CALC
  }
  if (isLambda(v)) return ERR.CALC
  return v ?? 0
}

defineFunctions([
  fn('TRUE', 0, 0, '', () => true),
  fn('FALSE', 0, 0, '', () => false),
  fn('NOT', 1, 1, 'b', ([b]) => !b),
  fn('AND', 1, MANY, 'r', (args, ctx) => {
    const bs = logicalValues(args, ctx)
    return isError(bs) ? bs : bs.every(Boolean)
  }),
  fn('OR', 1, MANY, 'r', (args, ctx) => {
    const bs = logicalValues(args, ctx)
    return isError(bs) ? bs : bs.some(Boolean)
  }),
  fn('XOR', 1, MANY, 'r', (args, ctx) => {
    const bs = logicalValues(args, ctx)
    return isError(bs) ? bs : bs.filter(Boolean).length % 2 === 1
  }),
  {
    name: 'IF',
    min: 1,
    max: 3,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => {
      const c = condition(ev(args[0], ctx), ctx)
      const pick = (test: Scalar): Value => {
        if (test instanceof CellError) return test
        const b = toBool(test)
        if (isError(b)) return b
        if (b) return args[1] === undefined ? true : args[1].t === 'missing' ? 0 : ev(args[1], ctx)
        return args.length < 3 ? false : args[2].t === 'missing' ? 0 : ev(args[2], ctx)
      }
      if (typeof c === 'object' && c !== null && 'arr' in c) {
        const arr = c.arr
        if (isError(arr)) return arr
        // Lift over the test: both branches are evaluated once and indexed per item.
        const yes = args[1] === undefined ? true : ev(args[1], ctx)
        const no = args.length < 3 ? false : ev(args[2], ctx)
        const ya = toArray(yes, ctx)
        const na = toArray(no, ctx)
        const data = arr.data.map((t, i) => {
          if (t instanceof CellError) return t
          const b = toBool(t)
          if (isError(b)) return b
          const src = b ? ya : na
          if (isError(src)) return src
          const r = Math.floor(i / arr.cols)
          const cc = i % arr.cols
          const rr = src.rows === 1 ? 0 : r
          const ci = src.cols === 1 ? 0 : cc
          return rr < src.rows && ci < src.cols ? (src.data[rr * src.cols + ci] ?? 0) : ERR.NA
        })
        return makeArray(arr.rows, arr.cols, data)
      }
      return pick(c as Scalar)
    }
  },
  {
    name: 'IFS',
    min: 2,
    max: 254,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => {
      if (args.length % 2) return ERR.VALUE
      for (let i = 0; i < args.length; i += 2) {
        const t = toScalar(ev(args[i], ctx), ctx)
        if (t instanceof CellError) return t
        const b = toBool(t)
        if (isError(b)) return b
        if (b) return ev(args[i + 1], ctx)
      }
      return ERR.NA
    }
  },
  {
    name: 'IFERROR',
    min: 2,
    max: 2,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => replaceErrors(ev(args[0], ctx), () => ev(args[1], ctx), ctx, () => true)
  },
  {
    name: 'IFNA',
    min: 2,
    max: 2,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => replaceErrors(ev(args[0], ctx), () => ev(args[1], ctx), ctx, (e) => e === ERR.NA)
  },
  {
    name: 'SWITCH',
    min: 3,
    max: 254,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => {
      const v = toScalar(ev(args[0], ctx), ctx)
      if (v instanceof CellError) return v
      const pairs = args.length - 1
      for (let i = 1; i + 1 < args.length; i += 2) {
        const m = toScalar(ev(args[i], ctx), ctx)
        if (m instanceof CellError) return m
        if (compareScalars(v, m) === 0 && typeof v === typeof m) return ev(args[i + 1], ctx)
      }
      return pairs % 2 === 1 ? ev(args[args.length - 1], ctx) : ERR.NA
    }
  },
  {
    name: 'LET',
    min: 3,
    max: 253,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => {
      if (args.length % 2 === 0) return ERR.VALUE
      const env = new Map(ctx.env ?? [])
      const inner: EvalContext = { ...ctx, env }
      for (let i = 0; i + 1 < args.length; i += 2) {
        const nameNode = args[i]
        if (nameNode.t !== 'name' || nameNode.sheet) return ERR.VALUE
        env.set(nameNode.name.toUpperCase(), evaluate(args[i + 1], inner))
      }
      return evaluate(args[args.length - 1], inner)
    }
  },
  {
    name: 'LAMBDA',
    min: 1,
    max: 254,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => {
      const params: string[] = []
      for (const p of args.slice(0, -1)) {
        if (p.t !== 'name' || p.sheet) return ERR.VALUE
        params.push(p.name.toUpperCase())
      }
      if (new Set(params).size !== params.length) return ERR.VALUE
      return { kind: 'lambda', params, body: args[args.length - 1], env: ctx.env }
    }
  },
  {
    name: 'ISOMITTED',
    min: 1,
    max: 1,
    params: 'r',
    lazy: true,
    fn: (args: Node[], ctx: EvalContext): Value => {
      const a = args[0]
      if (a.t !== 'name') return false
      return ctx.env?.has(a.name.toUpperCase()) ? ctx.env.get(a.name.toUpperCase()) === null : false
    }
  },
  fn('MAP', 2, MANY, 'r', (args, ctx) => {
    const f = lambdaOf(args[args.length - 1])
    if (isError(f)) return f
    return mapArrays(args.slice(0, -1), f, ctx)
  }),
  fn('REDUCE', 2, 3, 'r', (args, ctx) => {
    const [init, arr, lam] = args.length === 3 ? args : [null, args[0], args[1]]
    const f = lambdaOf(lam)
    if (isError(f)) return f
    const a = toArray(arr, ctx)
    if (isError(a)) return a
    let acc: Value = isRef(init) ? toScalar(init, ctx) : init
    for (const x of a.data) acc = invokeLambda(f, [acc, x], ctx)
    return acc
  }),
  fn('SCAN', 2, 3, 'r', (args, ctx) => {
    const [init, arr, lam] = args.length === 3 ? args : [null, args[0], args[1]]
    const f = lambdaOf(lam)
    if (isError(f)) return f
    const a = toArray(arr, ctx)
    if (isError(a)) return a
    let acc: Value = isRef(init) ? toScalar(init, ctx) : init
    const data: Scalar[] = []
    for (const x of a.data) {
      acc = invokeLambda(f, [acc, x], ctx)
      data.push(single(acc, ctx))
    }
    return makeArray(a.rows, a.cols, data)
  }),
  fn('BYROW', 2, 2, 'r', ([arr, lam], ctx) => {
    const f = lambdaOf(lam)
    if (isError(f)) return f
    const a = toArray(arr, ctx)
    if (isError(a)) return a
    const data: Scalar[] = []
    for (let r = 0; r < a.rows; r++) data.push(single(invokeLambda(f, [makeArray(1, a.cols, a.data.slice(r * a.cols, (r + 1) * a.cols))], ctx), ctx))
    return makeArray(a.rows, 1, data)
  }),
  fn('BYCOL', 2, 2, 'r', ([arr, lam], ctx) => {
    const f = lambdaOf(lam)
    if (isError(f)) return f
    const a = toArray(arr, ctx)
    if (isError(a)) return a
    const data: Scalar[] = []
    for (let c = 0; c < a.cols; c++) {
      const col: Scalar[] = []
      for (let r = 0; r < a.rows; r++) col.push(a.data[r * a.cols + c])
      data.push(single(invokeLambda(f, [makeArray(a.rows, 1, col)], ctx), ctx))
    }
    return makeArray(1, a.cols, data)
  }),
  fn('MAKEARRAY', 3, 3, 'nnr', ([rows, cols, lam], ctx) => {
    const f = lambdaOf(lam)
    if (isError(f)) return f
    rows = Math.trunc(rows)
    cols = Math.trunc(cols)
    if (rows < 1 || cols < 1) return ERR.VALUE
    const data: Scalar[] = []
    for (let r = 1; r <= rows; r++) for (let c = 1; c <= cols; c++) data.push(single(invokeLambda(f, [r, c], ctx), ctx))
    return makeArray(rows, cols, data)
  })
])

function replaceErrors(v: Value, alt: () => Value, ctx: EvalContext, which: (e: CellError) => boolean): Value {
  if (isRef(v) || isArray(v)) {
    const multi = isArray(v) ? v.rows * v.cols > 1 : v.areas.length > 1 || v.areas[0].r0 !== v.areas[0].r1 || v.areas[0].c0 !== v.areas[0].c1
    if (multi && !ctx.legacy) {
      const a = toArray(v, ctx)
      if (isError(a)) return which(a) ? alt() : a
      if (!a.data.some((x) => x instanceof CellError && which(x))) return v
      const fallback = toScalar(alt(), ctx)
      return makeArray(a.rows, a.cols, a.data.map((x) => (x instanceof CellError && which(x) ? fallback : x)))
    }
    const s = toScalar(v, ctx)
    return s instanceof CellError && which(s) ? alt() : v
  }
  if (v instanceof CellError && which(v)) return alt()
  if (isLambda(v)) return ERR.CALC
  return v === null ? 0 : v
}
