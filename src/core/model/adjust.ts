// Reference adjustment when rows or columns are inserted or deleted, or cells are moved.
import { mapNode, refCorner, type Node, type RefNode } from '../engine/parser'
import { MAX_COLS, MAX_ROWS, type Area } from '../engine/values'

export interface Shift {
  sheet: number
  axis: 'row' | 'col'
  /** First row/column affected. */
  at: number
  /** Positive: inserted; negative: deleted. */
  count: number
  /**
   * Insert/Delete Cells: only this band of the other axis moves (columns for a row shift). Ranges move or grow
   * only when they lie entirely inside the band; others stay as they are (as in Excel).
   */
  span?: [number, number]
}

/** New [lo, hi] of an index span after the shift, or null when it is deleted entirely. */
export function shiftSpan(lo: number, hi: number, at: number, count: number): [number, number] | null {
  if (count > 0) {
    return [lo >= at ? lo + count : lo, hi >= at ? hi + count : hi]
  }
  const n = -count
  const end = at + n - 1
  const nlo = lo > end ? lo - n : lo >= at ? at : lo
  const nhi = hi > end ? hi - n : hi >= at ? at - 1 : hi
  if (nhi < nlo) return null
  return [nlo, nhi]
}

export function shiftArea(a: Area, s: Shift): Area | null {
  if (a.sheet !== s.sheet) return a
  if (s.span) {
    const [lo, hi] = s.axis === 'row' ? [a.c0, a.c1] : [a.r0, a.r1]
    if (lo < s.span[0] || hi > s.span[1]) return a
  }
  if (s.axis === 'row') {
    if (a.r0 === 0 && a.r1 === MAX_ROWS - 1) return a
    const r = shiftSpan(a.r0, a.r1, s.at, s.count)
    if (!r) return null
    if (r[1] >= MAX_ROWS) r[1] = MAX_ROWS - 1
    if (r[0] >= MAX_ROWS) return null
    return { ...a, r0: r[0], r1: r[1] }
  }
  if (a.c0 === 0 && a.c1 === MAX_COLS - 1) return a
  const c = shiftSpan(a.c0, a.c1, s.at, s.count)
  if (!c) return null
  if (c[1] >= MAX_COLS) c[1] = MAX_COLS - 1
  if (c[0] >= MAX_COLS) return null
  return { ...a, c0: c[0], c1: c[1] }
}

/**
 * Rewrite a formula for a structural change. `from` is the formula's position before the change, `to` after it
 * (they differ when the formula cell itself moved). `sheet` is the formula's sheet id.
 * Returns the same object when nothing changed.
 */
export function adjustFormula(ast: Node, sheet: number, from: { row: number; col: number }, to: { row: number; col: number }, s: Shift): Node {
  return mapNode(ast, (n) => {
    if (n.t !== 'ref') return undefined
    const refSheet = n.sheet ? n.sheet.id : sheet
    const a = refCorner(n, from)
    if (!a) return undefined
    // 3-D references are only adjusted when the change is on their first sheet (Excel adjusts all; rare in practice).
    const target = refSheet === s.sheet && !n.sheet?.id2 ? shiftArea({ sheet: refSheet, ...a }, s) : { sheet: refSheet, ...a }
    if (!target) return { t: 'err', v: '#REF!' }
    return rebase(n, target, to, from)
  })
}

/** The node for absolute area `a`, placed in a formula at `to` (relative parts become offsets from `to`). */
function rebase(n: RefNode, a: Area, to: { row: number; col: number }, from: { row: number; col: number }): Node | undefined {
  const old = refCorner(n, from)!
  const same = old.r0 === a.r0 && old.r1 === a.r1 && old.c0 === a.c0 && old.c1 === a.c1 && from.row === to.row && from.col === to.col
  if (same) return undefined
  // Keep which corner is "first" as written.
  const firstTop = (n.ra0 ? n.r0 : n.r0 + from.row) <= (n.ra1 ? n.r1 : n.r1 + from.row)
  const firstLeft = (n.ca0 ? n.c0 : n.c0 + from.col) <= (n.ca1 ? n.c1 : n.c1 + from.col)
  const r0 = firstTop ? a.r0 : a.r1
  const r1 = firstTop ? a.r1 : a.r0
  const c0 = firstLeft ? a.c0 : a.c1
  const c1 = firstLeft ? a.c1 : a.c0
  return {
    ...n,
    r0: n.ra0 ? r0 : r0 - to.row,
    r1: n.ra1 ? r1 : r1 - to.row,
    c0: n.ca0 ? c0 : c0 - to.col,
    c1: n.ca1 ? c1 : c1 - to.col
  }
}

/**
 * Rewrite references into a moved block (cut and paste, drag): references to cells inside `src` now point into `dst`.
 * `pos` is the formula's own position (the formula didn't move, or `to` gives its new position).
 */
export function adjustForMove(ast: Node, sheet: number, from: { row: number; col: number }, to: { row: number; col: number }, src: Area, dst: { sheet: number; r0: number; c0: number }): Node {
  const dr = dst.r0 - src.r0
  const dc = dst.c0 - src.c0
  return mapNode(ast, (n) => {
    if (n.t !== 'ref') return undefined
    const refSheet = n.sheet ? n.sheet.id : sheet
    const a = refCorner(n, from)
    if (!a) return undefined
    const inside = refSheet === src.sheet && a.r0 >= src.r0 && a.r1 <= src.r1 && a.c0 >= src.c0 && a.c1 <= src.c1
    const target: Area = inside ? { sheet: dst.sheet, r0: a.r0 + dr, r1: a.r1 + dr, c0: a.c0 + dc, c1: a.c1 + dc } : { sheet: refSheet, ...a }
    if (target.r1 >= MAX_ROWS || target.c1 >= MAX_COLS) return { t: 'err', v: '#REF!' }
    const moved = rebase(n, target, to, from)
    if (inside && dst.sheet !== refSheet) return { ...((moved ?? n) as RefNode), sheet: { id: dst.sheet } }
    return moved
  })
}

/** Paste Special ▸ Transpose: relative references swap their row and column offsets (absolute parts stay). */
export function transposeRefs(ast: Node): Node {
  return mapNode(ast, (n) => {
    if (n.t !== 'ref' || n.span) return undefined
    const corner = (r: number, c: number, ra: boolean, ca: boolean) => (!ra && !ca ? [c, r] : [r, c])
    const [r0, c0] = corner(n.r0, n.c0, n.ra0, n.ca0)
    const [r1, c1] = corner(n.r1, n.c1, n.ra1, n.ca1)
    return { ...n, r0, c0, r1, c1 }
  })
}

/** Rewrite a formula copied from `from` to `to`: relative references move; those falling off the sheet become #REF!. */
export function relocate(ast: Node, to: { row: number; col: number }): Node {
  return mapNode(ast, (n) => (n.t === 'ref' && !refCorner(n, to) ? { t: 'err', v: '#REF!' } : undefined))
}
