// Sorting a range by one or more columns (Data → Sort, and the AutoFilter sort buttons).
import type { Doc } from './doc'
import { CellError, compareScalars, type Area, type Scalar } from './engine/values'
import type { Cell } from './model/workbook'

export interface SortKey {
  /** Absolute column (or row, for left-to-right sorts). */
  index: number
  order: 1 | -1
  /** Custom list order, e.g. ["Low","Medium","High"]. */
  list?: string[]
  by?: 'value' | 'fill' | 'fontColor'
}

export function sortRange(doc: Doc, a: Area, keys: SortKey[], header: boolean, leftToRight = false) {
  const sh = doc.sheet(a.sheet)
  const r0 = a.r0 + (header && !leftToRight ? 1 : 0)
  const c0 = a.c0 + (header && leftToRight ? 1 : 0)
  const lines: { idx: number; cells: (Cell | undefined)[]; keys: unknown[] }[] = []
  const count = leftToRight ? a.c1 - c0 + 1 : a.r1 - r0 + 1
  const width = leftToRight ? a.r1 - a.r0 + 1 : a.c1 - a.c0 + 1
  for (let i = 0; i < count; i++) {
    const cells: (Cell | undefined)[] = []
    for (let j = 0; j < width; j++) cells.push(leftToRight ? sh.get(a.r0 + j, c0 + i) : sh.get(r0 + i, a.c0 + j))
    const ks = keys.map((k) => {
      const cell = leftToRight ? sh.get(k.index, c0 + i) : sh.get(r0 + i, k.index)
      if (k.by === 'fill') return doc.wb.styles.get(cell?.s).fill ?? ''
      if (k.by === 'fontColor') return doc.wb.styles.get(cell?.s).color ?? ''
      return cell?.v ?? null
    })
    lines.push({ idx: i, cells, keys: ks })
  }
  lines.sort((x, y) => {
    for (let k = 0; k < keys.length; k++) {
      const p = x.keys[k] as Scalar
      const q = y.keys[k] as Scalar
      const key = keys[k]
      // Blanks always last, errors after values.
      if (p === null || q === null) {
        if (p === q) continue
        return p === null ? 1 : -1
      }
      if (p instanceof CellError || q instanceof CellError) {
        if (p instanceof CellError && q instanceof CellError) continue
        return p instanceof CellError ? 1 : -1
      }
      let c: number
      if (key.list) {
        const li = (v: unknown) => {
          const i = key.list!.findIndex((s) => typeof v === 'string' && s.toLowerCase() === v.toLowerCase())
          return i < 0 ? key.list!.length : i
        }
        c = li(p) - li(q)
      } else c = compareScalars(p, q)
      if (c) return c * key.order
    }
    return x.idx - y.idx
  })
  if (lines.every((l, i) => l.idx === i)) return
  doc.transact('Sort', (tx) => {
    lines.forEach((l, i) => {
      for (let j = 0; j < width; j++) {
        const r = leftToRight ? a.r0 + j : r0 + i
        const c = leftToRight ? c0 + i : a.c0 + j
        const cell = l.cells[j]
        tx.setCell(a.sheet, r, c, cell && cell.spillOf === undefined ? { ...cell } : cell?.f ? { ...cell } : cell ? { v: cell.v, s: cell.s } : undefined)
      }
    })
  })
}
