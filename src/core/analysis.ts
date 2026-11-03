// Data tools: Remove Duplicates, Text to Columns, Goal Seek, Descriptive Statistics.
import type { Doc } from './doc'
import type { Cell } from './model/workbook'
import { CellError, type Area, type Scalar } from './engine/values'
import { numberToText } from './engine/values'

const keyOf = (v: Scalar) => (v === null ? '' : typeof v === 'string' ? 's' + v.toLowerCase() : typeof v === 'number' ? 'n' + v : typeof v === 'boolean' ? 'b' + v : 'e' + (v as CellError).code)

/** Remove Duplicates: keep the first of each set of rows that match in `cols` (offsets in the area). */
export function removeDuplicates(doc: Doc, a: Area, cols: number[], headers: boolean): { removed: number; kept: number } {
  const sh = doc.sheet(a.sheet)
  const first = headers ? a.r0 + 1 : a.r0
  const last = Math.min(a.r1, sh.used().r)
  const width = a.c1 - a.c0 + 1
  const rows: (Cell | undefined)[][] = []
  for (let r = first; r <= last; r++) rows.push(Array.from({ length: width }, (_, j) => sh.get(r, a.c0 + j)))
  const seen = new Set<string>()
  const kept = rows.filter((row) => {
    const k = cols.map((j) => keyOf(row[j]?.v ?? null)).join('\u0000')
    if (seen.has(k)) return false
    seen.add(k)
    return true
  })
  const removed = rows.length - kept.length
  if (!removed) return { removed: 0, kept: kept.length }
  doc.transact('Remove duplicates', (tx) => {
    // Rows below move up within the range (cells outside its columns stay put, as in Excel).
    for (let i = 0; i < rows.length; i++)
      for (let j = 0; j < width; j++) {
        const cell = kept[i]?.[j]
        tx.setCell(a.sheet, first + i, a.c0 + j, cell ? { ...cell, f: cell.f ? { ...cell.f } : undefined } : undefined)
      }
  })
  return { removed, kept: kept.length }
}

/** Text to Columns (delimited): split the text of one column into the columns to its right. */
export function textToColumns(doc: Doc, a: Area, delims: string[], opts: { consecutive?: boolean; quote?: string } = {}): number {
  const sh = doc.sheet(a.sheet)
  if (a.c0 !== a.c1) throw new Error('Text to Columns works on one column at a time.')
  const q = opts.quote ?? '"'
  const split = (t: string): string[] => {
    const out: string[] = []
    let cur = ''
    let inQ = false
    for (let i = 0; i < t.length; i++) {
      const ch = t[i]
      if (q && ch === q) {
        if (inQ && t[i + 1] === q) {
          cur += q
          i++
        } else inQ = !inQ
      } else if (!inQ && delims.includes(ch)) {
        out.push(cur)
        cur = ''
        if (opts.consecutive) while (delims.includes(t[i + 1] ?? '')) i++
      } else cur += ch
    }
    out.push(cur)
    return out
  }
  const last = Math.min(a.r1, sh.used().r)
  let widest = 0
  const parts: (string[] | null)[] = []
  for (let r = a.r0; r <= last; r++) {
    const v = sh.get(r, a.c0)?.v
    const p = typeof v === 'string' ? split(v) : null
    parts.push(p)
    if (p) widest = Math.max(widest, p.length)
  }
  if (widest < 2) return 0
  doc.transact('Text to columns', (tx) => {
    parts.forEach((p, i) => {
      if (!p) return
      p.forEach((text, j) => doc.inputTx(tx, a.sheet, a.r0 + i, a.c0 + j, text.trim() === '' ? '' : text))
    })
  })
  return widest
}

/** Goal Seek: change `changing` until the formula in `target` equals `goal`. Returns the value found. */
export function goalSeek(doc: Doc, target: { sheet: number; r: number; c: number }, goal: number, changing: { sheet: number; r: number; c: number }): { ok: boolean; value: number; result: number } {
  const sh = doc.sheet(changing.sheet)
  const tc = doc.sheet(target.sheet).get(target.r, target.c)
  if (!tc?.f) throw new Error('The “Set cell” must contain a formula.')
  const cur = sh.get(changing.r, changing.c)
  if (cur?.f) throw new Error('The “By changing cell” must contain a value, not a formula.')
  const start = typeof cur?.v === 'number' ? cur.v : 0
  const original = cur ? { ...cur } : undefined
  const f = (x: number): number => {
    sh.set(changing.r, changing.c, { ...(cur ?? {}), v: x })
    doc.engine.touch(changing.sheet, changing.r, changing.c)
    doc.engine.recalc()
    const v = doc.sheet(target.sheet).get(target.r, target.c)?.v
    return typeof v === 'number' ? v - goal : NaN
  }
  // Secant steps, falling back to bisection once the goal is bracketed (Excel: 100 iterations, 0.001 change).
  let x0 = start
  let x1 = start === 0 ? 1 : start * 1.01
  let f0 = f(x0)
  let f1 = f(x1)
  const br: { lo: [number, number] | null; hi: [number, number] | null } = { lo: null, hi: null }
  const note = (x: number, y: number) => {
    if (Number.isNaN(y)) return
    if (y < 0 && (!br.lo || Math.abs(y) < Math.abs(br.lo[1]))) br.lo = [x, y]
    if (y > 0 && (!br.hi || y < br.hi[1])) br.hi = [x, y]
  }
  note(x0, f0)
  note(x1, f1)
  let found = Math.abs(f0) < 1e-9 ? x0 : NaN
  for (let i = 0; i < 100 && Number.isNaN(found); i++) {
    if (Math.abs(f1) < 1e-9) {
      found = x1
      break
    }
    let x2 = f1 !== f0 && Number.isFinite(f0) && Number.isFinite(f1) ? x1 - (f1 * (x1 - x0)) / (f1 - f0) : x1 * 2 + 1
    const { lo, hi } = br
    if (lo && hi && (!Number.isFinite(x2) || x2 < Math.min(lo[0], hi[0]) || x2 > Math.max(lo[0], hi[0]))) x2 = (lo[0] + hi[0]) / 2
    x0 = x1
    f0 = f1
    x1 = x2
    f1 = f(x1)
    note(x1, f1)
  }
  const ok = !Number.isNaN(found)
  // Restore, then make the change as one undoable edit (as Excel's OK does).
  sh.set(changing.r, changing.c, original)
  doc.engine.touch(changing.sheet, changing.r, changing.c)
  doc.engine.recalc()
  const value = ok ? found : x1
  if (ok) doc.input(changing.sheet, changing.r, changing.c, numberToText(value))
  const res = doc.sheet(target.sheet).get(target.r, target.c)?.v
  return { ok, value, result: typeof res === 'number' ? res : NaN }
}

/** Descriptive Statistics (Analysis ToolPak): the summary table for each column of the input. */
export function descriptiveStats(doc: Doc, input: Area, out: { sheet: number; r: number; c: number }, headers: boolean) {
  const sh = doc.sheet(input.sheet)
  const cols = input.c1 - input.c0 + 1
  const labels = ['Mean', 'Standard Error', 'Median', 'Mode', 'Standard Deviation', 'Sample Variance', 'Kurtosis', 'Skewness', 'Range', 'Minimum', 'Maximum', 'Sum', 'Count']
  return doc.transact('Descriptive statistics', (tx) => {
    for (let j = 0; j < cols; j++) {
      const c = input.c0 + j
      const xs: number[] = []
      for (let r = headers ? input.r0 + 1 : input.r0; r <= Math.min(input.r1, sh.used().r); r++) {
        const v = sh.get(r, c)?.v
        if (typeof v === 'number') xs.push(v)
      }
      const n = xs.length
      const mean = n ? xs.reduce((a, b) => a + b, 0) / n : NaN
      const dev = xs.map((x) => x - mean)
      const v2 = n > 1 ? dev.reduce((a, d) => a + d * d, 0) / (n - 1) : NaN
      const sd = Math.sqrt(v2)
      const sorted = [...xs].sort((a, b) => a - b)
      const median = n ? (n % 2 ? sorted[(n - 1) / 2] : (sorted[n / 2 - 1] + sorted[n / 2]) / 2) : NaN
      const counts = new Map<number, number>()
      for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1)
      const best = Math.max(0, ...counts.values())
      const mode = best > 1 ? xs.find((x) => counts.get(x) === best)! : null
      const kurt = n > 3 ? ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * dev.reduce((a, d) => a + (d / sd) ** 4, 0) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3)) : null
      const skew = n > 2 ? (n / ((n - 1) * (n - 2))) * dev.reduce((a, d) => a + (d / sd) ** 3, 0) : null
      const values: (number | string | null)[] = [mean, sd / Math.sqrt(n), median, mode ?? '#N/A', sd, v2, kurt ?? '#DIV/0!', skew ?? '#DIV/0!', n ? sorted[n - 1] - sorted[0] : 0, sorted[0] ?? 0, sorted[n - 1] ?? 0, xs.reduce((a, b) => a + b, 0), n]
      const x = out.c + j * 2
      const title = headers ? (doc.display(input.sheet, input.r0, c)?.t ?? `Column${j + 1}`) : `Column${j + 1}`
      doc.inputTx(tx, out.sheet, out.r, x, "'" + title)
      labels.forEach((l, i) => {
        doc.inputTx(tx, out.sheet, out.r + 2 + i, x, "'" + l)
        const v = values[i]
        doc.inputTx(tx, out.sheet, out.r + 2 + i, x + 1, v === null || (typeof v === 'number' && Number.isNaN(v)) ? '#DIV/0!' : typeof v === 'number' ? numberToText(v) : v)
      })
    }
  })
}
