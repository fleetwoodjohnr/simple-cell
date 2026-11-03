// Ctrl+Arrow navigation: jump to the edge of the current block of data, or to the next block.
import { MAX_COLS, MAX_ROWS } from './engine/values'
import type { Sheet } from './model/workbook'

export type Dir = 'up' | 'down' | 'left' | 'right' | 'end' | 'home'

const filled = (s: Sheet, r: number, c: number) => {
  const cell = s.get(r, c)
  return !!cell && cell.v !== null && cell.v !== ''
}

export function navigate(s: Sheet, r: number, c: number, dir: Dir): { r: number; c: number } {
  if (dir === 'home') return { r: s.freeze.rows, c: s.freeze.cols }
  const u = s.used()
  if (dir === 'end') return { r: Math.max(0, u.r), c: Math.max(0, u.c) }
  const dr = dir === 'up' ? -1 : dir === 'down' ? 1 : 0
  const dc = dir === 'left' ? -1 : dir === 'right' ? 1 : 0
  const inside = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS
  const hidden = (rr: number, cc: number) => (dr ? s.rowHeight(rr) === 0 : s.colWidth(cc) === 0)
  const step = (rr: number, cc: number): [number, number] => {
    let nr = rr + dr
    let nc = cc + dc
    while (inside(nr, nc) && hidden(nr, nc)) {
      nr += dr
      nc += dc
    }
    return [nr, nc]
  }
  let [nr, nc] = step(r, c)
  if (!inside(nr, nc)) return { r, c }
  const edge = { r: dr > 0 ? MAX_ROWS - 1 : dr < 0 ? 0 : r, c: dc > 0 ? MAX_COLS - 1 : dc < 0 ? 0 : c }
  if (filled(s, r, c) && filled(s, nr, nc)) {
    // Run to the last filled cell of this block.
    let pr = nr
    let pc = nc
    for (;;) {
      const [qr, qc] = step(pr, pc)
      if (!inside(qr, qc) || !filled(s, qr, qc)) return { r: pr, c: pc }
      pr = qr
      pc = qc
    }
  }
  // Skip empty cells to the next filled one; beyond the used range there is nothing more.
  for (;;) {
    if (filled(s, nr, nc)) return { r: nr, c: nc }
    if ((dr > 0 && nr > u.r) || (dc > 0 && nc > u.c)) return edge
    const [qr, qc] = step(nr, nc)
    if (!inside(qr, qc)) return { r: nr, c: nc }
    nr = qr
    nc = qc
  }
}
