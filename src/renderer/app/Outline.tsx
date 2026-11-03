// Outline (grouped rows/columns): bars, +/− buttons and level buttons in the gutters left of and above the headers.
import { useMemo } from 'react'
import { GUTTER_STEP, colX, rowY, visibleCols, visibleRows, type Layout } from '@/grid/layout'
import type { Book } from '../state/client'
import { useViews } from '../state/view'
import * as A from './actions'

/** Runs of consecutive indexes at each outline level (level >= L). */
export function outlineRuns(list: [number, number][]): { level: number; s: number; e: number }[] {
  const sorted = [...list].sort((a, b) => a[0] - b[0])
  const max = sorted.reduce((m, [, l]) => Math.max(m, l), 0)
  const out: { level: number; s: number; e: number }[] = []
  for (let level = 1; level <= max; level++) {
    let s = -1
    let prev = -2
    for (const [i, l] of sorted) {
      if (l < level) continue
      if (i !== prev + 1) {
        if (s >= 0) out.push({ level, s, e: prev })
        s = i
      }
      prev = i
    }
    if (s >= 0) out.push({ level, s, e: prev })
  }
  return out
}

const btn = 'absolute flex items-center justify-center rounded-[3px] border border-border-strong bg-panel text-[10px] font-semibold leading-none text-text-2 hover:bg-hover'

export function OutlineOverlay({ tab, book, L }: { tab: string; book: Book; L: Layout }) {
  const view = useViews((s) => s.views[tab])
  const meta = book.sheetMeta(view.sheet)
  const rowRuns = useMemo(() => outlineRuns(meta?.outlineRows ?? []), [meta?.outlineRows])
  const colRuns = useMemo(() => outlineRuns(meta?.outlineCols ?? []), [meta?.outlineCols])
  if (!meta || (!L.gutterW && !L.gutterH)) return null
  const z = L.zoom
  const act = (fn: (c: A.Ctx) => Promise<unknown>) => {
    const c = A.ctx()
    if (c) void A.edit(c, () => fn(c))
  }
  const [vr0, vr1] = visibleRows(L)
  const [vc0, vc1] = visibleCols(L)
  const rowLevels = rowRuns.reduce((m, r) => Math.max(m, r.level), 0)
  const colLevels = colRuns.reduce((m, r) => Math.max(m, r.level), 0)
  const hiddenRow = (r: number) => L.rows.size(r) === 0
  const hiddenCol = (c: number) => L.cols.size(c) === 0
  const allHidden = (s: number, e: number, hidden: (i: number) => boolean) => {
    for (let i = s; i <= e; i++) if (!hidden(i)) return false
    return true
  }
  const size = 12
  return (
    <>
      {L.gutterW > 0 && (
        <div className="pointer-events-none absolute left-0 overflow-hidden" style={{ top: L.headerH, width: L.gutterW, height: L.height - L.headerH }}>
          {rowRuns
            .filter((g) => g.e + 1 >= Math.min(vr0, 0) && g.s <= vr1 + 1)
            .map((g) => {
              const x = (g.level - 1) * GUTTER_STEP + 2
              const collapsed = allHidden(g.s, g.e, hiddenRow)
              const top = rowY(L, g.s) - L.headerH
              const bottom = rowY(L, g.e + 1) - L.headerH
              const sumH = L.rows.size(g.e + 1) * z
              const by = bottom + Math.max(0, (sumH - size) / 2)
              return (
                <div key={`r${g.level}:${g.s}`}>
                  {!collapsed && <div className="absolute border-l-2 border-t-2 border-text-3" style={{ left: x + size / 2 - 1, top: top + 2, width: 4, height: Math.max(0, bottom - top - 2) }} />}
                  <button
                    type="button"
                    aria-label={`${collapsed ? 'Show' : 'Hide'} rows ${g.s + 1}–${g.e + 1}`}
                    data-testid={`outline-row-${g.s + 1}-${g.e + 1}`}
                    className={`${btn} pointer-events-auto`}
                    style={{ left: x, top: by, width: size, height: size }}
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => act((c) => c.book.api.showDetail(c.view.sheet, 'row', g.s, g.e, collapsed))}
                  >
                    {collapsed ? '+' : '−'}
                  </button>
                </div>
              )
            })}
        </div>
      )}
      {L.gutterH > 0 && (
        <div className="pointer-events-none absolute top-0 overflow-hidden" style={{ left: L.headerW, width: L.width - L.headerW, height: L.gutterH }}>
          {colRuns
            .filter((g) => g.e + 1 >= Math.min(vc0, 0) && g.s <= vc1 + 1)
            .map((g) => {
              const y = (g.level - 1) * GUTTER_STEP + 2
              const collapsed = allHidden(g.s, g.e, hiddenCol)
              const left = colX(L, g.s) - L.headerW
              const right = colX(L, g.e + 1) - L.headerW
              const bx = right + Math.max(0, (L.cols.size(g.e + 1) * z - size) / 2)
              return (
                <div key={`c${g.level}:${g.s}`}>
                  {!collapsed && <div className="absolute border-l-2 border-t-2 border-text-3" style={{ top: y + size / 2 - 1, left: left + 2, height: 4, width: Math.max(0, right - left - 2) }} />}
                  <button
                    type="button"
                    aria-label={`${collapsed ? 'Show' : 'Hide'} columns`}
                    data-testid={`outline-col-${g.s}-${g.e}`}
                    className={`${btn} pointer-events-auto`}
                    style={{ top: y, left: bx, width: size, height: size }}
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => act((c) => c.book.api.showDetail(c.view.sheet, 'col', g.s, g.e, collapsed))}
                  >
                    {collapsed ? '+' : '−'}
                  </button>
                </div>
              )
            })}
        </div>
      )}
      {/* Level buttons: rows in the corner's left part, columns in its top part. */}
      {Array.from({ length: rowLevels ? rowLevels + 1 : 0 }, (_, i) => (
        <button key={`rl${i}`} type="button" aria-label={`Show row level ${i + 1}`} className={btn} style={{ left: i * GUTTER_STEP + 2, top: Math.max(L.gutterH, L.headerH - size - 4), width: size, height: size }} onMouseDown={(e) => e.preventDefault()} onClick={() => act((c) => c.book.api.showOutlineLevel(c.view.sheet, 'row', i + 1))}>
          {i + 1}
        </button>
      ))}
      {Array.from({ length: colLevels ? colLevels + 1 : 0 }, (_, i) => (
        <button key={`cl${i}`} type="button" aria-label={`Show column level ${i + 1}`} className={btn} style={{ top: i * GUTTER_STEP + 2, left: Math.max(L.gutterW, L.headerW - size - 4), width: size, height: size }} onMouseDown={(e) => e.preventDefault()} onClick={() => act((c) => c.book.api.showOutlineLevel(c.view.sheet, 'col', i + 1))}>
          {i + 1}
        </button>
      ))}
    </>
  )
}
