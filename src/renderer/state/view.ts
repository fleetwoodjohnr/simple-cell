// Per-tab view state: active sheet, selection, editing, scroll position.
import { create } from 'zustand'
import { MAX_COLS, MAX_ROWS, type Area } from '@/core/engine/values'

export interface Rect {
  r0: number
  c0: number
  r1: number
  c1: number
}

export interface Sel {
  /** Ranges in the order they were selected; the last one holds the active cell. */
  ranges: Rect[]
  /** The active cell (white cell inside the selection). */
  active: { r: number; c: number }
  /** Moving corner of the last range (Shift+Arrow extends it; the anchor is the active cell). */
  focus: { r: number; c: number }
}

export interface Editing {
  text: string
  /** Where the user is typing: in the cell, or in the formula bar. */
  where: 'cell' | 'bar'
  /** Started by typing (arrow keys commit and move) rather than F2 (arrow keys move the caret). */
  enterMode: boolean
  caret: number
  /** Formula point mode: a reference being inserted with the keyboard or mouse. */
  point?: { start: number; end: number; anchor: { r: number; c: number }; focus: { r: number; c: number } }
}

export interface TabView {
  sheet: number
  sel: Sel
  editing: Editing | null
  /** Marching ants around the copied range. */
  copy: (Area & { cut: boolean }) | null
  /** Scroll offsets per sheet (pixels at 100% zoom). */
  scroll: Record<number, { x: number; y: number }>
  zoom: number
  /** The active cell's validation dropdown is open (Alt+Down). */
  listOpen?: boolean
  /** Circle Invalid Data: cells marked on the current sheet. */
  circles?: { r: number; c: number }[]
  /** The open AutoFilter menu: filter range key and column offset. */
  filterMenu?: { key: string; col: number }
}

export const rectOf = (a: { r: number; c: number }, b: { r: number; c: number }): Rect => ({ r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) })

export const cellSel = (r: number, c: number): Sel => ({ ranges: [{ r0: r, c0: c, r1: r, c1: c }], active: { r, c }, focus: { r, c } })

export function freshView(sheet: number): TabView {
  return { sheet, sel: cellSel(0, 0), editing: null, copy: null, scroll: {}, zoom: 1 }
}

interface ViewStore {
  views: Record<string, TabView>
}

export const useViews = create<ViewStore>(() => ({ views: {} }))

export function getView(tab: string): TabView {
  return useViews.getState().views[tab]
}

export function setView(tab: string, patch: Partial<TabView> | ((v: TabView) => Partial<TabView>)) {
  useViews.setState((s) => {
    const cur = s.views[tab]
    if (!cur) return s
    const p = typeof patch === 'function' ? patch(cur) : patch
    return { views: { ...s.views, [tab]: { ...cur, ...p } } }
  })
}

export function initView(tab: string, sheet: number) {
  useViews.setState((s) => ({ views: { ...s.views, [tab]: freshView(sheet) } }))
}

export function dropView(tab: string) {
  useViews.setState((s) => {
    const { [tab]: _gone, ...rest } = s.views
    return { views: rest }
  })
}

export const clampCell = (r: number, c: number) => ({ r: Math.max(0, Math.min(MAX_ROWS - 1, r)), c: Math.max(0, Math.min(MAX_COLS - 1, c)) })

/** Selection areas on a sheet, for core calls. */
export function areasOf(v: TabView): Area[] {
  return v.sel.ranges.map((r) => ({ sheet: v.sheet, ...r }))
}

export const isWholeCols = (r: Rect) => r.r0 === 0 && r.r1 === MAX_ROWS - 1
export const isWholeRows = (r: Rect) => r.c0 === 0 && r.c1 === MAX_COLS - 1
