// Plain-JSON copy of a workbook, for crash recovery and version history. Formulas are stored as text.
import { parseFormula, printFormula } from './engine/parser'
import { CellError, type Scalar } from './engine/values'
import { Sheet, Workbook, type Cell } from './model/workbook'
import type { Style } from './model/styles'

type JScalar = number | string | boolean | null | { e: string }
export interface Snapshot {
  v: 1
  date1904: boolean
  calc: Workbook['calc']
  styles: Style[]
  names: { name: string; scope: number | null; text: string; comment?: string; hidden?: boolean }[]
  tables: Workbook['tables']
  props: Record<string, string>
  activeSheet: number
  sheets: {
    id: number
    name: string
    hidden: Sheet['hidden']
    tabColor: string | null
    freeze: Sheet['freeze']
    merges: Sheet['merges']
    defaultColWidth: number
    defaultRowHeight: number
    showGridlines: boolean
    zoom: number
    cols: [number, unknown][]
    rows: [number, unknown][]
    /** [row, col, value, style, formula, array area] */
    cells: [number, number, JScalar, number, string?, unknown?][]
  }[]
}

const enc = (v: Scalar): JScalar => (v instanceof CellError ? { e: v.code } : v)
const dec = (v: JScalar): Scalar => (v !== null && typeof v === 'object' ? CellError.of(v.e as never) : v)

export function snapshot(wb: Workbook): Snapshot {
  const name = (id: number) => wb.sheetById(id)?.name ?? null
  return {
    v: 1,
    date1904: wb.date1904,
    calc: wb.calc,
    styles: wb.styles.list,
    names: wb.names.map((n) => ({ name: n.name, scope: n.scope, text: n.text, comment: n.comment, hidden: n.hidden })),
    tables: wb.tables.map((t) => ({ ...t, carry: undefined })),
    props: wb.props,
    activeSheet: wb.activeSheet,
    sheets: wb.sheets.map((s) => {
      const cells: Snapshot['sheets'][number]['cells'] = []
      for (const [r, row] of s.rows)
        for (const [c, cell] of row) {
          if (cell.spillOf !== undefined && !cell.f) {
            if (cell.s) cells.push([r, c, null, cell.s])
            continue
          }
          const f = cell.f ? printFormula(cell.f.ast, { row: r, col: c, sheetName: name }) : undefined
          cells.push(f ? [r, c, enc(cell.v), cell.s ?? 0, f, cell.f!.array] : [r, c, enc(cell.v), cell.s ?? 0])
        }
      return {
        id: s.id,
        name: s.name,
        hidden: s.hidden,
        tabColor: s.tabColor,
        freeze: s.freeze,
        merges: s.merges,
        defaultColWidth: s.defaultColWidth,
        defaultRowHeight: s.defaultRowHeight,
        showGridlines: s.showGridlines,
        zoom: s.zoom,
        cols: [...s.cols],
        rows: [...s.rowInfo],
        cells
      }
    })
  }
}

export function restore(snap: Snapshot): Workbook {
  const wb = new Workbook()
  wb.date1904 = snap.date1904
  wb.calc = { ...wb.calc, ...snap.calc }
  for (const st of snap.styles.slice(1)) wb.styles.intern(st)
  wb.props = snap.props ?? {}
  for (const s of snap.sheets) {
    const sh = new Sheet(s.id, s.name)
    Object.assign(sh, { hidden: s.hidden, tabColor: s.tabColor, freeze: s.freeze, merges: s.merges, defaultColWidth: s.defaultColWidth, defaultRowHeight: s.defaultRowHeight, showGridlines: s.showGridlines, zoom: s.zoom ?? 1 })
    for (const [k, v] of s.cols) sh.cols.set(k, v as never)
    for (const [k, v] of s.rows) sh.rowInfo.set(k, v as never)
    wb.restoreSheet(sh, wb.sheets.length)
  }
  const sheetId = (n: string) => wb.sheetByName(n)?.id ?? null
  wb.tables = snap.tables ?? []
  wb.names = snap.names.map((n) => {
    let ast
    try {
      ast = parseFormula('=' + n.text, { row: 0, col: 0, sheetId })
    } catch {
      ast = undefined
    }
    return { ...n, ast }
  })
  snap.sheets.forEach((s, i) => {
    const sh = wb.sheets[i]
    for (const [r, c, v, st, f, array] of s.cells) {
      const cell: Cell = { v: dec(v) }
      if (st) cell.s = st
      if (f) {
        try {
          cell.f = { ast: parseFormula('=' + f, { row: r, col: c, sheetId }), array: array as never }
        } catch {
          cell.v = '=' + f
        }
      }
      sh.set(r, c, cell)
    }
  })
  wb.activeSheet = Math.min(snap.activeSheet ?? 0, wb.sheets.length - 1)
  return wb
}
