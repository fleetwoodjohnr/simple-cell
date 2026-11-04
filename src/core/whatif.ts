// What-if analysis: Data Tables (Excel's {=TABLE()}) and the Scenario Manager.
//
// A data table's result cells are values the engine fills in: for each one the input cell(s) are set to the values in
// the table's top row and left column, the workbook is recalculated, and the formula is read; then the inputs are put
// back. ponytail: one full recalculation per result cell, so a large table slows every edit; Excel has the same
// trade-off, which is why "Automatic except data tables" exists (calc mode 'autoNoTable').

import type { Doc, Tx } from './doc'
import type { Cell, DataTable, Scenario, ScenarioSet } from './model/workbook'
import { areaName, cellName, parseArea } from './refs'
import { CellError, numberToText, type Scalar } from './engine/values'

type Ref = { r: number; c: number }

const inTable = (t: DataTable, r: number, c: number) => r >= t.area.r0 && r <= t.area.r1 && c >= t.area.c0 && c <= t.area.c1

export function dataTableAt(doc: Doc, sheet: number, r: number, c: number): DataTable | undefined {
  return doc.wb.sheetById(sheet)?.dataTables.find((t) => inTable(t, r, c))
}

/** Set cells to values without undo (restored by the caller), and recalculate. */
function substitute(doc: Doc, sheet: number, cells: Ref[], values: Scalar[], saved: (Cell | undefined)[]) {
  const sh = doc.sheet(sheet)
  cells.forEach((p, i) => {
    sh.set(p.r, p.c, { ...(saved[i] ?? {}), f: undefined, v: values[i] })
    doc.engine.touch(sheet, p.r, p.c)
  })
  doc.engine.recalc()
}

function restore(doc: Doc, sheet: number, cells: Ref[], saved: (Cell | undefined)[]) {
  const sh = doc.sheet(sheet)
  cells.forEach((p, i) => {
    sh.set(p.r, p.c, saved[i])
    doc.engine.touch(sheet, p.r, p.c)
  })
  doc.engine.recalc()
}

/** A data table's results, row by row. */
export function computeDataTable(doc: Doc, sheet: number, t: DataTable): Scalar[][] {
  const sh = doc.sheet(sheet)
  const { r0, c0, r1, c1 } = t.area
  const inputs = [t.rowInput, t.colInput].filter((p): p is Ref => !!p)
  const saved = inputs.map((p) => sh.get(p.r, p.c))
  const val = (r: number, c: number): Scalar => sh.get(r, c)?.v ?? null
  // The values substituted are read before anything changes (they may be formulas too).
  const top = Array.from({ length: c1 - c0 + 1 }, (_, j) => val(r0 - 1, c0 + j))
  const left = Array.from({ length: r1 - r0 + 1 }, (_, i) => val(r0 + i, c0 - 1))
  const out: Scalar[][] = []
  try {
    for (let i = 0; i <= r1 - r0; i++) {
      const row: Scalar[] = []
      for (let j = 0; j <= c1 - c0; j++) {
        let at: Ref
        if (t.rowInput && t.colInput) {
          substitute(doc, sheet, [t.rowInput, t.colInput], [top[j], left[i]], saved)
          at = { r: r0 - 1, c: c0 - 1 }
        } else if (t.colInput) {
          substitute(doc, sheet, [t.colInput], [left[i]], saved)
          at = { r: r0 - 1, c: c0 + j }
        } else {
          substitute(doc, sheet, [t.rowInput!], [top[j]], saved)
          at = { r: r0 + i, c: c0 - 1 }
        }
        row.push(sh.get(at.r, at.c)?.v ?? 0)
      }
      out.push(row)
    }
  } finally {
    restore(doc, sheet, inputs, saved)
  }
  return out
}

/**
 * Recalculate every data table and write the results (called after a recalculation that changed something).
 * Returns whether any result changed.
 */
export function recalcDataTables(doc: Doc): boolean {
  let changed = false
  for (const sh of doc.wb.sheets)
    for (const t of sh.dataTables) {
      const values = computeDataTable(doc, sh.id, t)
      values.forEach((row, i) =>
        row.forEach((v, j) => {
          const r = t.area.r0 + i
          const c = t.area.c0 + j
          const cell = sh.get(r, c)
          if (cell && sameScalar(cell.v, v) && !cell.f) return
          sh.set(r, c, { ...(cell ?? {}), f: undefined, v })
          doc.engine.touch(sh.id, r, c)
          changed = true
        })
      )
    }
  if (changed) doc.engine.recalc()
  return changed
}

const sameScalar = (a: Scalar, b: Scalar) => a === b || (a instanceof CellError && b instanceof CellError && a.code === b.code)

/** Data ▸ What-If Analysis ▸ Data Table: `area` is the whole selection (formulas, input values and results). */
export function createDataTable(doc: Doc, sheet: number, area: { r0: number; c0: number; r1: number; c1: number }, rowInput: Ref | null, colInput: Ref | null) {
  if (!rowInput && !colInput) throw new Error('Enter a row input cell, a column input cell, or both.')
  if (area.r1 - area.r0 < 1 || area.c1 - area.c0 < 1) throw new Error('Select the formulas, the input values and the cells for the results (at least two rows and two columns).')
  const t: DataTable = { area: { r0: area.r0 + 1, c0: area.c0 + 1, r1: area.r1, c1: area.c1 }, rowInput, colInput }
  for (const p of [rowInput, colInput]) if (p && p.r >= area.r0 && p.r <= area.r1 && p.c >= area.c0 && p.c <= area.c1) throw new Error('The input cell must be outside the data table.')
  const sh = doc.sheet(sheet)
  if (sh.dataTables.some((o) => o.area.r0 <= t.area.r1 && o.area.r1 >= t.area.r0 && o.area.c0 <= t.area.c1 && o.area.c1 >= t.area.c0)) throw new Error('The selection overlaps another data table.')
  doc.tablesDirty = true
  return doc.transact('Data table', (tx) => {
    // The result cells start empty; the recalculation after the transaction fills them in.
    for (let r = t.area.r0; r <= t.area.r1; r++)
      for (let c = t.area.c0; c <= t.area.c1; c++) {
        const cell = sh.get(r, c)
        if (cell) tx.setCell(sheet, r, c, cell.s ? { v: null, s: cell.s } : undefined)
      }
    tx.prop(sh, 'dataTables', [...sh.dataTables, t])
  })
}

/** Tables entirely inside the areas (cleared together with them: a whole table can be deleted, not part of one). */
export function removeTablesInTx(doc: Doc, tx: Tx, sheet: number, areas: { r0: number; c0: number; r1: number; c1: number }[]) {
  const sh = doc.sheet(sheet)
  const covered = (t: DataTable) => areas.some((a) => a.r0 <= t.area.r0 && a.r1 >= t.area.r1 && a.c0 <= t.area.c0 && a.c1 >= t.area.c1)
  if (sh.dataTables.some(covered)) tx.prop(sh, 'dataTables', sh.dataTables.filter((t) => !covered(t)))
}

// ---------- Scenario Manager ----------

const MAX_CHANGING = 32

/** "B1:B3,D1" → cells, at most 32 (Excel's limit). */
export function scenarioCells(text: string): Ref[] {
  const out: Ref[] = []
  for (const raw of text.split(',')) {
    const t = raw.trim().replace(/^.*!/, '')
    if (!t) continue
    const a = parseArea(t)
    if (!a) throw new Error(`“${raw.trim()}” isn’t a cell or range.`)
    for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) out.push({ r, c })
    if (out.length > MAX_CHANGING) throw new Error(`A scenario can change at most ${MAX_CHANGING} cells.`)
  }
  if (!out.length) throw new Error('Enter the changing cells.')
  return out
}

/** "$B$1,$B$2" for a list of cells. */
export const cellsText = (cells: Ref[]) => cells.map((p) => absolute(cellName(p.r, p.c))).join(',')
const absolute = (a1: string) => a1.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')

/** Replace the sheet's scenarios (the dialog's Add, Edit and Delete). */
export function setScenarios(doc: Doc, sheet: number, list: Scenario[], current = 0) {
  const sh = doc.sheet(sheet)
  const names = new Set<string>()
  for (const s of list) {
    if (!s.name.trim()) throw new Error('Every scenario needs a name.')
    if (names.has(s.name.toLowerCase())) throw new Error(`There is already a scenario named “${s.name}”.`)
    names.add(s.name.toLowerCase())
    if (s.cells.length !== s.values.length) throw new Error(`“${s.name}” needs one value per changing cell.`)
  }
  const next: ScenarioSet | null = list.length ? { list, current: Math.min(current, list.length - 1), show: Math.min(sh.scenarios?.show ?? 0, list.length - 1), resultCells: sh.scenarios?.resultCells } : null
  return doc.transact('Scenarios', (tx) => tx.prop(sh, 'scenarios', next))
}

/** Show: put the scenario's values into its changing cells (one undoable step). */
export function showScenario(doc: Doc, sheet: number, index: number) {
  const sh = doc.sheet(sheet)
  const s = sh.scenarios?.list[index]
  if (!s) throw new Error('There is no such scenario.')
  return doc.transact(`Show scenario “${s.name}”`, (tx) => {
    s.cells.forEach((p, i) => doc.inputTx(tx, sheet, p.r, p.c, s.values[i]))
    tx.prop(sh, 'scenarios', { ...sh.scenarios!, current: index, show: index })
  })
}

/** Scenario Summary report on a new sheet: each scenario's changing values and the result cells' values under it. */
export function scenarioSummary(doc: Doc, sheet: number, resultText: string): number {
  const sh = doc.sheet(sheet)
  const set = sh.scenarios
  if (!set?.list.length) throw new Error('This sheet has no scenarios.')
  const results = resultText.trim() ? scenarioCells(resultText) : []
  // Every changing cell of any scenario, in first-seen order.
  const changing: Ref[] = []
  for (const s of set.list) for (const p of s.cells) if (!changing.some((q) => q.r === p.r && q.c === p.c)) changing.push(p)
  const read = (p: Ref): Scalar => sh.get(p.r, p.c)?.v ?? null
  const current = { changing: changing.map(read), results: results.map(read) }
  const saved = changing.map((p) => sh.get(p.r, p.c))
  const columns: { changing: Scalar[]; results: Scalar[] }[] = []
  try {
    for (const s of set.list) {
      const values = changing.map((p, i) => {
        const k = s.cells.findIndex((q) => q.r === p.r && q.c === p.c)
        return k < 0 ? current.changing[i] : typedValue(s.values[k])
      })
      substitute(doc, sheet, changing, values, saved)
      columns.push({ changing: values, results: results.map(read) })
    }
  } finally {
    restore(doc, sheet, changing, saved)
  }
  let id = 0
  doc.transact('Scenario summary', (tx) => {
    let name = 'Scenario Summary'
    for (let k = 2; doc.wb.sheetByName(name); k++) name = `Scenario Summary ${k}`
    const out = doc.wb.createSheet(name)
    id = out.id
    const at = doc.wb.sheetIndex(sheet) + 1
    tx.custom(
      () => doc.wb.sheets.splice(at, 0, out),
      () => doc.wb.sheets.splice(doc.wb.sheetIndex(out.id), 1)
    )
    tx.structural = true
    const put = (r: number, c: number, v: Scalar) => doc.inputTx(tx, out.id, r, c, v === null ? '' : typeof v === 'number' ? numberToText(v) : v instanceof CellError ? v.code : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : "'" + v)
    put(0, 1, 'Scenario Summary')
    put(1, 3, 'Current Values:')
    set.list.forEach((s, j) => put(1, 4 + j, s.name))
    put(2, 1, 'Changing Cells:')
    changing.forEach((p, i) => {
      put(3 + i, 2, absolute(cellName(p.r, p.c)))
      put(3 + i, 3, current.changing[i])
      columns.forEach((col, j) => put(3 + i, 4 + j, col.changing[i]))
    })
    let r = 3 + changing.length
    if (results.length) {
      put(r++, 1, 'Result Cells:')
      results.forEach((p, i) => {
        put(r, 2, absolute(cellName(p.r, p.c)))
        put(r, 3, current.results[i])
        columns.forEach((col, j) => put(r, 4 + j, col.results[i]))
        r++
      })
    }
    put(r + 1, 1, 'Notes: The Current Values column shows the changing cells as they were when this report was made.')
    tx.prop(sh, 'scenarios', { ...set, resultCells: results.length ? cellsText(results) : set.resultCells })
  })
  return id
}

/** A scenario value as Excel stores it ("10", "abc") → the cell value it stands for. */
function typedValue(text: string): Scalar {
  const n = Number(text)
  return text.trim() !== '' && Number.isFinite(n) ? n : text
}

/** For the dialog: describe a scenario's cells. */
export const describeCells = (s: Scenario) => (s.cells.length ? areaName({ r0: s.cells[0].r, c0: s.cells[0].c, r1: s.cells[0].r, c1: s.cells[0].c }) + (s.cells.length > 1 ? ` and ${s.cells.length - 1} more` : '') : '')
