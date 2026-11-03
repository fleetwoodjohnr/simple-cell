// The document host: every operation the UI can ask of an open workbook. It runs inside the core Web Worker
// (worker.ts) and directly in unit tests. All arguments and results are structured-clone friendly.

import { Doc, type CfSpec, type DvSpec, type GoToKind, type PasteSpecialOptions, type ViewCell } from './doc'
import type { DvAlert } from './validation'
import type { FilterColumn } from './filter'
import type { Link } from './model/workbook'
import type { PageSetup } from './print'
import type { Chart, ChartType } from './charts'
import { MAX_COLS, MAX_ROWS, isError, type Area, type Scalar } from './engine/values'
import { FormulaSyntaxError } from './engine/parser'
import type { Style } from './model/styles'
import { mean, sum, variance } from './engine/functions/helpers'
import { navigate, type Dir } from './nav'
import { cellKey, cellName, fromKey } from './refs'
import { snapshot, restore, type Snapshot } from './snapshot'
import { importWorkbook, exportWorkbook, type FormatId } from './formats/registry'
import { shiftXlsxCarry } from './formats/xlsx/shift'
import type { FormatWarning } from './formats/types'
import { CATALOG } from './engine/catalog'
import { getFunction } from './engine/evaluate'
import { findCells, replaceCells, type FindOptions } from './find'
import { sortRange, type SortKey } from './sort'
import { autofill } from './autofill'

export interface SheetMeta {
  id: number
  name: string
  hidden: boolean | string
  tabColor: string | null
  freeze: { rows: number; cols: number }
  defaultColWidth: number
  defaultRowHeight: number
  /** Sparse widths/heights; 0 = hidden. */
  cols: [number, number][]
  rows: [number, number][]
  merges: Area[]
  used: { r: number; c: number }
  showGridlines: boolean
  zoom: number
  spills: Area[]
  /** Rows hidden by an AutoFilter (to draw the filter state). */
  filtered: number[]
  /** The sheet has conditional formatting: any change may restyle other cells, so the client refetches. */
  cf: boolean
  /** AutoFilter ranges (the sheet's and tables'), for their header buttons. */
  filters: { key: string; area: Area; active: number[] }[]
  /** Outline (group) levels by row / column index. */
  outlineRows: [number, number][]
  outlineCols: [number, number][]
  charts: Omit<Chart, 'carry'>[]
}

export interface WorkbookMeta {
  sheets: SheetMeta[]
  styles: Style[]
  names: { name: string; scope: number | null; text: string; comment?: string }[]
  tables: ReturnType<Doc['tablesOn']>[number] extends infer T ? (T & { sheet: number })[] : never
  calcMode: 'auto' | 'manual'
  iterate: boolean
  canUndo: string | null
  canRedo: string | null
  circular: { sheet: number; r: number; c: number }[]
  version: number
  structureVersion: number
  /** Bumped when the display changed beyond the edited cells (conditional formats). */
  viewVersion: number
  date1904: boolean
}

export interface Stats {
  count: number
  numbers: number
  sum: number
  average: number | null
  min: number | null
  max: number | null
}

export class Host {
  doc: Doc
  warnings: FormatWarning[] = []
  format: FormatId | null = null

  constructor(doc?: Doc) {
    this.doc = doc ?? new Doc()
  }

  // ---------- files ----------

  static async open(bytes: Uint8Array, format: FormatId): Promise<Host> {
    const r = await importWorkbook(format, bytes)
    const h = new Host(new Doc(r.workbook))
    if (['xlsx', 'xlsm', 'xltx', 'xltm'].includes(format)) h.doc.carryShifters.xlsx = (tx, s) => shiftXlsxCarry(h.doc.wb, tx, s)
    h.warnings = [...r.warnings, ...h.compareCached()]
    h.format = format
    return h
  }

  /**
   * After loading, compare each formula's result with the one the file was saved with. Differences point at
   * calculation differences between Simple Cell and the program that saved the file; they are reported, not hidden.
   */
  compareCached(): FormatWarning[] {
    const cached = (this.doc.wb.carry as { cached?: { sheet: number; r: number; c: number; v: Scalar }[] } | null)?.cached
    if (!cached?.length) return []
    const diffs: string[] = []
    for (const { sheet, r, c, v } of cached) {
      const cell = this.doc.wb.sheetById(sheet)?.get(r, c)
      if (!cell?.f || cell.f.unsupported || this.doc.engine.volatile.has(cellKey(sheet, r, c))) continue
      const now = cell.v
      const same = now === v || (typeof now === 'number' && typeof v === 'number' && Math.abs(now - v) <= Math.max(1e-9, Math.abs(v) * 1e-9)) || (isError(now) && isError(v) && now.code === v.code) || (v === null && (now === 0 || now === ''))
      if (!same) diffs.push(`${this.doc.wb.sheetById(sheet)?.name}!${cellName(r, c)}`)
    }
    if (!diffs.length) return []
    return [{ level: 'loss', message: `${diffs.length} formula result(s) differ from the values saved in the file (${diffs.slice(0, 8).join(', ')}${diffs.length > 8 ? ', …' : ''}). This can mean a calculation difference between Simple Cell and the program that saved the file.` }]
  }

  async save(format: FormatId, opts: Record<string, unknown> = {}): Promise<{ data: Uint8Array; warnings: FormatWarning[] }> {
    return exportWorkbook(format, this.doc.wb, opts)
  }

  snapshot(): Snapshot {
    return snapshot(this.doc.wb)
  }

  static fromSnapshot(s: Snapshot): Host {
    return new Host(new Doc(restore(s)))
  }

  // ---------- reading ----------

  meta(): WorkbookMeta {
    const wb = this.doc.wb
    return {
      sheets: wb.sheets.map((s) => this.sheetMeta(s.id)),
      styles: wb.styles.list,
      names: wb.names.filter((n) => !n.hidden).map((n) => ({ name: n.name, scope: n.scope, text: n.text, comment: n.comment })),
      tables: wb.sheets.flatMap((s) => this.doc.tablesOn(s.id).map((t) => ({ ...t, sheet: s.id }))),
      calcMode: wb.calc.mode,
      iterate: wb.calc.iterate,
      canUndo: this.doc.undoStack.at(-1)?.label ?? null,
      canRedo: this.doc.redoStack.at(-1)?.label ?? null,
      circular: [...this.doc.engine.circular].map(fromKey),
      version: this.doc.version,
      structureVersion: this.doc.structureVersion,
      viewVersion: this.doc.viewVersion,
      date1904: wb.date1904
    }
  }

  sheetMeta(id: number): SheetMeta {
    const s = this.doc.sheet(id)
    const filtered: number[] = []
    const rows: [number, number][] = []
    const outlineRows: [number, number][] = []
    const outlineCols: [number, number][] = []
    for (const [r, ri] of s.rowInfo) {
      if (ri.outline) outlineRows.push([r, ri.outline])
      if (ri.hidden) rows.push([r, 0])
      else if (ri.height !== undefined) rows.push([r, ri.height])
      if (ri.filtered) filtered.push(r)
    }
    const cols: [number, number][] = []
    for (const [c, ci] of s.cols) {
      if (ci.outline) outlineCols.push([c, ci.outline])
      if (ci.hidden) cols.push([c, 0])
      else if (ci.width !== undefined) cols.push([c, ci.width])
    }
    const spills = [...this.doc.engine.spills].filter(([k]) => fromKey(k).sheet === id).map(([, a]) => a)
    return {
      id: s.id,
      name: s.name,
      hidden: s.hidden,
      tabColor: s.tabColor,
      freeze: s.freeze,
      defaultColWidth: s.defaultColWidth,
      defaultRowHeight: s.defaultRowHeight,
      cols,
      rows,
      merges: s.merges,
      used: s.used(),
      showGridlines: s.showGridlines,
      zoom: s.zoom,
      spills,
      filtered,
      cf: s.condFormats.length > 0,
      filters: this.doc.filtersOn(s.id),
      outlineRows,
      outlineCols,
      charts: s.charts.map(({ carry: _c, ...c }) => c)
    }
  }

  /** Cells to draw in a window of the sheet. `chars` maps column → characters that fit (for General). */
  viewport(sheet: number, r0: number, r1: number, c0: number, c1: number): ViewCell[] {
    const s = this.doc.wb.sheetById(sheet)
    if (!s) return []
    const out: ViewCell[] = []
    const cf = this.doc.condFormatsIn(sheet, { sheet, r0, r1, c0, c1 })
    const chars = (c: number) => Math.max(1, Math.floor(s.colWidth(c) / 7))
    const tables = this.doc.wb.tables.filter((t) => t.sheet === sheet && t.area.r0 <= r1 && t.area.r1 >= r0 && t.area.c0 <= c1 && t.area.c1 >= c0)
    const ts = (r: number, c: number) => (tables.length ? this.doc.tableStyleAt(sheet, r, c) : null)
    const seen = new Set<number>()
    s.forEachIn({ sheet, r0, r1, c0, c1 }, (r, c, cell) => {
      const k = r * 16384 + c
      seen.add(k)
      out.push(this.doc.viewOf(cell, r, c, chars(c), cf.get(k), ts(r, c)))
    })
    // Notes and links on empty cells still show (red corner, link style).
    for (const m of [s.notes, s.links] as Map<number, unknown>[])
      for (const k of m.keys()) {
        const r = Math.floor(k / 16384)
        const c = k % 16384
        if (r < r0 || r > r1 || c < c0 || c > c1 || seen.has(k) || cf.has(k)) continue
        seen.add(k)
        out.push(this.doc.viewOf({ v: null }, r, c, 1, undefined, ts(r, c)))
      }
    // Empty cells that conditional formatting or a table style colours.
    for (const [k, res] of cf) if (!seen.has(k)) {
      seen.add(k)
      out.push(this.doc.viewOf({ v: null }, Math.floor(k / 16384), k % 16384, 1, res, ts(Math.floor(k / 16384), k % 16384)))
    }
    for (const t of tables)
      for (let r = Math.max(r0, t.area.r0); r <= Math.min(r1, t.area.r1); r++)
        for (let c = Math.max(c0, t.area.c0); c <= Math.min(c1, t.area.c1); c++) {
          const k = r * 16384 + c
          if (seen.has(k)) continue
          seen.add(k)
          out.push(this.doc.viewOf({ v: null }, r, c, 1, undefined, ts(r, c)))
        }
    if (s.notes.size || s.links.size)
      for (const v of out) {
        const k = v.r * 16384 + v.c
        if (s.notes.has(k)) v.note = 1
        if (s.links.has(k)) v.link = 1
      }
    return out
  }

  /** Cells whose display changed since the last call (any sheet). */
  changes(): { sheet: number; cells: ViewCell[]; cleared: [number, number][] }[] {
    const by = new Map<number, { cells: ViewCell[]; cleared: [number, number][] }>()
    for (const k of this.doc.takeChanges()) {
      const { sheet, r, c } = fromKey(k)
      const s = this.doc.wb.sheetById(sheet)
      if (!s) continue
      let e = by.get(sheet)
      if (!e) by.set(sheet, (e = { cells: [], cleared: [] }))
      const cell = s.get(r, c)
      const cf = s.condFormats.length ? this.doc.condFormatsIn(sheet, { sheet, r0: r, r1: r, c0: c, c1: c }).get(r * 16384 + c) : undefined
      const ts = this.doc.tableStyleAt(sheet, r, c)
      const marks = s.notes.has(r * 16384 + c) || s.links.has(r * 16384 + c)
      if (cell || cf || ts || marks) {
        const v = this.doc.viewOf(cell ?? { v: null }, r, c, Math.max(1, Math.floor(s.colWidth(c) / 7)), cf, ts)
        if (s.notes.has(r * 16384 + c)) v.note = 1
        if (s.links.has(r * 16384 + c)) v.link = 1
        e.cells.push(v)
      }
      else e.cleared.push([r, c])
    }
    return [...by].map(([sheet, e]) => ({ sheet, ...e }))
  }

  editText(sheet: number, r: number, c: number) {
    return this.doc.editText(sheet, r, c)
  }

  styleAt(sheet: number, r: number, c: number): Style {
    return this.doc.styleAt(sheet, r, c)
  }

  cellInfo(sheet: number, r: number, c: number) {
    const cell = this.doc.wb.sheetById(sheet)?.get(r, c)
    const v = cell?.v ?? null
    return {
      editText: this.doc.editText(sheet, r, c),
      formula: !!cell?.f,
      unsupported: cell?.f?.unsupported ?? null,
      stale: !!cell?.stale,
      error: isError(v) ? v.code : null,
      type: v === null ? 'empty' : typeof v === 'number' ? 'number' : typeof v === 'string' ? 'text' : typeof v === 'boolean' ? 'logical' : 'error',
      spill: cell?.spillOf !== undefined ? fromKey(cell.spillOf) : null,
      style: this.doc.styleAt(sheet, r, c),
      table: this.doc.wb.tableAt(sheet, r, c)?.name ?? null
    }
  }

  stats(areas: Area[]): Stats {
    const nums: number[] = []
    let count = 0
    const seen = new Set<string>()
    for (const a of areas) {
      const s = this.doc.wb.sheetById(a.sheet)
      s?.forEachIn(a, (r, c, cell) => {
        const k = `${r},${c}`
        if (seen.has(k)) return
        seen.add(k)
        if (cell.v === null) return
        count++
        if (typeof cell.v === 'number') nums.push(cell.v)
      })
    }
    const v = variance(nums, true)
    void v
    return {
      count,
      numbers: nums.length,
      sum: sum(nums),
      average: nums.length ? mean(nums) : null,
      min: nums.length ? Math.min(...nums) : null,
      max: nums.length ? Math.max(...nums) : null
    }
  }

  navigate(sheet: number, r: number, c: number, dir: Dir): { r: number; c: number } {
    return navigate(this.doc.sheet(sheet), r, c, dir)
  }

  /** Values for copying: TSV and an HTML table (with basic formatting). */
  copyText(a: Area): { tsv: string; html: string } {
    const s = this.doc.sheet(a.sheet)
    const rows: string[] = []
    const html: string[] = []
    for (let r = a.r0; r <= a.r1; r++) {
      if (s.rowInfo.get(r)?.filtered) continue
      const line: string[] = []
      const tds: string[] = []
      for (let c = a.c0; c <= a.c1; c++) {
        const vc = this.doc.display(a.sheet, r, c, 255)
        const t = vc?.t ?? ''
        line.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t)
        const st = this.doc.styleAt(a.sheet, r, c)
        const css = [st.bold && 'font-weight:bold', st.italic && 'font-style:italic', st.color && `color:${st.color}`, st.fill && `background:${st.fill}`, st.hAlign && st.hAlign !== 'general' && `text-align:${st.hAlign}`].filter(Boolean).join(';')
        tds.push(`<td${css ? ` style="${css}"` : ''}>${t.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</td>`)
      }
      rows.push(line.join('\t'))
      html.push(`<tr>${tds.join('')}</tr>`)
    }
    return { tsv: rows.join('\n'), html: `<table>${html.join('')}</table>` }
  }

  functionList() {
    return CATALOG.map((e) => ({ name: e.n, category: e.cat, description: e.d, implemented: !!getFunction(e.n) }))
  }

  // ---------- editing (each returns an error message or null) ----------

  private run(fn: () => unknown): { ok: true } | { ok: false; error: string; pos?: number } {
    try {
      fn()
      return { ok: true }
    } catch (e) {
      return { ok: false, error: (e as Error).message, pos: e instanceof FormulaSyntaxError ? e.pos : undefined }
    }
  }

  /**
   * Typing into a cell. A value that breaks the cell's data validation comes back with the alert to show: for a
   * "stop" rule the entry is undone here; for "warning" and "information" it stays until the user says otherwise.
   */
  input(sheet: number, r: number, c: number, text: string, array?: Area): ReturnType<Host['run']> & { alert?: DvAlert } {
    const before = this.doc.undoStack.length
    const res = this.run(() => this.doc.input(sheet, r, c, text, { array }))
    if (!res.ok || array) return res
    const alert = this.doc.checkValidation(sheet, r, c)
    if (!alert) return res
    if (alert.style === 'stop') {
      if (this.doc.undoStack.length > before) {
        this.doc.undo()
        this.doc.redoStack.pop()
      }
      return { ok: false, error: alert.message, alert }
    }
    return { ...res, alert }
  }

  /** Enter the same text into every cell of the areas (Ctrl+Enter); formulas adjust per cell. */
  inputAll(areas: Area[], text: string) {
    return this.run(() =>
      this.doc.transact('Typing', (tx) => {
        for (const a of areas) for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) this.doc.inputTx(tx, a.sheet, r, c, text)
      })
    )
  }

  /** Paste rows of text (TSV from another app): numbers, dates and formulas are recognised. */
  pasteText(sheet: number, r: number, c: number, rows: string[][]) {
    return this.run(() =>
      this.doc.transact('Paste', (tx) => {
        rows.forEach((line, i) => line.forEach((t, j) => this.doc.inputTx(tx, sheet, r + i, c + j, t)))
      })
    )
  }

  /** Paste cells copied inside Simple Cell (formulas move relative to the new position). */
  pasteCells(src: Area, dst: Area, what: 'all' | 'values' | 'formats' | 'formulas' = 'all', cut = false) {
    return this.run(() => {
      const block = this.doc.block(src)
      this.doc.transact(cut ? 'Move' : 'Paste', (tx) => {
        if (cut) {
          const s = this.doc.sheet(src.sheet)
          s.forEachIn(src, (r, c) => tx.setCell(src.sheet, r, c, undefined))
        }
        this.doc.pasteTx(tx, dst, block, what)
      })
    })
  }

  goToSpecial(sheet: number, within: Area[], kind: GoToKind, types?: string): { ok: true; areas: Area[] } | { ok: false; error: string } {
    try {
      return { ok: true, areas: this.doc.goToSpecial(sheet, within, kind, types) }
    } catch (e) {
      return { ok: false, error: (e as Error).message }
    }
  }
  flashFill(sheet: number, r: number, c: number): { ok: true; filled: number } | { ok: false; error: string } {
    try {
      return { ok: true, filled: this.doc.flashFill(sheet, r, c) }
    } catch (e) {
      return { ok: false, error: (e as Error).message }
    }
  }
  pasteSpecial(src: Area, dst: Area, o: PasteSpecialOptions) {
    return this.run(() => this.doc.pasteSpecial(src, dst, o))
  }
  clear(areas: Area[], what: 'contents' | 'formats' | 'all') {
    return this.run(() => this.doc.clear(areas, what))
  }
  formatCells(areas: Area[], patch: Partial<Style>) {
    return this.run(() => this.doc.format(areas, patch))
  }
  /** Toggle a boolean style property based on the active cell (Bold, Italic…). */
  toggle(areas: Area[], key: 'bold' | 'italic' | 'strike' | 'wrap', active: { sheet: number; r: number; c: number }) {
    const on = !this.doc.styleAt(active.sheet, active.r, active.c)[key]
    return this.run(() => this.doc.format(areas, { [key]: on || undefined }))
  }
  insert(sheet: number, axis: 'row' | 'col', at: number, count: number) {
    return this.run(() => (axis === 'row' ? this.doc.insertRows(sheet, at, count) : this.doc.insertCols(sheet, at, count)))
  }
  remove(sheet: number, axis: 'row' | 'col', at: number, count: number) {
    return this.run(() => (axis === 'row' ? this.doc.deleteRows(sheet, at, count) : this.doc.deleteCols(sheet, at, count)))
  }
  undo() {
    return this.doc.undo()?.selection ?? null
  }
  redo() {
    return this.doc.redo()?.selection ?? null
  }
  addSheet(name?: string, index?: number) {
    let id = -1
    const r = this.run(() => (id = this.doc.addSheet(name, index)))
    return { ...r, id }
  }
  deleteSheet(id: number) {
    return this.run(() => this.doc.deleteSheet(id))
  }
  renameSheet(id: number, name: string) {
    return this.run(() => this.doc.renameSheet(id, name))
  }
  moveSheet(id: number, to: number) {
    return this.run(() => this.doc.moveSheet(id, to))
  }
  sheetProps(id: number, props: { tabColor?: string | null; hidden?: boolean; showGridlines?: boolean; zoom?: number }) {
    return this.run(() =>
      this.doc.transact('Sheet', (tx) => {
        const s = this.doc.sheet(id)
        if (props.tabColor !== undefined) tx.prop(s, 'tabColor', props.tabColor)
        if (props.hidden !== undefined) {
          if (props.hidden && this.doc.wb.sheets.filter((x) => !x.hidden && x.id !== id).length === 0) throw new Error('A workbook must contain at least one visible sheet.')
          tx.prop(s, 'hidden', props.hidden ? 'hidden' : false)
        }
        if (props.showGridlines !== undefined) tx.prop(s, 'showGridlines', props.showGridlines)
        if (props.zoom !== undefined) tx.prop(s, 'zoom', props.zoom)
      })
    )
  }
  setColWidths(sheet: number, cols: number[], width: number | undefined) {
    return this.run(() => this.doc.setColWidths(sheet, cols, width))
  }
  setRowHeights(sheet: number, rows: number[], height: number | undefined) {
    return this.run(() => this.doc.setRowHeights(sheet, rows, height))
  }
  /** Best width for columns from their displayed text (double-click on a column border). */
  autofitCols(sheet: number, cols: number[], measure: (text: string, style: number) => number) {
    void measure
    return this.run(() => {
      const s = this.doc.sheet(sheet)
      const widths: number[] = cols.map((c) => {
        let w = 0
        s.forEachIn({ sheet, r0: 0, r1: MAX_ROWS - 1, c0: c, c1: c }, (r, cc, cell) => {
          const t = this.doc.viewOf(cell, r, cc, 255).t
          w = Math.max(w, t.length * 7 + 10)
        })
        return w ? Math.min(Math.max(w, 20), 800) : s.defaultColWidth
      })
      this.doc.transact('AutoFit column width', (tx) => {
        cols.forEach((c, i) => tx.mapSet(s.cols, c, { ...(s.cols.get(c) ?? {}), width: widths[i], hidden: undefined }))
      })
    })
  }
  /** Column texts for autofit measured by the UI (which knows fonts). */
  columnTexts(sheet: number, col: number, limit = 5000): { t: string; s: number }[] {
    const s = this.doc.sheet(sheet)
    const out: { t: string; s: number }[] = []
    s.forEachIn({ sheet, r0: 0, r1: MAX_ROWS - 1, c0: col, c1: col }, (r, c, cell) => {
      if (out.length >= limit) return
      const vc = this.doc.viewOf(cell, r, c, 255)
      if (vc.t) out.push({ t: vc.t, s: vc.s })
    })
    return out
  }
  hide(sheet: number, axis: 'row' | 'col', indexes: number[], hidden: boolean) {
    return this.run(() => this.doc.hide(sheet, axis, indexes, hidden))
  }
  freeze(sheet: number, rows: number, cols: number) {
    return this.run(() => this.doc.freeze(sheet, rows, cols))
  }
  merge(area: Area, mode: 'all' | 'across' | 'center') {
    return this.run(() => this.doc.merge(area, mode))
  }
  unmerge(area: Area) {
    return this.run(() => this.doc.unmerge(area))
  }
  defineName(name: string, text: string, scope: number | null, comment?: string) {
    return this.run(() => this.doc.defineName(name, text, scope, comment))
  }
  deleteName(name: string, scope: number | null) {
    return this.run(() => this.doc.deleteName(name, scope))
  }
  createTable(sheet: number, area: Area, hasHeaders: boolean, style?: string) {
    return this.run(() => this.doc.createTable(sheet, area, hasHeaders, style))
  }
  setTableOptions(name: string, patch: Parameters<Doc['setTableOptions']>[1]) {
    return this.run(() => this.doc.setTableOptions(name, patch))
  }
  convertTableToRange(name: string) {
    return this.run(() => this.doc.convertTableToRange(name))
  }
  toggleAutoFilter(sheet: number, area: Area | null) {
    return this.run(() => this.doc.toggleAutoFilter(sheet, area))
  }
  setFilter(sheet: number, key: string, col: number, fc: Omit<FilterColumn, 'col'> | null) {
    return this.run(() => this.doc.setFilter(sheet, key, col, fc))
  }
  clearFilters(sheet: number, key: string) {
    return this.run(() => this.doc.clearFilters(sheet, key))
  }
  reapplyFilter(sheet: number, key: string) {
    return this.run(() => this.doc.reapplyFilter(sheet, key))
  }
  filterValues(sheet: number, key: string, col: number) {
    return this.doc.filterValues(sheet, key, col)
  }
  addChartFromRange(sheet: number, area: Area, type?: ChartType): { ok: true; id: number } | { ok: false; error: string } {
    try {
      return { ok: true, id: this.doc.addChartFromRange(sheet, area, type) }
    } catch (e) {
      return { ok: false, error: (e as Error).message }
    }
  }
  updateChart(sheet: number, id: number, patch: Partial<Omit<Chart, 'id'>>) {
    return this.run(() => this.doc.updateChart(sheet, id, patch))
  }
  deleteChart(sheet: number, id: number) {
    return this.run(() => this.doc.deleteChart(sheet, id))
  }
  chartData(sheet: number, id: number) {
    return this.doc.chartData(sheet, id)
  }
  printJob(sheet: number, area?: Area) {
    return this.doc.printJob(sheet, area)
  }
  setPageSetup(sheet: number, patch: Partial<PageSetup>) {
    return this.run(() => this.doc.setPageSetup(sheet, patch))
  }
  setPrintName(sheet: number, which: 'area' | 'titles', text: string | null) {
    return this.run(() => this.doc.setPrintName(sheet, which, text))
  }
  pageSetup(sheet: number) {
    return { ...this.doc.sheet(sheet).pageSetup, ...this.doc.printNames(sheet) }
  }
  group(sheet: number, axis: 'row' | 'col', from: number, to: number, delta: 1 | -1) {
    return this.run(() => this.doc.group(sheet, axis, from, to, delta))
  }
  showDetail(sheet: number, axis: 'row' | 'col', from: number, to: number, show: boolean) {
    return this.run(() => this.doc.showDetail(sheet, axis, from, to, show))
  }
  showOutlineLevel(sheet: number, axis: 'row' | 'col', level: number) {
    return this.run(() => this.doc.showOutlineLevel(sheet, axis, level))
  }
  insertCells(area: Area, dir: 'down' | 'right') {
    return this.run(() => this.doc.insertCells(area, dir))
  }
  deleteCells(area: Area, dir: 'up' | 'left') {
    return this.run(() => this.doc.deleteCells(area, dir))
  }
  setNote(sheet: number, r: number, c: number, text: string | null, author?: string) {
    return this.run(() => this.doc.setNote(sheet, r, c, text, author))
  }
  setLink(areas: Area[], link: Link | null, display?: string) {
    return this.run(() => this.doc.setLink(areas, link, display))
  }
  noteAt(sheet: number, r: number, c: number) {
    return this.doc.noteAt(sheet, r, c)
  }
  linkAt(sheet: number, r: number, c: number) {
    return this.doc.linkAt(sheet, r, c)
  }
  setValidation(sheet: number, areas: Area[], spec: DvSpec) {
    return this.run(() => this.doc.setValidation(sheet, areas, spec))
  }
  validationInfo(sheet: number, r: number, c: number) {
    return this.doc.validationInfo(sheet, r, c)
  }
  invalidCells(sheet: number) {
    return this.doc.invalidCells(sheet)
  }
  addCondFormat(sheet: number, ranges: Area[], spec: CfSpec) {
    return this.run(() => this.doc.addCondFormat(sheet, ranges, spec))
  }
  clearCondFormats(sheet: number, areas: Area[] | null) {
    return this.run(() => this.doc.clearCondFormats(sheet, areas))
  }
  deleteCondRule(sheet: number, id: number) {
    return this.run(() => this.doc.deleteCondRule(sheet, id))
  }
  moveCondRule(sheet: number, id: number, dir: -1 | 1) {
    return this.run(() => this.doc.moveCondRule(sheet, id, dir))
  }
  listCondFormats(sheet: number) {
    return this.doc.listCondFormats(sheet)
  }
  fill(a: Area, dir: 'down' | 'right' | 'up' | 'left') {
    return this.run(() => this.doc.fill(a, dir))
  }
  autofill(src: Area, dst: Area) {
    return this.run(() => autofill(this.doc, src, dst))
  }
  sort(a: Area, keys: SortKey[], header: boolean) {
    return this.run(() => sortRange(this.doc, a, keys, header))
  }
  find(opts: FindOptions) {
    return findCells(this.doc, opts)
  }
  replace(opts: FindOptions & { replacement: string; all: boolean; at?: { sheet: number; r: number; c: number } }) {
    let n = 0
    const r = this.run(() => (n = replaceCells(this.doc, opts)))
    return { ...r, count: n }
  }
  setCalc(mode: 'auto' | 'manual', iterate?: boolean) {
    this.doc.wb.calc.mode = mode
    if (iterate !== undefined) this.doc.wb.calc.iterate = iterate
    if (mode === 'auto') this.doc.recalc(true)
    return { ok: true }
  }
  calculate(full = false) {
    if (full) this.doc.recalcAll()
    else this.doc.calculateNow()
    return { ok: true }
  }
  /** Cells referenced by a formula (Trace Precedents) and cells that use a cell (Trace Dependents). */
  trace(sheet: number, r: number, c: number) {
    return {
      precedents: this.doc.engine.precedents(sheet, r, c),
      dependents: this.doc.engine.dependents(sheet, r, c).map(fromKey)
    }
  }

  /** Value of one cell (tests and macros). */
  value(sheet: number, r: number, c: number): Scalar {
    return this.doc.value(sheet, r, c)
  }
}

export const LIMITS = { MAX_ROWS, MAX_COLS }
