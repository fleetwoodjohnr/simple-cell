// A workbook being edited: the model, the calculation engine, and undoable transactions.
// Every change goes through `transact`, which records how to revert it.

import { Engine } from './engine/engine'
import './engine/functions'
import { FormulaSyntaxError, parseFormula, printFormula, type Node } from './engine/parser'
import { CellError, ERR, MAX_COLS, MAX_ROWS, isError, numberToText, setDateParser, type Area, type Scalar } from './engine/values'
import { Sheet, Workbook, type Cell, type ColInfo, type DefinedName, type RowInfo } from './model/workbook'
import type { Style } from './model/styles'
import { adjustFormula, relocate, shiftArea, transposeRefs, type Shift } from './model/adjust'
import { areaName, cellKey, cellName, fromKey, quoteSheet } from './refs'
import { parseInput } from './input'
import { formatValue, isDateFormat } from './numfmt/format'
import { parseDateTime } from './numfmt/dates'
import { CfEvaluator, newRuleId, type CfResult, type CfRule, type CondFormat } from './condfmt'
import { alertFor, isValid, listItems, validationAt, type DataValidation, type DvAlert } from './validation'
import { columnValues, hiddenRows, type AutoFilter, type FilterColumn } from './filter'
import { tableCellStyle } from './tablestyle'
import { learn, run as runProgram } from './flashfill'
import type { Link, Note, Table } from './model/workbook'

setDateParser((s) => parseDateTime(s))

interface Change {
  apply(): void
  revert(): void
}

export interface Transaction {
  label: string
  changes: Change[]
  /** Selection to restore on undo/redo (UI state, opaque to the core). */
  selection?: unknown
}

const cloneCell = (c: Cell | undefined): Cell | undefined => (c ? { ...c } : undefined)

export class Tx {
  readonly changes: Change[] = []
  /** Cells whose content changed (the engine is told after the transaction). */
  readonly touched = new Set<number>()
  structural = false
  /** Display changed beyond the edited cells (conditional formats): the client refetches what is on screen. */
  refetch = false
  /** Rows were hidden or shown: SUBTOTAL/AGGREGATE recalculate. */
  rowsShown = false
  namesChanged = new Set<string>()

  constructor(private doc: Doc) {}

  setCell(sheet: number, r: number, c: number, cell: Cell | undefined) {
    const sh = this.doc.wb.sheetById(sheet)
    if (!sh) return
    const before = cloneCell(sh.get(r, c))
    // A cell whose content is replaced is no longer part of a spill.
    const after = cell ? { ...cell } : undefined
    if (after && after.spillOf !== undefined && (after.f || after.v !== null)) delete after.spillOf
    const doc = this.doc
    const ch: Change = {
      apply: () => {
        doc.writeCell(sh, r, c, after ? { ...after } : undefined)
      },
      revert: () => {
        doc.writeCell(sh, r, c, before ? { ...before } : undefined)
      }
    }
    ch.apply()
    this.changes.push(ch)
    this.touched.add(cellKey(sheet, r, c))
  }

  /** Change a property of any model object (sheet name, freeze panes, workbook settings…). */
  prop<T extends object, K extends keyof T>(obj: T, key: K, value: T[K]) {
    const before = obj[key]
    if (before === value) return
    const ch: Change = {
      apply: () => {
        obj[key] = value
      },
      revert: () => {
        obj[key] = before
      }
    }
    ch.apply()
    this.changes.push(ch)
  }

  mapSet<K, V>(map: Map<K, V>, key: K, value: V | undefined) {
    const had = map.has(key)
    const before = map.get(key)
    const ch: Change = {
      apply: () => {
        if (value === undefined) map.delete(key)
        else map.set(key, value)
      },
      revert: () => {
        if (had) map.set(key, before as V)
        else map.delete(key)
      }
    }
    ch.apply()
    this.changes.push(ch)
  }

  custom(apply: () => void, revert: () => void) {
    apply()
    this.changes.push({ apply, revert })
  }
}

/** A conditional formatting rule as the UI describes it: formulas as text. */
export type CfSpec = Omit<CfRule, 'id' | 'priority' | 'formulas' | 'carry'> & { formulas?: string[] }

export interface CfListItem {
  id: number
  type: CfRule['type']
  priority: number
  ranges: string
  operator?: string
  text?: string
  rank?: number
  percent?: boolean
  bottom?: boolean
  aboveAverage?: boolean
  timePeriod?: string
  iconSet?: string
  colors?: string[]
  stopIfTrue?: boolean
  dxf?: Style
  formulas?: string[]
}

export type GoToKind = 'blanks' | 'constants' | 'formulas' | 'notes' | 'condFormats' | 'validation' | 'visible' | 'lastCell'

export interface PasteSpecialOptions {
  what: 'all' | 'formulas' | 'values' | 'formats' | 'notes' | 'validation' | 'allExceptBorders' | 'colWidths' | 'formulasNumFmt' | 'valuesNumFmt'
  op?: 'add' | 'subtract' | 'multiply' | 'divide'
  skipBlanks?: boolean
  transpose?: boolean
}

/** A data validation rule as the UI describes it: formulas as text (a list may be "a,b,c"). */
export type DvSpec = Omit<DataValidation, 'ranges' | 'formulas' | 'carry'> & { formulas?: string[] }
export type DvInfo = Omit<DvSpec, never> & { ranges: string; anchor: { r: number; c: number }; items?: string[] }

/** The parts of `a` outside `cut` (up to four rectangles). */
export function subtractArea(a: Area, cut: Area): Area[] {
  if (cut.sheet !== a.sheet || cut.r0 > a.r1 || cut.r1 < a.r0 || cut.c0 > a.c1 || cut.c1 < a.c0) return [a]
  const out: Area[] = []
  if (cut.r0 > a.r0) out.push({ ...a, r1: cut.r0 - 1 })
  if (cut.r1 < a.r1) out.push({ ...a, r0: cut.r1 + 1 })
  const r0 = Math.max(a.r0, cut.r0)
  const r1 = Math.min(a.r1, cut.r1)
  if (cut.c0 > a.c0) out.push({ ...a, r0, r1, c1: cut.c0 - 1 })
  if (cut.c1 < a.c1) out.push({ ...a, r0, r1, c0: cut.c1 + 1 })
  return out
}

export interface ViewCell {
  r: number
  c: number
  /** Displayed text. */
  t: string
  /** Style id. */
  s: number
  /** Value kind: n number, s text, b logical, e error, d date. */
  k: 'n' | 's' | 'b' | 'e' | 'd' | ''
  /** Colour from the number format ([Red]). */
  fc?: string
  /** Formula cell (for the "show formulas" view and error indicators). */
  f?: 1
  /** Part of a spilled array. */
  sp?: 1
  /** Unsupported function: value came from the file. */
  st?: 1
  /** Effective style when conditional formatting changes it (the cell's style with the rule's format on top). */
  x?: Style
  /** Conditional formatting data bar: fraction of the cell width, colour. */
  bar?: { w: number; color: string }
  /** Conditional formatting icon. */
  icon?: { set: string; i: number }
  /** The rule hides the value (data bar or icon only). */
  hide?: 1
  /** The cell has a note (red corner) or a hyperlink. */
  note?: 1
  link?: 1
}

export class Doc {
  readonly wb: Workbook
  readonly engine: Engine
  undoStack: Transaction[] = []
  redoStack: Transaction[] = []
  /** Bumped on every change, so callers can tell when to refresh. */
  version = 0
  /** Changed cell keys since the last call to `takeChanges`. */
  private changedKeys = new Set<number>()
  structureVersion = 0
  viewVersion = 0

  readonly cf: CfEvaluator

  constructor(wb?: Workbook) {
    this.wb = wb ?? Doc.blank()
    this.engine = new Engine(this.wb)
    this.cf = new CfEvaluator(this.engine, () => this.version)
    if (wb) {
      this.engine.rebuild()
      this.recalc()
    }
  }

  static blank(): Workbook {
    const wb = new Workbook()
    wb.addSheet('Sheet1')
    return wb
  }

  // ---------- low level ----------

  /** Store a cell and keep the engine's dependency graph in step (used by Tx changes). */
  writeCell(sh: Sheet, r: number, c: number, cell: Cell | undefined) {
    const key = cellKey(sh.id, r, c)
    const old = sh.get(r, c)
    if (old?.f) this.engine.unregister(key)
    if (old && this.engine.spills.has(key) && (!cell?.f || cell.f !== old.f)) this.engine.clearSpill(key)
    // Keep spilled-into state if the new cell is still empty.
    if (old?.spillOf !== undefined && cell && !cell.f && cell.v === null && cell.spillOf === undefined) cell.spillOf = old.spillOf
    sh.set(r, c, cell && (cell.v !== null || cell.f || cell.s || cell.spillOf !== undefined || cell.rich) ? cell : undefined)
    if (cell?.f) this.engine.register(sh.id, r, c, cell.f)
    this.engine.touch(sh.id, r, c)
    // An edit inside another formula's spill range blocks (or unblocks) it.
    const spillOf = old?.spillOf ?? cell?.spillOf
    if (spillOf !== undefined) {
      const a = fromKey(spillOf)
      this.engine.touch(a.sheet, a.r, a.c)
    }
    // Anchors showing #SPILL! may now have room.
    for (const k of this.engine.blocked) {
      const a = fromKey(k)
      if (a.sheet === sh.id) this.engine.touch(a.sheet, a.r, a.c)
    }
  }

  transact(label: string, fn: (tx: Tx) => void, selection?: unknown): Transaction | null {
    const tx = new Tx(this)
    try {
      fn(tx)
    } catch (e) {
      // Undo whatever was applied before the failure.
      for (let i = tx.changes.length - 1; i >= 0; i--) tx.changes[i].revert()
      if (tx.structural) this.afterStructure()
      this.recalc()
      throw e
    }
    if (!tx.changes.length) return null
    if (tx.structural) this.afterStructure()
    if (tx.refetch) this.viewVersion++
    if (tx.rowsShown) this.engine.rowsVisibilityChanged()
    for (const n of tx.namesChanged) this.engine.nameChanged(n)
    const t: Transaction = { label, changes: tx.changes, selection }
    this.undoStack.push(t)
    if (this.undoStack.length > 500) this.undoStack.shift() // ponytail: fixed depth; make it a setting if asked
    this.redoStack = []
    this.recalc()
    return t
  }

  undo(): Transaction | null {
    const t = this.undoStack.pop()
    if (!t) return null
    for (let i = t.changes.length - 1; i >= 0; i--) t.changes[i].revert()
    this.afterStructure()
    this.redoStack.push(t)
    this.recalc()
    return t
  }

  redo(): Transaction | null {
    const t = this.redoStack.pop()
    if (!t) return null
    for (const ch of t.changes) ch.apply()
    this.afterStructure()
    this.undoStack.push(t)
    this.recalc()
    return t
  }

  /** After rows/columns moved: rebuild dependency registrations (positions changed). */
  private afterStructure() {
    this.engine.rebuild()
    this.structureVersion++
  }

  recalc(force = false) {
    if (this.wb.calc.mode === 'auto' || force) this.engine.recalc({ volatile: true })
    for (const k of this.engine.takeChanged()) this.changedKeys.add(k)
    this.version++
  }

  /** Full recalculation (Ctrl+Alt+F9). */
  recalcAll() {
    this.engine.recalcAll()
    for (const k of this.engine.takeChanged()) this.changedKeys.add(k)
    this.version++
  }

  takeChanges(): number[] {
    const out = [...this.changedKeys]
    this.changedKeys.clear()
    return out
  }

  // ---------- reading ----------

  sheet(id: number): Sheet {
    const s = this.wb.sheetById(id)
    if (!s) throw new Error('No such sheet')
    return s
  }

  value(sheet: number, r: number, c: number): Scalar {
    return this.wb.sheetById(sheet)?.get(r, c)?.v ?? null
  }

  sheetIdByName = (n: string) => this.wb.sheetByName(n)?.id ?? null
  sheetName = (id: number) => this.wb.sheetById(id)?.name ?? null

  parse(text: string, sheet: number, r: number, c: number): Node {
    void sheet
    return parseFormula(text, { row: r, col: c, sheetId: this.sheetIdByName })
  }

  formulaText(sheet: number, r: number, c: number): string | null {
    const f = this.wb.sheetById(sheet)?.get(r, c)?.f
    if (!f) return null
    return '=' + printFormula(f.ast, { row: r, col: c, sheetName: this.sheetName })
  }

  /** What the formula bar shows for a cell. */
  editText(sheet: number, r: number, c: number): string {
    const sh = this.wb.sheetById(sheet)
    const cell = sh?.get(r, c)
    if (!cell || cell.spillOf !== undefined) return cell?.spillOf !== undefined ? (this.formulaText(sheet, fromKey(cell.spillOf).r, fromKey(cell.spillOf).c) ?? '') : ''
    if (cell.f) {
      const t = this.formulaText(sheet, r, c) ?? ''
      return cell.f.array ? `{${t}}` : t
    }
    const v = cell.v
    if (v === null) return ''
    if (typeof v === 'number') {
      const fmt = this.wb.styles.get(cell.s).numFmt
      if (isDateFormat(fmt)) {
        const hasTime = v % 1 !== 0
        const datePart = Math.floor(v) !== 0 || !hasTime
        const code = datePart ? (hasTime ? 'm/d/yyyy h:mm:ss AM/PM' : 'm/d/yyyy') : 'h:mm:ss AM/PM'
        return formatValue(v, code, this.wb.date1904).text
      }
      if (fmt && /%/.test(fmt.split(';')[0])) return numberToText(v * 100) + '%'
      return numberToText(v)
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
    if (v instanceof CellError) return v.code
    const s = this.wb.styles.get(cell.s)
    // Text that would otherwise turn into a number or formula keeps its apostrophe.
    if (s.carry?.quotePrefix || (parseInput(v).value !== v && !/^\s/.test(v))) return "'" + v
    return v
  }

  /** Conditional formatting for a window of a sheet (empty when the sheet has none). */
  condFormatsIn(sheet: number, win: Area): Map<number, CfResult> {
    const sh = this.wb.sheetById(sheet)
    if (!sh?.condFormats.length) return new Map()
    return this.cf.forArea(sh, sh.condFormats, win)
  }

  /** Displayed text and kind for one cell. `chars` is roughly how many characters fit (General format). */
  display(sheet: number, r: number, c: number, chars = 11): ViewCell | null {
    const sh = this.wb.sheetById(sheet)
    const cell = sh?.get(r, c)
    if (!cell) return null
    return this.viewOf(cell, r, c, chars)
  }

  /** View of one cell. `table` is the table style beneath the cell's own formatting; `cf` goes on top. */
  viewOf(cell: Cell, r: number, c: number, chars = 11, cf?: CfResult, table?: Style | null): ViewCell {
    const base = this.wb.styles.get(cell.s)
    const style = cf?.style || table ? { ...table, ...base, ...cf?.style } : base
    const v = cell.v
    const out: ViewCell = { r, c, t: '', s: cell.s ?? 0, k: '' }
    if (cell.f) out.f = 1
    if (cell.spillOf !== undefined) out.sp = 1
    if (cell.f?.unsupported && cell.stale) out.st = 1
    if (table && !cf?.style) out.x = style
    if (cf) {
      if (cf.style || table) out.x = style
      if (cf.bar) out.bar = cf.bar
      if (cf.icon) out.icon = cf.icon
      if (cf.hide) out.hide = 1
    }
    if (v === null) return out
    const fmt = formatValue(v, style.numFmt, this.wb.date1904, chars)
    out.t = fmt.text
    if (fmt.color) out.fc = fmt.color
    out.k = typeof v === 'number' ? (fmt.isDate ? 'd' : 'n') : typeof v === 'string' ? 's' : typeof v === 'boolean' ? 'b' : 'e'
    return out
  }

  // ---------- editing ----------

  /** Enter what the user typed into one cell (Enter in the cell or formula bar). */
  input(sheet: number, r: number, c: number, text: string, opts: { array?: Area } = {}) {
    return this.transact('Typing', (tx) => this.inputTx(tx, sheet, r, c, text, opts))
  }

  inputTx(tx: Tx, sheet: number, r: number, c: number, text: string, opts: { array?: Area } = {}) {
    const sh = this.sheet(sheet)
    const old = sh.get(r, c)
    const p = parseInput(text, this.wb.date1904)
    let s = old?.s
    if (p.formula) {
      let ast: Node
      try {
        ast = this.parse(p.formula, sheet, r, c)
      } catch (e) {
        if (e instanceof FormulaSyntaxError) throw e
        throw new FormulaSyntaxError(String(e), 0)
      }
      if (opts.array) {
        // Legacy array formula (Ctrl+Shift+Enter): one formula over the whole area.
        for (let rr = opts.array.r0; rr <= opts.array.r1; rr++)
          for (let cc = opts.array.c0; cc <= opts.array.c1; cc++) if (rr !== r || cc !== c) tx.setCell(sheet, rr, cc, sh.get(rr, cc)?.s ? { v: null, s: sh.get(rr, cc)!.s } : undefined)
        tx.setCell(sheet, opts.array.r0, opts.array.c0, { v: 0, f: { ast, array: opts.array }, s })
        return
      }
      tx.setCell(sheet, r, c, { v: 0, f: { ast }, s })
      return
    }
    if (p.format && !this.wb.styles.get(s).numFmt) s = this.wb.styles.derive(s, { numFmt: p.format })
    if (p.quoted) s = this.wb.styles.derive(s, { carry: { ...(this.wb.styles.get(s).carry ?? {}), quotePrefix: true } })
    else if (this.wb.styles.get(s).carry?.quotePrefix && typeof p.value !== 'string') s = this.wb.styles.derive(s, { carry: undefined })
    tx.setCell(sheet, r, c, p.value === null && !s ? undefined : { v: p.value ?? null, s })
    // Typing over a table's header cell renames the column.
    const t = this.wb.tableAt(sheet, r, c)
    if (t?.headerRow && r === t.area.r0) {
      const i = c - t.area.c0
      const typed = (this.display(sheet, r, c)?.t ?? '').trim() || `Column${i + 1}`
      let name = typed
      for (let k = 2; t.columns.some((x, j) => j !== i && x.name.toLowerCase() === name.toLowerCase()); k++) name = `${typed}${k}`
      if (name !== t.columns[i].name) {
        tx.prop(t, 'columns', t.columns.map((x, j) => (j === i ? { ...x, name } : x)))
        if (name !== typed) tx.setCell(sheet, r, c, { v: name, s })
        this.engine.tableChanged(t.name)
      }
    }
  }

  /** Clear contents, formats or both from areas (Delete key clears contents). */
  clear(areas: Area[], what: 'contents' | 'formats' | 'all' = 'contents') {
    return this.transact(what === 'contents' ? 'Clear contents' : what === 'formats' ? 'Clear formats' : 'Clear', (tx) => {
      for (const a of areas) {
        const sh = this.sheet(a.sheet)
        const cells: [number, number, Cell][] = []
        sh.forEachIn(a, (r, c, cell) => cells.push([r, c, cell]))
        for (const [r, c, cell] of cells) {
          if (cell.spillOf !== undefined && !cell.f && what === 'contents') continue
          if (what === 'contents') tx.setCell(a.sheet, r, c, cell.s ? { v: null, s: cell.s } : undefined)
          else if (what === 'formats') tx.setCell(a.sheet, r, c, { ...cell, s: undefined })
          else tx.setCell(a.sheet, r, c, undefined)
        }
        // Clear All also removes notes and hyperlinks.
        if (what === 'all')
          for (const m of [sh.notes, sh.links] as Map<number, unknown>[])
            for (const k of [...m.keys()]) {
              const r = Math.floor(k / 16384)
              const c = k % 16384
              if (r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1) tx.mapSet(m, k, undefined)
            }
      }
    })
  }

  /** Insert Cells: shift the cells at and beyond the area down or right (Ctrl+Plus). */
  insertCells(area: Area, dir: 'down' | 'right') {
    const isRow = dir === 'down'
    const s: Shift = { sheet: area.sheet, axis: isRow ? 'row' : 'col', at: isRow ? area.r0 : area.c0, count: isRow ? area.r1 - area.r0 + 1 : area.c1 - area.c0 + 1, span: isRow ? [area.c0, area.c1] : [area.r0, area.r1] }
    return this.transact('Insert cells', (tx) => this.shiftTx(tx, s))
  }

  /** Delete Cells: remove the area and shift the cells beyond it up or left (Ctrl+Minus). */
  deleteCells(area: Area, dir: 'up' | 'left') {
    const isRow = dir === 'up'
    const s: Shift = { sheet: area.sheet, axis: isRow ? 'row' : 'col', at: isRow ? area.r0 : area.c0, count: -(isRow ? area.r1 - area.r0 + 1 : area.c1 - area.c0 + 1), span: isRow ? [area.c0, area.c1] : [area.r0, area.r1] }
    return this.transact('Delete cells', (tx) => this.shiftTx(tx, s))
  }

  /**
   * Go To Special: the cells of a kind within `within` (the used range when it is a single cell), as row runs.
   * Types filter constants and formulas by value: n numbers, t text, l logicals, e errors.
   */
  goToSpecial(sheet: number, within: Area[], kind: GoToKind, types = 'ntle'): Area[] {
    const sh = this.sheet(sheet)
    const u = sh.used()
    if (kind === 'lastCell') return u.r < 0 ? [{ sheet, r0: 0, c0: 0, r1: 0, c1: 0 }] : [{ sheet, r0: u.r, c0: u.c, r1: u.r, c1: u.c }]
    const single = within.length === 1 && within[0].r0 === within[0].r1 && within[0].c0 === within[0].c1
    const areas = single ? [{ sheet, r0: 0, c0: 0, r1: Math.max(0, u.r), c1: Math.max(0, u.c) }] : within.map((a) => ({ ...a, r1: Math.min(a.r1, Math.max(u.r, a.r0)), c1: Math.min(a.c1, Math.max(u.c, a.c0)) }))
    const typeOf = (v: Scalar) => (typeof v === 'number' ? 'n' : typeof v === 'string' ? 't' : typeof v === 'boolean' ? 'l' : v instanceof CellError ? 'e' : '')
    const hiddenRow = (r: number) => !!sh.rowInfo.get(r)?.hidden || sh.rowInfo.get(r)?.height === 0
    const hiddenCol = (c: number) => !!sh.cols.get(c)?.hidden || sh.cols.get(c)?.width === 0
    const test = (r: number, c: number): boolean => {
      const cell = sh.get(r, c)
      switch (kind) {
        case 'blanks':
          return !cell || (cell.v === null && !cell.f) || cell.v === ''
        case 'constants':
          return !!cell && !cell.f && cell.spillOf === undefined && cell.v !== null && types.includes(typeOf(cell.v))
        case 'formulas':
          return !!cell && (!!cell.f || cell.spillOf !== undefined) && types.includes(typeOf(cell.v))
        case 'notes':
          return sh.notes.has(r * 16384 + c)
        case 'condFormats':
          return sh.condFormats.some((b) => b.ranges.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1))
        case 'validation':
          return !!validationAt(sh, r, c)
        case 'visible':
          return !hiddenRow(r) && !hiddenCol(c)
      }
    }
    const out: Area[] = []
    let cells = 0
    for (const a of areas)
      for (let r = a.r0; r <= a.r1; r++) {
        // Sparse kinds skip rows without content quickly.
        if ((kind === 'constants' || kind === 'formulas') && !sh.rows.get(r)?.size) continue
        let start = -1
        for (let c = a.c0; c <= a.c1 + 1; c++) {
          const ok = c <= a.c1 && test(r, c)
          if (ok && start < 0) start = c
          if (!ok && start >= 0) {
            const prev = out[out.length - 1]
            // Join with the run above when it spans the same columns.
            if (prev && prev.r1 === r - 1 && prev.c0 === start && prev.c1 === c - 1) prev.r1 = r
            else out.push({ sheet, r0: r, c0: start, r1: r, c1: c - 1 })
            cells += c - start
            start = -1
          }
        }
        if (out.length > 5000 || cells > 2_000_000) throw new Error('Too many cells match to select them all.')
      }
    if (!out.length) throw new Error('No cells were found.')
    return out
  }

  /**
   * Flash Fill (Ctrl+E) for the column of (r, c): learns from the column's filled cells, using the neighbouring
   * columns of the same rows as inputs, and fills the empty cells. Returns how many cells it filled.
   */
  flashFill(sheet: number, r: number, c: number): number {
    const sh = this.sheet(sheet)
    const has = (rr: number, cc: number) => {
      const v = sh.get(rr, cc)?.v
      return v !== null && v !== undefined && v !== ''
    }
    const text = (rr: number, cc: number) => this.display(sheet, rr, cc)?.t ?? ''
    // Input columns: the filled columns next to this one (on this row or the one above).
    const filled = (cc: number) => has(r, cc) || (r > 0 && has(r - 1, cc))
    const srcCols: number[] = []
    for (let cc = c - 1; cc >= 0 && filled(cc) && srcCols.length < 8; cc--) srcCols.unshift(cc)
    for (let cc = c + 1; cc < MAX_COLS && filled(cc) && srcCols.length < 16; cc++) srcCols.push(cc)
    if (!srcCols.length) throw new Error('Flash Fill needs data next to this column to learn from.')
    const rowHas = (rr: number) => srcCols.some((cc) => has(rr, cc))
    let r0 = r
    while (r0 > 0 && rowHas(r0 - 1)) r0--
    let r1 = r
    while (r1 < MAX_ROWS - 1 && rowHas(r1 + 1)) r1++
    // The first row is a header when its cell in this column is text and the inputs below are not like it.
    const examples: { inputs: string[]; output: string; r: number }[] = []
    for (let rr = r0; rr <= r1; rr++) if (has(rr, c)) examples.push({ r: rr, inputs: srcCols.map((cc) => text(rr, cc)), output: text(rr, c) })
    if (!examples.length) throw new Error('Type an example of the result in this column first, then press Ctrl+E.')
    let prog = learn(examples)
    let start = r0
    if (!prog && examples.length > 1 && examples[0].r === r0) {
      prog = learn(examples.slice(1))
      start = r0 + 1
    }
    if (!prog) throw new Error('Flash Fill couldn’t find a pattern. Type another example or two and try again.')
    const todo: { r: number; v: string }[] = []
    for (let rr = start; rr <= r1; rr++) {
      if (has(rr, c) || !rowHas(rr)) continue
      const v = runProgram(prog, srcCols.map((cc) => text(rr, cc)))
      if (v !== null && v !== '') todo.push({ r: rr, v })
    }
    if (!todo.length) throw new Error('There are no empty cells in this column to fill.')
    this.transact('Flash Fill', (tx) => {
      for (const t of todo) this.inputTx(tx, sheet, t.r, c, "'" + t.v)
    })
    return todo.length
  }

  // ---------- notes and hyperlinks ----------

  /** Add, change or (with null) delete the note on a cell. */
  setNote(sheet: number, r: number, c: number, text: string | null, author = '') {
    const sh = this.sheet(sheet)
    const k = r * 16384 + c
    const cur = sh.notes.get(k)
    return this.transact(text === null ? 'Delete note' : cur ? 'Edit note' : 'New note', (tx) => {
      tx.refetch = true
      tx.mapSet(sh.notes, k, text === null ? undefined : { author: cur?.author ?? author, text })
    })
  }

  /** Link the areas' cells (or remove links with null). `display` replaces the text of an empty or linked cell. */
  setLink(areas: Area[], link: Link | null, display?: string) {
    return this.transact(link ? 'Hyperlink' : 'Remove hyperlink', (tx) => {
      tx.refetch = true
      for (const a of areas) {
        const sh = this.sheet(a.sheet)
        if ((a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1) > 10_000) throw new Error('Select fewer cells to link (at most 10,000).')
        for (let r = a.r0; r <= a.r1; r++)
          for (let c = a.c0; c <= a.c1; c++) {
            tx.mapSet(sh.links, r * 16384 + c, link ?? undefined)
            const cell = sh.get(r, c)
            if (link) {
              const text = display ?? (cell && cell.v !== null ? undefined : (link.url ?? link.location ?? '').replace(/^mailto:/, ''))
              if (text !== undefined && !cell?.f) this.inputTx(tx, a.sheet, r, c, "'" + text)
              // The Hyperlink cell style: blue, underlined.
              const own = this.wb.styles.get(sh.get(r, c)?.s)
              if (!own.color && !own.underline) {
                const s = this.wb.styles.derive(sh.get(r, c)?.s, { color: '#0563C1', underline: 'single' })
                tx.setCell(a.sheet, r, c, { ...(sh.get(r, c) ?? { v: null }), s })
              }
            } else if (cell?.s) {
              const own = this.wb.styles.get(cell.s)
              if (own.color === '#0563C1' && own.underline === 'single') tx.setCell(a.sheet, r, c, { ...cell, s: this.wb.styles.derive(cell.s, { color: undefined, underline: undefined }) })
            }
          }
      }
    })
  }

  noteAt(sheet: number, r: number, c: number): Note | null {
    return this.sheet(sheet).notes.get(r * 16384 + c) ?? null
  }

  linkAt(sheet: number, r: number, c: number): Link | null {
    return this.sheet(sheet).links.get(r * 16384 + c) ?? null
  }

  /** Apply a style patch to areas. Whole rows/columns also set the row/column default style. */
  format(areas: Area[], patch: Partial<Style> | ((s: Style) => Partial<Style>), label = 'Format cells') {
    return this.transact(label, (tx) => {
      for (const a of areas) {
        const sh = this.sheet(a.sheet)
        const used = sh.used()
        const wholeCols = a.r0 === 0 && a.r1 === MAX_ROWS - 1
        const wholeRows = a.c0 === 0 && a.c1 === MAX_COLS - 1
        const r1 = wholeCols ? Math.max(used.r, -1) : a.r1
        const c1 = wholeRows ? Math.max(used.c, -1) : a.c1
        const p = (base: number | undefined) => this.wb.styles.derive(base, typeof patch === 'function' ? patch(this.wb.styles.get(base)) : patch)
        if (wholeCols) for (let c = a.c0; c <= a.c1; c++) tx.mapSet(sh.cols, c, { ...(sh.cols.get(c) ?? {}), style: p(sh.cols.get(c)?.style) })
        if (wholeRows) for (let r = a.r0; r <= a.r1; r++) tx.mapSet(sh.rowInfo, r, { ...(sh.rowInfo.get(r) ?? {}), style: p(sh.rowInfo.get(r)?.style) })
        if ((r1 - a.r0 + 1) * (c1 - a.c0 + 1) > 2_000_000) throw new Error('That range is too large to format cell by cell')
        for (let r = a.r0; r <= r1; r++)
          for (let c = a.c0; c <= c1; c++) {
            const cell = sh.get(r, c)
            const base = cell?.s ?? sh.rowInfo.get(r)?.style ?? sh.cols.get(c)?.style
            const s = p(base)
            if (cell) {
              if (cell.s !== s) tx.setCell(a.sheet, r, c, { ...cell, s })
            } else if (s !== 0 && !wholeCols && !wholeRows) tx.setCell(a.sheet, r, c, { v: null, s })
          }
      }
    })
  }

  /** Style of a cell including row/column defaults. */
  styleAt(sheet: number, r: number, c: number): Style {
    const sh = this.wb.sheetById(sheet)
    const s = sh?.get(r, c)?.s ?? sh?.rowInfo.get(r)?.style ?? sh?.cols.get(c)?.style
    return this.wb.styles.get(s)
  }

  // ---------- structure ----------

  insertRows(sheet: number, at: number, count: number) {
    return this.transact(count === 1 ? 'Insert row' : 'Insert rows', (tx) => this.shiftTx(tx, { sheet, axis: 'row', at, count }))
  }
  deleteRows(sheet: number, at: number, count: number) {
    return this.transact(count === 1 ? 'Delete row' : 'Delete rows', (tx) => this.shiftTx(tx, { sheet, axis: 'row', at, count: -count }))
  }
  insertCols(sheet: number, at: number, count: number) {
    return this.transact(count === 1 ? 'Insert column' : 'Insert columns', (tx) => this.shiftTx(tx, { sheet, axis: 'col', at, count }))
  }
  deleteCols(sheet: number, at: number, count: number) {
    return this.transact(count === 1 ? 'Delete column' : 'Delete columns', (tx) => this.shiftTx(tx, { sheet, axis: 'col', at, count: -count }))
  }

  /** Insert or delete rows/columns: moves cells, row/column info, merges, and rewrites every affected reference. */
  shiftTx(tx: Tx, s: Shift) {
    const sh = this.sheet(s.sheet)
    const isRow = s.axis === 'row'
    const limit = isRow ? MAX_ROWS : MAX_COLS
    if (s.count > 0) {
      // Refuse to push content off the sheet (Excel: "To prevent possible loss of data…").
      const u = sh.used()
      if ((isRow ? u.r : u.c) + s.count >= limit) throw new Error('Simple Cell can’t insert here: it would push cells off the end of the sheet.')
    }
    tx.structural = true
    const delEnd = s.count < 0 ? s.at - s.count - 1 : -1
    const inBand = (r: number, c: number) => !s.span || ((isRow ? c : r) >= s.span[0] && (isRow ? c : r) <= s.span[1])
    if (s.span) {
      // Excel refuses to split merged cells or tables with a cell shift.
      // Something straddling the band's edge, at or beyond the shift, would be split.
      const crosses = (a: Area) => {
        const [lo, hi] = isRow ? [a.c0, a.c1] : [a.r0, a.r1]
        const partial = lo <= s.span![1] && hi >= s.span![0] && (lo < s.span![0] || hi > s.span![1])
        return partial && (isRow ? a.r1 : a.c1) >= s.at
      }
      if (sh.merges.some(crosses)) throw new Error('Simple Cell can’t shift part of a merged cell. Unmerge it, or select whole merged cells.')
      if (this.wb.tables.some((t) => t.sheet === sh.id && crosses(t.area))) throw new Error('That would shift cells in a table. Insert or delete table rows or columns instead.')
    }
    // 1. Snapshot of every cell on this sheet with a new position, and formula rewrites everywhere.
    const moves: { from: [number, number]; to: [number, number] | null; cell: Cell }[] = []
    for (const [r, row] of sh.rows)
      for (const [c, cell] of row) {
        const idx = isRow ? r : c
        let to: [number, number] | null = [r, c]
        if (!inBand(r, c)) {
          moves.push({ from: [r, c], to, cell })
          continue
        }
        if (s.count > 0 && idx >= s.at) to = isRow ? [r + s.count, c] : [r, c + s.count]
        if (s.count < 0) {
          if (idx >= s.at && idx <= delEnd) to = null
          else if (idx > delEnd) to = isRow ? [r + s.count, c] : [r, c + s.count]
        }
        moves.push({ from: [r, c], to, cell })
      }
    // Formulas on other sheets that refer to this one.
    const others: { sheet: Sheet; r: number; c: number; cell: Cell; ast: Node }[] = []
    for (const o of this.wb.sheets) {
      if (o.id === sh.id) continue
      for (const [r, row] of o.rows)
        for (const [c, cell] of row) {
          if (!cell.f) continue
          const ast = adjustFormula(cell.f.ast, o.id, { row: r, col: c }, { row: r, col: c }, s)
          if (ast !== cell.f.ast) others.push({ sheet: o, r, c, cell, ast })
        }
    }
    const beforeRows = new Map(sh.rows)
    const afterRows = new Map<number, Map<number, Cell>>()
    for (const m of moves) {
      if (!m.to) continue
      let cell = m.cell
      if (cell.spillOf !== undefined && !cell.f) continue // spilled values are recomputed
      if (cell.f) {
        const ast = adjustFormula(cell.f.ast, sh.id, { row: m.from[0], col: m.from[1] }, { row: m.to[0], col: m.to[1] }, s)
        const array = cell.f.array ? (shiftArea(cell.f.array, s) ?? undefined) : undefined
        if (ast !== cell.f.ast || array !== cell.f.array) cell = { ...cell, f: { ...cell.f, ast, array } }
      }
      let row = afterRows.get(m.to[0])
      if (!row) afterRows.set(m.to[0], (row = new Map()))
      row.set(m.to[1], cell)
    }
    const swapRows = (rows: Map<number, Map<number, Cell>>) => {
      sh.rows.clear()
      for (const [r, row] of rows) sh.rows.set(r, row)
      sh.invalidateExtent()
    }
    tx.custom(
      () => swapRows(afterRows),
      () => swapRows(beforeRows)
    )
    for (const o of others) {
      const before = o.cell.f!
      const after = { ...before, ast: o.ast }
      tx.custom(
        () => (o.cell.f = after),
        () => (o.cell.f = before)
      )
    }
    // 2. Row/column info shifts (whole rows/columns only).
    const info = (isRow ? sh.rowInfo : sh.cols) as Map<number, RowInfo | ColInfo>
    if (!s.span) {
    const beforeInfo = new Map(info)
    const afterInfo = new Map<number, RowInfo | ColInfo>()
    for (const [i, v] of info) {
      if (s.count > 0) afterInfo.set(i >= s.at ? i + s.count : i, v)
      else if (i < s.at) afterInfo.set(i, v)
      else if (i > delEnd) afterInfo.set(i + s.count, v)
    }
    // Inserted rows/columns take the formatting of the one before them (Excel's default "Format same as above").
    if (s.count > 0 && s.at > 0 && info.has(s.at - 1)) {
      const prev = info.get(s.at - 1)!
      for (let i = 0; i < s.count; i++) afterInfo.set(s.at + i, { ...prev, hidden: false, outline: prev.outline })
    }
    const swapInfo = (m: Map<number, RowInfo | ColInfo>) => {
      info.clear()
      for (const [k, v] of m) info.set(k, v)
    }
    tx.custom(
      () => swapInfo(afterInfo),
      () => swapInfo(beforeInfo)
    )
    }
    // 3. Merges, freeze panes, names, tables.
    const merges = sh.merges.map((m) => shiftArea(m, s)).filter((m): m is Area => !!m && !(m.r0 === m.r1 && m.c0 === m.c1))
    tx.prop(sh, 'merges', merges)
    for (const n of this.wb.names) {
      if (!n.ast) continue
      const ast = adjustFormula(n.ast, n.scope ?? sh.id, { row: 0, col: 0 }, { row: 0, col: 0 }, s)
      if (ast !== n.ast) {
        tx.prop(n, 'ast', ast)
        tx.prop(n, 'text', printFormula(ast, { row: 0, col: 0, sheetName: this.sheetName }))
        tx.namesChanged.add(n.name.toLowerCase())
      }
    }
    for (const t of this.wb.tables) {
      if (t.sheet !== sh.id) continue
      const a = shiftArea(t.area, s)
      if (a) tx.prop(t, 'area', a)
    }
    // Notes and hyperlinks move with their cells; those on deleted cells go.
    for (const m of [sh.notes, sh.links] as Map<number, unknown>[]) {
      if (!m.size) continue
      const next = new Map<number, unknown>()
      for (const [k, v] of m) {
        const r = Math.floor(k / 16384)
        const c = k % 16384
        if (!inBand(r, c)) {
          next.set(k, v)
          continue
        }
        const a = shiftArea({ sheet: sh.id, r0: r, c0: c, r1: r, c1: c }, s)
        if (a) next.set(a.r0 * 16384 + a.c0, v)
      }
      const before = new Map(m)
      tx.custom(
        () => {
          m.clear()
          for (const [k, v] of next) m.set(k, v)
        },
        () => {
          m.clear()
          for (const [k, v] of before) m.set(k, v)
        }
      )
    }
    // Conditional formats: ranges move with the cells; formulas are relative to the first range's top-left cell.
    if (sh.condFormats.length) {
      const next: CondFormat[] = []
      for (const b of sh.condFormats) {
        const ranges = b.ranges.map((a) => shiftArea(a, s)).filter((a): a is Area => !!a)
        if (!ranges.length) continue
        const from = { row: b.ranges[0].r0, col: b.ranges[0].c0 }
        const to = { row: ranges[0].r0, col: ranges[0].c0 }
        const rules = b.rules.map((r) => (r.formulas ? { ...r, formulas: r.formulas.map((f) => adjustFormula(f, sh.id, from, to, s)) } : r))
        next.push({ ...b, ranges, rules })
      }
      tx.prop(sh, 'condFormats', next)
    }
    if (sh.validations.length) {
      const next: DataValidation[] = []
      for (const v of sh.validations) {
        const ranges = v.ranges.map((a) => shiftArea(a, s)).filter((a): a is Area => !!a)
        if (!ranges.length) continue
        const from = { row: v.ranges[0].r0, col: v.ranges[0].c0 }
        const to = { row: ranges[0].r0, col: ranges[0].c0 }
        next.push({ ...v, ranges, formulas: v.formulas?.map((f) => adjustFormula(f, sh.id, from, to, s)) })
      }
      tx.prop(sh, 'validations', next)
    }
    for (const [key, fn] of Object.entries(this.carryShifters)) {
      void key
      fn(tx, s)
    }
  }

  /** Format modules (conditional formats, validations, drawings…) register how their ranges move. */
  carryShifters: Record<string, (tx: Tx, s: Shift) => void> = {}

  // ---------- sheets ----------

  addSheet(name?: string, index?: number): number {
    const n = name ?? this.wb.uniqueSheetName('Sheet')
    if (this.wb.sheetByName(n)) throw new Error(`A sheet named “${n}” already exists.`)
    const s = this.wb.createSheet(n)
    this.transact('Insert sheet', (tx) => {
      const at = index ?? this.wb.sheets.length
      tx.custom(
        () => this.wb.sheets.splice(at, 0, s),
        () => this.wb.sheets.splice(this.wb.sheetIndex(s.id), 1)
      )
      tx.structural = true
    })
    return s.id
  }

  deleteSheet(id: number) {
    if (this.wb.sheets.filter((s) => !s.hidden).length <= 1 && !this.wb.sheetById(id)?.hidden) throw new Error('A workbook must contain at least one visible sheet.')
    return this.transact('Delete sheet', (tx) => {
      const idx = this.wb.sheetIndex(id)
      const s = this.wb.sheets[idx]
      tx.custom(
        () => this.wb.sheets.splice(idx, 1),
        () => this.wb.sheets.splice(idx, 0, s)
      )
      const names = this.wb.names.filter((n) => n.scope === id)
      if (names.length) tx.prop(this.wb, 'names', this.wb.names.filter((n) => n.scope !== id))
      const tables = this.wb.tables.filter((t) => t.sheet === id)
      if (tables.length) tx.prop(this.wb, 'tables', this.wb.tables.filter((t) => t.sheet !== id))
      if (this.wb.activeSheet >= this.wb.sheets.length) tx.prop(this.wb, 'activeSheet', this.wb.sheets.length - 1)
      tx.structural = true
    })
  }

  renameSheet(id: number, name: string) {
    const clean = name.trim()
    if (!clean || clean.length > 31 || /[\\/?*[\]:]/.test(clean) || clean.startsWith("'") || clean.endsWith("'")) throw new Error('Sheet names can’t be blank, longer than 31 characters, or contain \\ / ? * [ ] :')
    if (clean.toLowerCase() === 'history') throw new Error('“History” is a reserved sheet name.')
    const other = this.wb.sheetByName(clean)
    if (other && other.id !== id) throw new Error(`A sheet named “${clean}” already exists.`)
    return this.transact('Rename sheet', (tx) => tx.prop(this.sheet(id), 'name', clean))
  }

  moveSheet(id: number, to: number) {
    return this.transact('Move sheet', (tx) => {
      const from = this.wb.sheetIndex(id)
      if (from === to) return
      tx.custom(
        () => {
          const [s] = this.wb.sheets.splice(from, 1)
          this.wb.sheets.splice(to, 0, s)
        },
        () => {
          const [s] = this.wb.sheets.splice(to, 1)
          this.wb.sheets.splice(from, 0, s)
        }
      )
      tx.structural = true
    })
  }

  // ---------- rows & columns ----------

  setColWidths(sheet: number, cols: number[], width: number | undefined) {
    return this.transact('Column width', (tx) => {
      const sh = this.sheet(sheet)
      for (const c of cols) tx.mapSet(sh.cols, c, cleanInfo({ ...(sh.cols.get(c) ?? {}), width, hidden: width === 0 ? true : sh.cols.get(c)?.hidden && width === undefined ? true : undefined }))
    })
  }

  setRowHeights(sheet: number, rows: number[], height: number | undefined) {
    return this.transact('Row height', (tx) => {
      const sh = this.sheet(sheet)
      for (const r of rows) tx.mapSet(sh.rowInfo, r, cleanInfo({ ...(sh.rowInfo.get(r) ?? {}), height, custom: height !== undefined || undefined, hidden: height === 0 ? true : undefined }))
    })
  }

  hide(sheet: number, axis: 'row' | 'col', indexes: number[], hidden: boolean) {
    return this.transact(hidden ? 'Hide' : 'Unhide', (tx) => {
      if (axis === 'row') tx.rowsShown = true
      const sh = this.sheet(sheet)
      const map = (axis === 'row' ? sh.rowInfo : sh.cols) as Map<number, RowInfo | ColInfo>
      for (const i of indexes) {
        const cur = map.get(i) ?? {}
        const next: RowInfo = { ...cur, hidden: hidden || undefined }
        if (!hidden && (next.height === 0 || (next as ColInfo).width === 0)) {
          delete next.height
          delete (next as ColInfo).width
        }
        if (!hidden) delete next.filtered
        tx.mapSet(map, i, cleanInfo(next))
      }
    })
  }

  freeze(sheet: number, rows: number, cols: number) {
    return this.transact(rows || cols ? 'Freeze panes' : 'Unfreeze panes', (tx) => tx.prop(this.sheet(sheet), 'freeze', { rows, cols }))
  }

  merge(area: Area, mode: 'all' | 'across' | 'center' = 'all') {
    return this.transact('Merge cells', (tx) => {
      const sh = this.sheet(area.sheet)
      const parts: Area[] = mode === 'across' ? Array.from({ length: area.r1 - area.r0 + 1 }, (_, i) => ({ ...area, r0: area.r0 + i, r1: area.r0 + i })) : [area]
      let merges = sh.merges.filter((m) => !(m.r0 <= area.r1 && m.r1 >= area.r0 && m.c0 <= area.c1 && m.c1 >= area.c0))
      for (const p of parts) {
        if (p.r0 === p.r1 && p.c0 === p.c1) continue
        // Keep only the top-left value (Excel warns; the UI asks first).
        sh.forEachIn(p, (r, c, cell) => {
          if ((r !== p.r0 || c !== p.c0) && (cell.v !== null || cell.f)) tx.setCell(p.sheet, r, c, cell.s ? { v: null, s: cell.s } : undefined)
        })
        merges = [...merges, p]
      }
      tx.prop(sh, 'merges', merges)
      if (mode === 'center') {
        const s = this.wb.styles.derive(sh.get(area.r0, area.c0)?.s, { hAlign: 'center' })
        const cell = sh.get(area.r0, area.c0)
        tx.setCell(area.sheet, area.r0, area.c0, { ...(cell ?? { v: null }), s })
      }
    })
  }

  unmerge(area: Area) {
    return this.transact('Unmerge cells', (tx) => {
      const sh = this.sheet(area.sheet)
      tx.prop(
        sh,
        'merges',
        sh.merges.filter((m) => !(m.r0 <= area.r1 && m.r1 >= area.r0 && m.c0 <= area.c1 && m.c1 >= area.c0))
      )
    })
  }

  // ---------- names ----------

  // ---------- tables and filters ----------

  /** The filter at a key: 'sheet' for the sheet's AutoFilter, or a table name. */
  filterOf(sheet: number, key: string): { af: AutoFilter; table?: Table } | null {
    const sh = this.sheet(sheet)
    if (key === 'sheet') return sh.autoFilter ? { af: sh.autoFilter } : null
    const t = this.wb.tableByName(key)
    if (!t || t.sheet !== sheet || !t.headerRow || t.autoFilter === false) return null
    return { af: { area: { ...t.area, r1: t.totalsRow ? t.area.r1 - 1 : t.area.r1 }, columns: t.filter ?? [] }, table: t }
  }

  /**
   * Hide the rows any filter on the sheet excludes and show filtered rows that every filter now passes (a row is
   * hidden if any filter hides it). Call after changing the model; `extra` adds ranges whose filter was removed.
   */
  private refilterTx(tx: Tx, sheet: number, extra: Area[] = []) {
    const sh = this.sheet(sheet)
    const hide = new Set<number>()
    const rows = new Set<number>()
    const filters = this.filtersOn(sheet).map((f) => this.filterOf(sheet, f.key)!.af)
    for (const a of [...filters.map((f) => f.area), ...extra]) for (let r = a.r0 + 1; r <= a.r1; r++) rows.add(r)
    for (const af of filters) for (const r of hiddenRows(this.wb, sh, af)) hide.add(r)
    tx.rowsShown = true
    for (const r of rows) {
      const cur = sh.rowInfo.get(r)
      if (hide.has(r)) {
        if (!cur?.hidden || !cur.filtered) tx.mapSet(sh.rowInfo, r, { ...cur, hidden: true, filtered: true })
      } else if (cur?.filtered) tx.mapSet(sh.rowInfo, r, cleanInfo({ ...cur, hidden: undefined, filtered: undefined }))
    }
  }

  /** Set (or clear, with null) one column's criteria and filter. */
  setFilter(sheet: number, key: string, col: number, fc: Omit<FilterColumn, 'col'> | null) {
    const f = this.filterOf(sheet, key)
    if (!f) throw new Error('There is no filter here.')
    const columns = [...f.af.columns.filter((x) => x.col !== col), ...(fc ? [{ ...fc, col }] : [])].sort((a, b) => a.col - b.col)
    return this.transact(fc ? 'Filter' : 'Clear filter', (tx) => {
      if (f.table) tx.prop(f.table, 'filter', columns)
      else tx.prop(this.sheet(sheet), 'autoFilter', { ...f.af, columns })
      this.refilterTx(tx, sheet)
    })
  }

  clearFilters(sheet: number, key: string) {
    const f = this.filterOf(sheet, key)
    if (!f) return null
    return this.transact('Clear filters', (tx) => {
      if (f.table) tx.prop(f.table, 'filter', [])
      else tx.prop(this.sheet(sheet), 'autoFilter', { ...f.af, columns: [] })
      this.refilterTx(tx, sheet)
    })
  }

  /** Filter again with the same criteria (after the data changed). */
  reapplyFilter(sheet: number, key: string) {
    const f = this.filterOf(sheet, key)
    if (!f) return null
    return this.transact('Reapply filter', (tx) => this.refilterTx(tx, sheet))
  }

  /** Turn the sheet's AutoFilter on over `area` (header row first), or off when there is one. */
  toggleAutoFilter(sheet: number, area: Area | null) {
    const sh = this.sheet(sheet)
    return this.transact(sh.autoFilter ? 'Remove filter' : 'Filter', (tx) => {
      if (sh.autoFilter) {
        const area = sh.autoFilter.area
        tx.prop(sh, 'autoFilter', null)
        this.refilterTx(tx, sheet, [area])
      } else if (area) tx.prop(sh, 'autoFilter', { area: { ...area, sheet }, columns: [] })
    })
  }

  /** Values for a column's filter checklist, and its current criteria. */
  filterValues(sheet: number, key: string, col: number) {
    const f = this.filterOf(sheet, key)
    if (!f) return null
    return { ...columnValues(this.wb, this.sheet(sheet), f.af, col), current: f.af.columns.find((x) => x.col === col) ?? null }
  }

  /** Filter ranges on a sheet, for drawing their header buttons. */
  filtersOn(sheet: number): { key: string; area: Area; active: number[] }[] {
    const sh = this.sheet(sheet)
    const out: { key: string; area: Area; active: number[] }[] = []
    if (sh.autoFilter) out.push({ key: 'sheet', area: sh.autoFilter.area, active: sh.autoFilter.columns.map((x) => x.col) })
    for (const t of this.wb.tables) {
      if (t.sheet !== sheet) continue
      const f = this.filterOf(sheet, t.name)
      if (f) out.push({ key: t.name, area: f.af.area, active: f.af.columns.map((x) => x.col) })
    }
    return out
  }

  /** Format as Table (Ctrl+T). Without headers, a header row is inserted above the data. */
  createTable(sheet: number, area: Area, hasHeaders: boolean, style = 'TableStyleMedium2') {
    const sh = this.sheet(sheet)
    if (this.wb.tables.some((t) => t.sheet === sheet && t.area.r0 <= area.r1 && t.area.r1 >= area.r0 && t.area.c0 <= area.c1 && t.area.c1 >= area.c0)) throw new Error('A table can’t overlap another table.')
    if (sh.merges.some((m) => m.r0 <= area.r1 && m.r1 >= area.r0 && m.c0 <= area.c1 && m.c1 >= area.c0)) throw new Error('A table can’t contain merged cells. Unmerge them first.')
    let n = 1
    while (this.wb.tableByName(`Table${n}`) || this.wb.findName(`Table${n}`, null)) n++
    return this.transact('Create table', (tx) => {
      tx.refetch = true
      let a = { ...area, sheet }
      if (!hasHeaders) {
        // ponytail: inserts a whole sheet row (Excel shifts only the table's columns down).
        this.shiftTx(tx, { sheet, axis: 'row', at: area.r0, count: 1 })
        a = { ...a, r1: a.r1 + 1 }
      }
      const names: string[] = []
      for (let c = a.c0; c <= a.c1; c++) {
        let name = hasHeaders ? this.display(sheet, a.r0, c)?.t.trim() || '' : ''
        if (!name) name = `Column${c - a.c0 + 1}`
        let k = 2
        const base = name
        while (names.some((x) => x.toLowerCase() === name.toLowerCase())) name = `${base}${k++}`
        names.push(name)
        if (!hasHeaders || this.display(sheet, a.r0, c)?.t !== name) this.inputTx(tx, sheet, a.r0, c, "'" + name)
      }
      const t: Table = { id: Math.max(0, ...this.wb.tables.map((x) => x.id)) + 1, name: `Table${n}`, displayName: `Table${n}`, sheet, area: a, headerRow: true, totalsRow: false, columns: names.map((name) => ({ name })), style, showRowStripes: true, autoFilter: true, filter: [] }
      tx.prop(this.wb, 'tables', [...this.wb.tables, t])
      // A sheet AutoFilter over the same cells gives way to the table's.
      if (sh.autoFilter && sh.autoFilter.area.r0 <= a.r1 && sh.autoFilter.area.r1 >= a.r0 && sh.autoFilter.area.c0 <= a.c1 && sh.autoFilter.area.c1 >= a.c0) tx.prop(sh, 'autoFilter', null)
      this.engine.tableChanged(t.name)
    })
  }

  /** Table Design options: style, stripes, first/last column, filter buttons, total row, name. */
  setTableOptions(name: string, patch: { style?: string; showRowStripes?: boolean; showColumnStripes?: boolean; showFirstColumn?: boolean; showLastColumn?: boolean; autoFilter?: boolean; totalsRow?: boolean; name?: string }) {
    const t = this.wb.tableByName(name)
    if (!t) throw new Error(`There is no table called “${name}”.`)
    if (patch.name !== undefined) {
      const nn = patch.name.trim()
      if (!/^[A-Za-z_\\][\w.\\]*$/.test(nn) || /^[A-Za-z]{1,3}\d+$/.test(nn)) throw new Error(`“${nn}” isn’t a valid table name.`)
      if (nn.toLowerCase() !== t.name.toLowerCase() && (this.wb.tableByName(nn) || this.wb.findName(nn, null))) throw new Error(`The name “${nn}” is already in use.`)
    }
    return this.transact('Table options', (tx) => {
      tx.refetch = true
      for (const k of ['style', 'showRowStripes', 'showColumnStripes', 'showFirstColumn', 'showLastColumn', 'autoFilter'] as const) if (patch[k] !== undefined) tx.prop(t, k, patch[k] as never)
      if (patch.autoFilter === false && t.filter?.length) {
        tx.prop(t, 'filter', [])
        this.refilterTx(tx, t.sheet, [t.area])
      }
      if (patch.name !== undefined && patch.name.trim() !== t.name) {
        const old = t.name
        tx.prop(t, 'name', patch.name.trim())
        tx.prop(t, 'displayName', patch.name.trim())
        this.engine.tableChanged(old)
        this.engine.tableChanged(t.name)
      }
      if (patch.totalsRow !== undefined && patch.totalsRow !== t.totalsRow) this.totalsRowTx(tx, t, patch.totalsRow)
    })
  }

  /** Add or remove the total row: "Total" in the first column, SUBTOTAL in the last (sum for numbers, count otherwise). */
  private totalsRowTx(tx: Tx, t: Table, on: boolean) {
    const sh = this.sheet(t.sheet)
    const { r1, c0, c1 } = t.area
    if (on) {
      const below = r1 + 1
      let busy = false
      for (let c = c0; c <= c1; c++) if (sh.get(below, c)) busy = true
      if (busy) this.shiftTx(tx, { sheet: t.sheet, axis: 'row', at: below, count: 1 })
      const lastName = t.columns[t.columns.length - 1].name
      let numeric = true
      for (let r = t.area.r0 + 1; r <= r1; r++) {
        const v = sh.get(r, c1)?.v
        if (v !== null && v !== undefined && typeof v !== 'number') numeric = false
      }
      const fn = numeric ? 'sum' : 'count'
      tx.prop(t, 'area', { ...t.area, r1: below })
      tx.prop(t, 'totalsRow', true)
      tx.prop(t, 'columns', t.columns.map((col, i) => (i === 0 && c1 > c0 ? { ...col, totalsLabel: 'Total' } : i === t.columns.length - 1 ? { ...col, totalsFunction: fn } : col)))
      this.engine.tableChanged(t.name)
      if (c1 > c0) this.inputTx(tx, t.sheet, below, c0, 'Total')
      this.inputTx(tx, t.sheet, below, c1, `=SUBTOTAL(${numeric ? 109 : 103},${t.name}[${lastName.replace(/([[\]#'])/g, "'$1")}])`)
    } else {
      for (let c = c0; c <= c1; c++) if (sh.get(r1, c)) tx.setCell(t.sheet, r1, c, undefined)
      tx.prop(t, 'area', { ...t.area, r1: r1 - 1 })
      tx.prop(t, 'totalsRow', false)
      tx.prop(t, 'columns', t.columns.map(({ totalsFunction: _f, totalsLabel: _l, ...col }) => col))
      this.engine.tableChanged(t.name)
    }
  }

  /** Convert to Range: the table goes; its look is kept as cell formatting. Formulas keep their results' references. */
  convertTableToRange(name: string) {
    const t = this.wb.tableByName(name)
    if (!t) throw new Error(`There is no table called “${name}”.`)
    const sh = this.sheet(t.sheet)
    return this.transact('Convert to range', (tx) => {
      tx.refetch = true
      for (let r = t.area.r0; r <= t.area.r1; r++)
        for (let c = t.area.c0; c <= t.area.c1; c++) {
          const ts = tableCellStyle(t, r, c, this.wb.theme ?? undefined)
          if (!ts) continue
          const cell = sh.get(r, c)
          const own = this.wb.styles.get(cell?.s)
          const s = this.wb.styles.intern({ ...ts, ...own })
          tx.setCell(t.sheet, r, c, cell ? { ...cell, s } : { v: null, s })
        }
      // Structured references become ordinary ones.
      for (const o of this.wb.sheets)
        for (const [r, row] of o.rows)
          for (const [c, cell] of row) {
            if (!cell.f || !printFormula(cell.f.ast, { row: r, col: c, sheetName: this.sheetName }).includes('[')) continue
            const text = this.formulaText(o.id, r, c)
            if (!text) continue
            const plain = this.unstructure(text, t, o.id, r, c)
            if (plain !== text) this.inputTx(tx, o.id, r, c, plain)
          }
      tx.prop(this.wb, 'tables', this.wb.tables.filter((x) => x !== t))
      if (t.filter?.length) this.refilterTx(tx, t.sheet, [t.area])
      this.engine.tableChanged(t.name)
    })
  }

  /** Replace structured references to table `t` in formula text with A1 references. */
  private unstructure(text: string, t: Table, sheet: number, r: number, c: number): string {
    const name = t.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const q = (a: Area) => (a.sheet === sheet ? '' : quoteSheet(this.sheetName(a.sheet) ?? '') + '!') + areaName(a).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')
    const dataArea = { ...t.area, r0: t.headerRow ? t.area.r0 + 1 : t.area.r0, r1: t.totalsRow ? t.area.r1 - 1 : t.area.r1 }
    const col = (n: string) => t.columns.findIndex((x) => x.name.toLowerCase() === n.replace(/'(.)/g, '$1').toLowerCase())
    return text.replace(new RegExp(`${name}\\[(\\[[^\\]]*\\]|[^\\[\\]]*)\\]`, 'gi'), (m, inner: string) => {
      const spec = inner.replace(/^\[|\]$/g, '')
      if (spec === '' || /^#data$/i.test(spec)) return q(dataArea)
      if (/^#all$/i.test(spec)) return q(t.area)
      if (/^#headers$/i.test(spec)) return q({ ...t.area, r1: t.area.r0 })
      if (/^#totals$/i.test(spec)) return q({ ...t.area, r0: t.area.r1 })
      const i = col(spec.replace(/^@/, ''))
      if (i < 0) return m
      if (spec.startsWith('@')) return cellName(r, t.area.c0 + i)
      return q({ ...dataArea, c0: t.area.c0 + i, c1: t.area.c0 + i })
    })
  }

  /** Tables and filters for the sheet's meta. */
  tablesOn(sheet: number) {
    return this.wb.tables.filter((t) => t.sheet === sheet).map((t) => ({ name: t.name, area: t.area, style: t.style ?? '', showRowStripes: t.showRowStripes !== false, showColumnStripes: !!t.showColumnStripes, showFirstColumn: !!t.showFirstColumn, showLastColumn: !!t.showLastColumn, totalsRow: t.totalsRow, autoFilter: t.autoFilter !== false && t.headerRow }))
  }

  /** The table style's formatting at a cell (null outside tables). */
  tableStyleAt(sheet: number, r: number, c: number) {
    const t = this.wb.tableAt(sheet, r, c)
    return t ? tableCellStyle(t, r, c, this.wb.theme ?? undefined) : null
  }

  // ---------- data validation ----------

  /** Set the rule for the areas (replacing any rule there). `spec.type` 'none' with no messages clears them. */
  setValidation(sheet: number, areas: Area[], spec: DvSpec) {
    const sh = this.sheet(sheet)
    const anchor = areas[0]
    const formulas = spec.formulas?.filter((f) => f.trim() !== '').map((f) => {
      const t = f.trim()
      // A list typed as items ("Yes, No") rather than a reference or formula.
      if (spec.type === 'list' && !t.startsWith('=')) return { t: 'str', v: t } as Node
      return this.parse(t.startsWith('=') ? t : '=' + t, sheet, anchor.r0, anchor.c0)
    })
    return this.transact('Data validation', (tx) => {
      const next: DataValidation[] = []
      for (const v of sh.validations) {
        let ranges = v.ranges
        for (const cut of areas) ranges = ranges.flatMap((a) => subtractArea(a, cut))
        if (ranges.length) next.push(ranges === v.ranges ? v : { ...v, ranges })
      }
      const keep = spec.type !== 'none' || spec.prompt || spec.promptTitle
      if (keep) {
        const dv: DataValidation = { ...spec, ranges: areas.map((a) => ({ ...a, sheet })), formulas }
        if (!formulas?.length) delete dv.formulas
        next.push(dv)
      }
      tx.prop(sh, 'validations', next)
    })
  }

  /** The rule at a cell, with formulas as text (for the dialog), its dropdown items and input message. */
  validationInfo(sheet: number, r: number, c: number): DvInfo | null {
    const sh = this.sheet(sheet)
    const dv = validationAt(sh, r, c)
    if (!dv) return null
    const a = dv.ranges[0]
    const formulas = dv.formulas?.map((f) => (f.t === 'str' && dv.type === 'list' ? f.v : '=' + printFormula(f, { row: r, col: c, sheetName: this.sheetName })))
    const { ranges: _r, formulas: _f, carry: _c, ...rest } = dv
    void _r
    void _f
    void _c
    const info: DvInfo = { ...rest, formulas, ranges: dv.ranges.map((x) => areaName(x)).join(' '), anchor: { r: a.r0, c: a.c0 } }
    if (dv.type === 'list' && !dv.noDropdown) info.items = listItems(this.engine, sh, dv, r, c).map((x) => x.t)
    return info
  }

  /** The alert to show if the cell's value breaks its rule (null when it is fine). */
  checkValidation(sheet: number, r: number, c: number): DvAlert | null {
    const sh = this.sheet(sheet)
    const dv = validationAt(sh, r, c)
    if (!dv || dv.showError === false || isValid(this.engine, sh, dv, r, c)) return null
    return alertFor(dv)
  }

  /** Cells in the areas whose values break their rules (Circle Invalid Data). */
  invalidCells(sheet: number): { r: number; c: number }[] {
    const sh = this.sheet(sheet)
    const out: { r: number; c: number }[] = []
    for (const dv of sh.validations)
      for (const a of dv.ranges)
        sh.forEachIn(a, (r, c) => {
          if (validationAt(sh, r, c) === dv && !isValid(this.engine, sh, dv, r, c)) out.push({ r, c })
        })
    return out
  }

  // ---------- conditional formatting ----------

  /** Add a rule over `ranges` (formulas as text, relative to the first range's top-left cell). New rules come first. */
  addCondFormat(sheet: number, ranges: Area[], spec: CfSpec) {
    const sh = this.sheet(sheet)
    const anchor = ranges[0]
    const formulas = spec.formulas?.map((f) => this.parse(f.startsWith('=') ? f : '=' + f, sheet, anchor.r0, anchor.c0))
    const rule: CfRule = { ...spec, formulas, id: newRuleId(), priority: 1 }
    if (!formulas) delete rule.formulas
    return this.transact('Conditional formatting', (tx) => {
      tx.refetch = true
      const bumped = sh.condFormats.map((b) => ({ ...b, rules: b.rules.map((r) => ({ ...r, priority: r.priority + 1 })) }))
      tx.prop(sh, 'condFormats', [...bumped, { ranges: ranges.map((a) => ({ ...a, sheet })), rules: [rule] }])
    })
  }

  /** Remove conditional formatting from an area of the sheet, or from the whole sheet. */
  clearCondFormats(sheet: number, areas: Area[] | null) {
    const sh = this.sheet(sheet)
    return this.transact('Clear rules', (tx) => {
      tx.refetch = true
      if (!areas) return tx.prop(sh, 'condFormats', [])
      const next: CondFormat[] = []
      for (const b of sh.condFormats) {
        let ranges = b.ranges
        for (const cut of areas) ranges = ranges.flatMap((a) => subtractArea(a, cut))
        if (ranges.length) next.push(ranges === b.ranges ? b : { ...b, ranges })
      }
      tx.prop(sh, 'condFormats', next)
    })
  }

  deleteCondRule(sheet: number, id: number) {
    const sh = this.sheet(sheet)
    return this.transact('Delete rule', (tx) => {
      tx.refetch = true
      tx.prop(
        sh,
        'condFormats',
        sh.condFormats.map((b) => ({ ...b, rules: b.rules.filter((r) => r.id !== id) })).filter((b) => b.rules.length)
      )
    })
  }

  /** Move a rule up (-1, higher priority) or down (+1). */
  moveCondRule(sheet: number, id: number, dir: -1 | 1) {
    const sh = this.sheet(sheet)
    const order = sh.condFormats.flatMap((b) => b.rules).sort((a, b) => a.priority - b.priority)
    const i = order.findIndex((r) => r.id === id)
    const j = i + dir
    if (i < 0 || j < 0 || j >= order.length) return null
    ;[order[i], order[j]] = [order[j], order[i]]
    const prio = new Map(order.map((r, k) => [r.id, k + 1]))
    return this.transact('Move rule', (tx) => {
      tx.refetch = true
      tx.prop(
        sh,
        'condFormats',
        sh.condFormats.map((b) => ({ ...b, rules: b.rules.map((r) => ({ ...r, priority: prio.get(r.id)! })) }))
      )
    })
  }

  /** The sheet's rules in priority order, with formulas as text (for the Manage Rules dialog). */
  listCondFormats(sheet: number): CfListItem[] {
    const sh = this.sheet(sheet)
    return sh.condFormats
      .flatMap((b) =>
        b.rules.map((r) => ({
          id: r.id,
          type: r.type,
          priority: r.priority,
          ranges: b.ranges.map((a) => areaName(a)).join(' '),
          operator: r.operator,
          text: r.text,
          rank: r.rank,
          percent: r.percent,
          bottom: r.bottom,
          aboveAverage: r.aboveAverage,
          timePeriod: r.timePeriod,
          iconSet: r.iconSet?.set,
          colors: r.colorScale?.colors ?? (r.dataBar ? [r.dataBar.color] : undefined),
          stopIfTrue: r.stopIfTrue,
          dxf: r.dxf,
          formulas: r.formulas?.map((f) => '=' + printFormula(f, { row: b.ranges[0].r0, col: b.ranges[0].c0, sheetName: this.sheetName }))
        }))
      )
      .sort((a, b) => a.priority - b.priority)
  }

  defineName(name: string, text: string, scope: number | null = null, comment?: string) {
    const n = name.trim()
    if (!/^[A-Za-z_\\À-￿][\w.\\?À-￿]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^(R|C|R\d+C\d+)$/i.test(n) || n.length > 255) throw new Error(`“${n}” isn’t a valid name. Names start with a letter or underscore, contain no spaces, and can’t look like a cell reference.`)
    const src = text.startsWith('=') ? text : '=' + text
    const ast = this.parse(src, scope ?? 0, 0, 0)
    return this.transact('Define name', (tx) => {
      const existing = this.wb.names.find((d) => d.name.toLowerCase() === n.toLowerCase() && d.scope === scope)
      const def: DefinedName = { name: n, scope, text: src.slice(1), ast, comment }
      tx.prop(this.wb, 'names', [...this.wb.names.filter((d) => d !== existing), def])
      tx.namesChanged.add(n.toLowerCase())
    })
  }

  deleteName(name: string, scope: number | null) {
    return this.transact('Delete name', (tx) => {
      tx.prop(
        this.wb,
        'names',
        this.wb.names.filter((d) => !(d.name.toLowerCase() === name.toLowerCase() && d.scope === scope))
      )
      tx.namesChanged.add(name.toLowerCase())
    })
  }

  // ---------- copy / fill ----------

  /** Cells of an area as a block (for copy, fill and paste). Formulas keep their relative AST. */
  block(a: Area): { rows: number; cols: number; cells: (Cell | undefined)[] } {
    const sh = this.sheet(a.sheet)
    const rows = a.r1 - a.r0 + 1
    const cols = a.c1 - a.c0 + 1
    const cells: (Cell | undefined)[] = new Array(rows * cols)
    sh.forEachIn(a, (r, c, cell) => {
      const own = cell.spillOf !== undefined && !cell.f ? { v: cell.v, s: cell.s } : cell
      cells[(r - a.r0) * cols + (c - a.c0)] = { ...own, spillOf: undefined, stale: undefined }
    })
    return { rows, cols, cells }
  }

  /** Write a block into a destination (tiling it over a larger destination, like Excel's paste and fill). */
  pasteTx(tx: Tx, dst: Area, block: { rows: number; cols: number; cells: (Cell | undefined)[] }, what: 'all' | 'values' | 'formats' | 'formulas' = 'all') {
    const rows = Math.max(block.rows, Math.floor((dst.r1 - dst.r0 + 1) / block.rows) * block.rows)
    const cols = Math.max(block.cols, Math.floor((dst.c1 - dst.c0 + 1) / block.cols) * block.cols)
    if (dst.r0 + rows > MAX_ROWS || dst.c0 + cols > MAX_COLS) throw new Error('The paste area extends beyond the edge of the sheet.')
    const sh = this.sheet(dst.sheet)
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        const src = block.cells[(i % block.rows) * block.cols + (j % block.cols)]
        const r = dst.r0 + i
        const c = dst.c0 + j
        const cur = sh.get(r, c)
        if (what === 'formats') {
          if (src?.s !== cur?.s) tx.setCell(dst.sheet, r, c, cur ? { ...cur, s: src?.s } : src?.s ? { v: null, s: src.s } : undefined)
          continue
        }
        let next: Cell | undefined = src ? { ...src } : undefined
        if (next?.f) {
          if (what === 'values') next = { v: next.v, s: next.s }
          else next.f = { ...next.f, ast: relocate(next.f.ast, { row: r, col: c }), array: undefined }
        }
        if (what === 'values' || what === 'formulas') next = next ? { ...next, s: cur?.s } : cur?.s ? { v: null, s: cur.s } : undefined
        tx.setCell(dst.sheet, r, c, next)
      }
  }

  /**
   * Paste Special: which parts of the copied cells to paste, an arithmetic operation with what is there,
   * Skip Blanks and Transpose (Excel's dialog). `dst` is the top-left cell, or a larger area to tile.
   */
  pasteSpecial(src: Area, dst: Area, o: PasteSpecialOptions) {
    const from = this.sheet(src.sheet)
    const to = this.sheet(dst.sheet)
    let block = this.block(src)
    const srcAt = (i: number, j: number) => (o.transpose ? { r: src.r0 + j, c: src.c0 + i } : { r: src.r0 + i, c: src.c0 + j })
    if (o.transpose) {
      const t: (Cell | undefined)[] = new Array(block.rows * block.cols)
      for (let i = 0; i < block.rows; i++) for (let j = 0; j < block.cols; j++) t[j * block.rows + i] = block.cells[i * block.cols + j]
      block = { rows: block.cols, cols: block.rows, cells: t }
    }
    const rows = Math.max(block.rows, Math.floor((dst.r1 - dst.r0 + 1) / block.rows) * block.rows)
    const cols = Math.max(block.cols, Math.floor((dst.c1 - dst.c0 + 1) / block.cols) * block.cols)
    if (dst.r0 + rows > MAX_ROWS || dst.c0 + cols > MAX_COLS) throw new Error('The paste area extends beyond the edge of the sheet.')
    const what = o.what
    return this.transact('Paste special', (tx) => {
      tx.refetch = true
      if (what === 'colWidths') {
        for (let j = 0; j < cols; j++) {
          const sc = o.transpose ? src.c0 : src.c0 + (j % block.cols)
          tx.mapSet(to.cols, dst.c0 + j, cleanInfo({ ...(to.cols.get(dst.c0 + j) ?? {}), width: from.cols.get(sc)?.width }))
        }
        return
      }
      for (let i = 0; i < rows; i++)
        for (let j = 0; j < cols; j++) {
          const bi = i % block.rows
          const bj = j % block.cols
          const s = block.cells[bi * block.cols + bj]
          const r = dst.r0 + i
          const c = dst.c0 + j
          const cur = to.get(r, c)
          const sp = srcAt(bi, bj)
          if (o.skipBlanks && (!s || (s.v === null && !s.f))) continue
          if (what === 'notes') {
            tx.mapSet(to.notes, r * 16384 + c, from.notes.get(sp.r * 16384 + sp.c))
            continue
          }
          if (what === 'validation') continue
          if (what === 'formats') {
            tx.setCell(dst.sheet, r, c, cur ? { ...cur, s: s?.s } : s?.s ? { v: null, s: s.s } : undefined)
            continue
          }
          const srcStyle = this.wb.styles.get(s?.s)
          let style = cur?.s
          if (what === 'all') style = s?.s
          else if (what === 'allExceptBorders') style = this.wb.styles.derive(s?.s, { borderTop: undefined, borderBottom: undefined, borderLeft: undefined, borderRight: undefined, borderDiagUp: undefined, borderDiagDown: undefined })
          else if (what === 'formulasNumFmt' || what === 'valuesNumFmt') style = this.wb.styles.derive(cur?.s, { numFmt: srcStyle.numFmt })
          const asValue = what === 'values' || what === 'valuesNumFmt'
          let next: Cell | undefined = s ? (asValue || !s.f ? { v: s.v } : { v: s.v, f: { ...s.f, ast: relocate(o.transpose ? transposeRefs(s.f.ast) : s.f.ast, { row: r, col: c }), array: undefined } }) : { v: null }
          if (o.op) {
            // Excel: a non-number source pastes as it is; a non-number destination is left alone; a formula
            // destination becomes =(formula)op value; numbers are combined.
            const a = cur?.v ?? null
            const y = typeof next.v === 'number' ? next.v : next.v === null ? 0 : null
            if (y !== null) {
              if (typeof a === 'string' || typeof a === 'boolean' || a instanceof CellError) continue
              const sym = { add: '+', subtract: '-', multiply: '*', divide: '/' }[o.op]
              if (cur?.f) {
                this.inputTx(tx, dst.sheet, r, c, `=(${this.formulaText(dst.sheet, r, c)!.slice(1)})${sym}${numberToText(y)}`)
                continue
              }
              const x = a ?? 0
              next = { v: o.op === 'add' ? x + y : o.op === 'subtract' ? x - y : o.op === 'multiply' ? x * y : y === 0 ? ERR.DIV0 : x / y }
              style = cur?.s
            }
          }
          if (next.v === null && !next.f && !style) tx.setCell(dst.sheet, r, c, undefined)
          else tx.setCell(dst.sheet, r, c, { ...next, ...(style ? { s: style } : {}) })
        }
      if (what === 'validation' || what === 'all') {
        // Validation rules of the source cells cover the destination (all: as Excel does on a normal paste).
        for (const dv of from.validations) {
          if (!dv.ranges.some((a) => a.r0 <= src.r1 && a.r1 >= src.r0 && a.c0 <= src.c1 && a.c1 >= src.c0)) continue
          if (what === 'all' && src.sheet === dst.sheet) continue
          const area: Area = { sheet: dst.sheet, r0: dst.r0, c0: dst.c0, r1: dst.r0 + rows - 1, c1: dst.c0 + cols - 1 }
          const others = to.validations.flatMap((v) => {
            const ranges = v.ranges.flatMap((x) => subtractArea(x, area))
            return ranges.length ? [{ ...v, ranges }] : []
          })
          tx.prop(to, 'validations', [...others, { ...dv, ranges: [area], formulas: dv.formulas?.map((f) => relocate(f, { row: area.r0, col: area.c0 })) }])
        }
      }
    })
  }

  /** Fill down/right/up/left from the edge row/column of the selection (Ctrl+D / Ctrl+R). */
  fill(a: Area, dir: 'down' | 'right' | 'up' | 'left') {
    return this.transact(`Fill ${dir}`, (tx) => {
      const src: Area =
        dir === 'down' ? { ...a, r1: a.r0 } : dir === 'up' ? { ...a, r0: a.r1 } : dir === 'right' ? { ...a, c1: a.c0 } : { ...a, c0: a.c1 }
      const dst: Area =
        dir === 'down' ? { ...a, r0: a.r0 + 1 } : dir === 'up' ? { ...a, r1: a.r1 - 1 } : dir === 'right' ? { ...a, c0: a.c0 + 1 } : { ...a, c1: a.c1 - 1 }
      if (dst.r0 > dst.r1 || dst.c0 > dst.c1) return
      this.pasteTx(tx, dst, this.block(src))
    })
  }

  /** Recalculate (F9; in manual mode this is the only time formulas update). */
  calculateNow() {
    this.engine.recalc({ volatile: true })
    for (const k of this.engine.takeChanged()) this.changedKeys.add(k)
    this.version++
  }
}

function cleanInfo<T extends object>(o: T): T | undefined {
  const out = Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined && v !== false)) as T
  return Object.keys(out).length ? out : undefined
}

export { isError, ERR }
