// A workbook being edited: the model, the calculation engine, and undoable transactions.
// Every change goes through `transact`, which records how to revert it.

import { Engine } from './engine/engine'
import './engine/functions'
import { FormulaSyntaxError, parseFormula, printFormula, type Node } from './engine/parser'
import { CellError, ERR, MAX_COLS, MAX_ROWS, isError, numberToText, setDateParser, type Area, type Scalar } from './engine/values'
import { Sheet, Workbook, type Cell, type ColInfo, type DefinedName, type RowInfo } from './model/workbook'
import type { Style } from './model/styles'
import { adjustFormula, relocate, shiftArea, type Shift } from './model/adjust'
import { areaName, cellKey, fromKey } from './refs'
import { parseInput } from './input'
import { formatValue, isDateFormat } from './numfmt/format'
import { parseDateTime } from './numfmt/dates'
import { CfEvaluator, newRuleId, type CfResult, type CfRule, type CondFormat } from './condfmt'
import { alertFor, isValid, listItems, validationAt, type DataValidation, type DvAlert } from './validation'

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

  viewOf(cell: Cell, r: number, c: number, chars = 11, cf?: CfResult): ViewCell {
    const base = this.wb.styles.get(cell.s)
    const style = cf?.style ? { ...base, ...cf.style } : base
    const v = cell.v
    const out: ViewCell = { r, c, t: '', s: cell.s ?? 0, k: '' }
    if (cell.f) out.f = 1
    if (cell.spillOf !== undefined) out.sp = 1
    if (cell.f?.unsupported && cell.stale) out.st = 1
    if (cf) {
      if (cf.style) out.x = style
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
      }
    })
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
    // 1. Snapshot of every cell on this sheet with a new position, and formula rewrites everywhere.
    const moves: { from: [number, number]; to: [number, number] | null; cell: Cell }[] = []
    for (const [r, row] of sh.rows)
      for (const [c, cell] of row) {
        const idx = isRow ? r : c
        let to: [number, number] | null = [r, c]
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
    // 2. Row/column info shifts.
    const info = (isRow ? sh.rowInfo : sh.cols) as Map<number, RowInfo | ColInfo>
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
