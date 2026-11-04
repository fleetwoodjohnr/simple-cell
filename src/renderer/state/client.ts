// The renderer's side of the core worker: typed calls, and a per-workbook cache of what is on screen.
import CoreWorker from '@/core/worker?worker'
import { getView, setView, cellSel } from './view'
import type { Host, SheetMeta, WorkbookMeta } from '@/core/host'
import type { ViewCell } from '@/core/doc'
import type { Area } from '@/core/engine/values'
import type { FormatWarning, FormatId } from '@/core/formats/types'

type Methods = { [K in keyof Host]: Host[K] extends (...a: infer A) => infer R ? (...a: A) => Promise<Awaited<R>> : never }

let worker: Worker | null = null
let seq = 0
const pending = new Map<number, { doc: string; resolve: (v: unknown) => void; reject: (e: Error) => void }>()
/** Calls still on their way per workbook, and who waits for them to finish (see Book.idle). */
const inflight = new Map<string, number>()
const idleWaiters = new Map<string, (() => void)[]>()

function settled(doc: string) {
  const n = (inflight.get(doc) ?? 1) - 1
  if (n > 0) return void inflight.set(doc, n)
  inflight.delete(doc)
  // After the callers' own continuations (queued when their promise resolved), so their markDirty has run.
  for (const w of idleWaiters.get(doc) ?? []) w()
  idleWaiters.delete(doc)
}
const books = new Map<string, Book>()

function ensureWorker(): Worker {
  if (worker) return worker
  worker = new CoreWorker()
  worker.onmessage = (e: MessageEvent) => {
    const m = e.data
    if (m.event === 'changes') {
      books.get(m.doc)?.applyChanges(m.changes, m.meta)
      return
    }
    const p = pending.get(m.id)
    if (!p) return
    pending.delete(m.id)
    if ('error' in m) p.reject(new Error(m.error))
    else p.resolve(m.result)
    settled(p.doc)
  }
  worker.onerror = (e) => console.error('Core worker error', e)
  return worker
}

function call<T>(doc: string, method: string, args: unknown[], transfer: Transferable[] = []): Promise<T> {
  const w = ensureWorker()
  const id = ++seq
  return new Promise<T>((resolve, reject) => {
    pending.set(id, { doc, resolve: resolve as (v: unknown) => void, reject })
    inflight.set(doc, (inflight.get(doc) ?? 0) + 1)
    w.postMessage({ id, doc, method, args }, transfer)
  })
}

const VIEW_MARGIN = 40 // extra rows/columns fetched around what is visible

/** One open workbook as the UI sees it. */
export class Book {
  meta: WorkbookMeta
  /** Cached cells per sheet: key r * 16384 + c. */
  cells = new Map<number, Map<number, ViewCell>>()
  /** Area already fetched per sheet. */
  private fetched = new Map<number, { r0: number; r1: number; c0: number; c1: number }[]>()
  private listeners = new Set<() => void>()
  /** Bumped whenever the cache or meta changes; the grid redraws on change. */
  version = 0
  readonly api: Methods

  constructor(
    readonly id: string,
    meta: WorkbookMeta
  ) {
    this.meta = meta
    books.set(id, this)
    this.api = new Proxy({} as Methods, { get: (_t, method: string) => (...args: unknown[]) => call(id, method, args) })
  }

  /** Resolves once every call sent for this workbook has been answered (and its caller has carried on). */
  idle(): Promise<void> {
    if (!inflight.get(this.id)) return Promise.resolve()
    return new Promise((resolve) => {
      idleWaiters.set(this.id, [...(idleWaiters.get(this.id) ?? []), () => queueMicrotask(resolve)])
      // A worker that died never answers: don't hold up closing for ever.
      setTimeout(resolve, 3000)
    })
  }

  static async create(id: string): Promise<Book> {
    const meta = await call<WorkbookMeta>(id, 'create', [])
    return new Book(id, meta)
  }

  static async open(id: string, bytes: Uint8Array, format: FormatId): Promise<{ book: Book; warnings: FormatWarning[] }> {
    const copy = bytes.slice()
    const r = await call<{ meta: WorkbookMeta; warnings: FormatWarning[] }>(id, 'open', [copy, format], [copy.buffer])
    return { book: new Book(id, r.meta), warnings: r.warnings }
  }

  static async restore(id: string, snapshot: unknown): Promise<Book> {
    const meta = await call<WorkbookMeta>(id, 'restore', [snapshot])
    return new Book(id, meta)
  }

  close() {
    books.delete(this.id)
    void call(this.id, 'close', [])
  }

  subscribe(fn: () => void) {
    this.listeners.add(fn)
    return () => {
      this.listeners.delete(fn)
    }
  }

  private emit() {
    this.version++
    for (const l of this.listeners) l()
  }

  sheetMeta(id: number): SheetMeta | undefined {
    return this.meta.sheets.find((s) => s.id === id)
  }

  cell(sheet: number, r: number, c: number): ViewCell | undefined {
    return this.cells.get(sheet)?.get(r * 16384 + c)
  }

  /**
   * Make sure the visible windows (the scrolling pane plus any frozen panes) are cached, with a margin.
   * Cells outside every window are dropped to bound memory.
   */
  async ensure(sheet: number, wins: { r0: number; r1: number; c0: number; c1: number }[]) {
    const have = this.fetched.get(sheet) ?? []
    const covered = (w: (typeof wins)[number]) => have.some((f) => w.r0 >= f.r0 && w.r1 <= f.r1 && w.c0 >= f.c0 && w.c1 <= f.c1)
    if (wins.every(covered)) return
    const want = wins.map((w) => ({ r0: Math.max(0, w.r0 - VIEW_MARGIN), r1: w.r1 + VIEW_MARGIN * 2, c0: Math.max(0, w.c0 - 10), c1: w.c1 + 20 }))
    this.fetched.set(sheet, want)
    const results = await Promise.all(want.map((w) => call<ViewCell[]>(this.id, 'viewport', [sheet, w.r0, w.r1, w.c0, w.c1])))
    if (this.fetched.get(sheet) !== want) return // a newer request replaced this one
    let map = this.cells.get(sheet)
    if (!map) this.cells.set(sheet, (map = new Map()))
    // The fetched windows replace what was cached there (cells emptied since are dropped too); cells outside every
    // window are dropped to bound memory.
    map.clear()
    for (const cells of results) for (const v of cells) map.set(v.r * 16384 + v.c, v)
    this.emit()
  }

  /** Forget cached cells (after structural changes the positions moved). */
  invalidate(sheet?: number) {
    if (sheet === undefined) {
      this.cells.clear()
      this.fetched.clear()
    } else {
      this.cells.delete(sheet)
      this.fetched.delete(sheet)
    }
  }

  applyChanges(changes: { sheet: number; cells: ViewCell[]; cleared: [number, number][] }[], meta: WorkbookMeta) {
    const structural = meta.structureVersion !== this.meta.structureVersion
    if (meta.viewVersion !== this.meta.viewVersion) this.fetched.clear()
    const widths = (m: WorkbookMeta) => new Map(m.sheets.map((s) => [s.id, JSON.stringify(s.cols)]))
    const before = widths(this.meta)
    this.meta = meta
    // Undo can remove the active result sheet. Select a surviving sheet before subscribers render.
    const view = getView(this.id)
    if (view && !meta.sheets.some((s) => s.id === view.sheet)) {
      const first = meta.sheets.find((s) => !s.hidden) ?? meta.sheets[0]
      if (first) setView(this.id, { sheet: first.id, editing: null, sel: cellSel(0, 0) })
    }
    // Column widths change how General numbers are shown: refetch those sheets.
    for (const [id, w] of widths(meta)) if (before.get(id) !== w) this.invalidate(id)
    if (structural) this.invalidate()
    else
      for (const ch of changes) {
        const map = this.cells.get(ch.sheet)
        if (!map) continue
        for (const v of ch.cells) map.set(v.r * 16384 + v.c, v)
        for (const [r, c] of ch.cleared) map.delete(r * 16384 + c)
        // Conditional formats depend on other cells: refetch the visible windows (patched cells show meanwhile).
        const sm = meta.sheets.find((s) => s.id === ch.sheet)
        if (sm?.cf || sm?.sparklines) this.fetched.delete(ch.sheet)
      }
    this.emit()
  }

  async refreshMeta() {
    this.meta = await this.api.meta()
    this.emit()
  }
}

export function getBook(id: string) {
  return books.get(id)
}

export type { Area }
