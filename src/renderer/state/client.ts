// The renderer's side of the core worker: typed calls, and a per-workbook cache of what is on screen.
import CoreWorker from '@/core/worker?worker'
import type { Host, SheetMeta, WorkbookMeta } from '@/core/host'
import type { ViewCell } from '@/core/doc'
import type { Area } from '@/core/engine/values'
import type { FormatWarning, FormatId } from '@/core/formats/types'

type Methods = { [K in keyof Host]: Host[K] extends (...a: infer A) => infer R ? (...a: A) => Promise<Awaited<R>> : never }

let worker: Worker | null = null
let seq = 0
const pending = new Map<number, { resolve: (v: unknown) => void; reject: (e: Error) => void }>()
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
  }
  worker.onerror = (e) => console.error('Core worker error', e)
  return worker
}

function call<T>(doc: string, method: string, args: unknown[], transfer: Transferable[] = []): Promise<T> {
  const w = ensureWorker()
  const id = ++seq
  return new Promise<T>((resolve, reject) => {
    pending.set(id, { resolve: resolve as (v: unknown) => void, reject })
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
    for (const k of map.keys()) {
      const r = Math.floor(k / 16384)
      const c = k % 16384
      if (!want.some((w) => r >= w.r0 && r <= w.r1 && c >= w.c0 && c <= w.c1)) map.delete(k)
    }
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
    const widths = (m: WorkbookMeta) => new Map(m.sheets.map((s) => [s.id, JSON.stringify(s.cols)]))
    const before = widths(this.meta)
    this.meta = meta
    // Column widths change how General numbers are shown: refetch those sheets.
    for (const [id, w] of widths(meta)) if (before.get(id) !== w) this.invalidate(id)
    if (structural) this.invalidate()
    else
      for (const ch of changes) {
        const map = this.cells.get(ch.sheet)
        if (!map) continue
        for (const v of ch.cells) map.set(v.r * 16384 + v.c, v)
        for (const [r, c] of ch.cleared) map.delete(r * 16384 + c)
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
