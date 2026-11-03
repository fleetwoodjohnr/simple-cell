// The core runs here, off the UI thread. Messages: { id, doc, method, args } → { id, result } | { id, error }.
// After every call that may change a workbook, the worker sends { event: 'changes', doc, ... } with what to redraw.
import { Host } from './host'
import type { FormatId } from './formats/types'

const docs = new Map<string, Host>()

type Msg = { id: number; doc: string; method: string; args: unknown[] }

const MUTATING = new Set(['input', 'inputAll', 'pasteText', 'pasteCells', 'clear', 'formatCells', 'toggle', 'insert', 'remove', 'undo', 'redo', 'addSheet', 'deleteSheet', 'renameSheet', 'moveSheet', 'sheetProps', 'setColWidths', 'setRowHeights', 'autofitCols', 'hide', 'freeze', 'merge', 'unmerge', 'defineName', 'deleteName', 'setValidation', 'setNote', 'insertCells', 'deleteCells', 'setLink', 'createTable', 'setTableOptions', 'convertTableToRange', 'toggleAutoFilter', 'setFilter', 'clearFilters', 'reapplyFilter', 'addCondFormat', 'clearCondFormats', 'deleteCondRule', 'moveCondRule', 'fill', 'autofill', 'sort', 'replace', 'setCalc', 'calculate'])

async function handle(m: Msg): Promise<unknown> {
  if (m.method === 'create') {
    docs.set(m.doc, new Host())
    return docs.get(m.doc)!.meta()
  }
  if (m.method === 'open') {
    const [bytes, format] = m.args as [Uint8Array, FormatId]
    const h = await Host.open(bytes, format)
    docs.set(m.doc, h)
    return { meta: h.meta(), warnings: h.warnings }
  }
  if (m.method === 'restore') {
    const h = Host.fromSnapshot(m.args[0] as never)
    docs.set(m.doc, h)
    return h.meta()
  }
  if (m.method === 'close') {
    docs.delete(m.doc)
    return null
  }
  const h = docs.get(m.doc)
  if (!h) throw new Error('This workbook is not open')
  const fn = (h as unknown as Record<string, (...a: unknown[]) => unknown>)[m.method]
  if (typeof fn !== 'function') throw new Error(`Unknown method ${m.method}`)
  const result = await fn.apply(h, m.args)
  if (MUTATING.has(m.method)) postMessage({ event: 'changes', doc: m.doc, changes: h.changes(), meta: h.meta() })
  return result
}

self.onmessage = async (e: MessageEvent<Msg>) => {
  const m = e.data
  try {
    const result = await handle(m)
    postMessage({ id: m.id, result })
  } catch (err) {
    postMessage({ id: m.id, error: (err as Error).message ?? String(err) })
  }
}
