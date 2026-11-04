// Opening, saving, closing workbooks; AutoSave and crash recovery.
import { formatFromPath, FORMAT_LABELS, type FormatId, type FormatWarning } from '@/core/formats/types'
import type { OpenedFile, RecoveryEntry } from '@/shared/ipc-types'
import { Book } from './client'
import { HOME, books, get, set, toast, updateTab, type TabMeta } from './store'
import { dropView, getView, initView, setView } from './view'

const api = () => window.simpleCell

let seq = 0
const newId = () => `t${Date.now().toString(36)}${(seq++).toString(36)}`
const baseName = (p: string) => p.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, '')
const EXT: Record<FormatId, string> = { xlsx: 'xlsx', xlsm: 'xlsm', xltx: 'xltx', xltm: 'xltm', ods: 'ods', csv: 'csv', tsv: 'tsv', json: 'json', html: 'html', xml: 'xml', xls: 'xls', xlsb: 'xlsb' }
/** Formats Simple Cell can write (the save dialog offers these). XLS goes through LibreOffice. */
export const SAVE_FORMATS: FormatId[] = ['xlsx', 'xlsm', 'xltx', 'ods', 'csv', 'tsv', 'json', 'html', 'xml', 'xls']
/** Formats that keep a whole workbook (AutoSave writes only these). */
const FULL_FORMATS: FormatId[] = ['xlsx', 'xlsm', 'ods']
/** Formats converted through LibreOffice. */
const VIA_LO: FormatId[] = ['xls', 'xlsb']

// ---------- dirty tracking, recovery, autosave ----------

const timers = new Map<string, { recovery?: number; autosave?: number }>()

function scheduleBackground(id: string) {
  const t = timers.get(id) ?? {}
  timers.set(id, t)
  const s = get().settings
  window.clearTimeout(t.recovery)
  t.recovery = window.setTimeout(() => void writeRecovery(id), s.recoveryIntervalSec * 1000)
  const tab = get().tabs.find((x) => x.id === id)
  if (s.autosave && tab?.path && tab.format && !tab.autosaveBlocked && FULL_FORMATS.includes(tab.format)) {
    window.clearTimeout(t.autosave)
    t.autosave = window.setTimeout(() => void autosave(id), s.autosaveIntervalSec * 1000)
  }
}

function clearTimers(id: string) {
  const t = timers.get(id)
  if (t) {
    window.clearTimeout(t.recovery)
    window.clearTimeout(t.autosave)
  }
  timers.delete(id)
}

async function writeRecovery(id: string) {
  const book = books.get(id)
  const tab = get().tabs.find((t) => t.id === id)
  if (!book || !tab?.dirty) return
  try {
    const snap = await book.api.snapshot()
    await api().recoveryWrite(id, JSON.stringify({ title: tab.title, path: tab.path, format: tab.format, snapshot: snap }))
  } catch {
    /* recovery is best-effort; the next change retries */
  }
}

async function autosave(id: string) {
  const tab = get().tabs.find((t) => t.id === id)
  if (!tab?.dirty || !tab.path || !tab.format || tab.autosaveBlocked) return
  await save(id, { auto: true })
}

/** Called after every edit. */
export function markDirty(id: string) {
  const tab = get().tabs.find((t) => t.id === id)
  if (!tab) return
  if (!tab.dirty || tab.status !== 'unsaved') updateTab(id, { dirty: true, status: 'unsaved' })
  scheduleBackground(id)
}

// ---------- tabs ----------

function untitled(): string {
  const used = new Set(get().tabs.map((t) => t.title))
  for (let n = 1; ; n++) {
    const name = `Book${n}`
    if (!used.has(name)) return name
  }
}

function addTab(meta: Partial<TabMeta> & { id: string; title: string }, book: Book) {
  books.set(meta.id, book)
  initView(meta.id, book.meta.sheets.find((s) => !s.hidden)?.id ?? book.meta.sheets[0].id)
  const tab: TabMeta = { path: null, format: null, dirty: false, status: 'saved', warnings: [], autosaveBlocked: false, savedAt: null, ...meta }
  set((s) => ({ tabs: [...s.tabs, tab], active: tab.id, backstage: false }))
  focusGrid()
}

export function focusGrid() {
  requestAnimationFrame(() => document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus())
}

export async function newWorkbook(title?: string): Promise<string> {
  const id = newId()
  const book = await Book.create(id)
  addTab({ id, title: title ?? untitled() }, book)
  return id
}

export function activate(id: string) {
  set({ active: id, backstage: false })
  focusGrid()
}

export async function openFiles(files: OpenedFile[]) {
  for (const f of files) {
    const existing = f.path ? get().tabs.find((t) => t.path === f.path) : null
    if (existing) {
      activate(existing.id)
      continue
    }
    let format = formatFromPath(f.name)
    if (!format) {
      toast(`${f.name}: this file type can’t be opened`, 'error')
      continue
    }
    try {
      let data = f.data
      const warnings: FormatWarning[] = []
      if (VIA_LO.includes(format)) {
        if (!(await api().libreOfficeAvailable())) {
          toast(`${f.name}: opening ${format.toUpperCase()} files needs LibreOffice, which isn’t installed.`, 'error')
          continue
        }
        data = await api().libreOfficeConvert(f.data, format, 'xlsx')
        warnings.push({ level: 'info', message: `Converted from ${FORMAT_LABELS[format]} with LibreOffice. Saving writes an Excel workbook (.xlsx).` })
        format = 'xlsx'
      }
      const id = newId()
      const { book, warnings: w } = await Book.open(id, data, format)
      warnings.push(...w)
      const loss = warnings.some((x) => x.level === 'loss')
      const sourceFormat = formatFromPath(f.name)!
      // Templates open as a new, untitled workbook based on them (as in Excel).
      const template = sourceFormat === 'xltx' || sourceFormat === 'xltm'
      const keepPath = f.path && !VIA_LO.includes(sourceFormat) && !template
      addTab({ id, title: template ? untitled() : baseName(f.name), path: keepPath ? f.path : null, format: sourceFormat === 'xls' ? 'xls' : VIA_LO.includes(sourceFormat) || template ? null : format, warnings, autosaveBlocked: loss, savedAt: f.mtime }, book)
      if (f.path) void addRecent(f.path, f.name)
      if (loss) toast(`Opened ${f.name} with compatibility notes — see the ⚠ badge`, 'info')
    } catch (e) {
      toast(`Couldn’t open ${f.name}: ${(e as Error).message}`, 'error')
    }
  }
}

export async function openDialog() {
  try {
    await openFiles(await api().openDialog())
  } catch (e) {
    toast((e as Error).message, 'error')
  }
}

export async function openPath(path: string) {
  try {
    await openFiles([await api().readFile(path)])
  } catch (e) {
    toast(`Couldn’t open ${baseName(path)}: ${(e as Error).message}`, 'error')
  }
}

export async function openDroppedFiles(files: File[]) {
  const opened: OpenedFile[] = []
  for (const file of files) {
    const path = api().grantDroppedFile(file)
    opened.push({ path: path ?? '', name: file.name, data: new Uint8Array(await file.arrayBuffer()), mtime: file.lastModified })
  }
  await openFiles(opened)
}

async function addRecent(path: string, name: string) {
  try {
    set({ recents: await api().recentAdd({ path, name, openedAt: Date.now() }) })
  } catch {
    /* non-critical */
  }
}

// ---------- save ----------

function confirmWarnings(tabId: string, format: FormatId, warnings: FormatWarning[]): Promise<'save' | 'xlsx' | 'cancel'> {
  return new Promise((resolve) => set({ dialog: { kind: 'export-warnings', tabId, format, warnings, resolve } }))
}

export async function save(id = get().active, opts: { as?: boolean; format?: FormatId; auto?: boolean } = {}): Promise<boolean> {
  const tab = get().tabs.find((t) => t.id === id)
  const book = books.get(id)
  if (!tab || !book) return false
  let path = tab.path
  let format: FormatId = opts.format ?? (tab.format && SAVE_FORMATS.includes(tab.format) ? tab.format : get().settings.defaultSaveFormat)
  if (!SAVE_FORMATS.includes(format)) format = SAVE_FORMATS[0]
  const asNew = opts.as || !path || opts.format != null || !tab.format || !SAVE_FORMATS.includes(tab.format)
  if (asNew) {
    if (opts.auto) return false
    const others = SAVE_FORMATS.filter((f) => f !== format)
    const chosen = await api().saveDialog(`${tab.title}.${EXT[format]}`, [EXT[format], ...others.map((f) => EXT[f])])
    if (!chosen) return false
    path = chosen
    format = formatFromPath(chosen) ?? format
    if (!SAVE_FORMATS.includes(format)) {
      toast(`Saving as ${format.toUpperCase()} isn’t supported yet.`, 'error')
      return false
    }
  }
  try {
    updateTab(id, { status: 'saving' })
    const view = getView(id)
    const saveOpts = { sheet: view?.sheet, guard: get().settings.csvFormulaGuard, defaultFont: { name: get().settings.defaultFont, size: get().settings.defaultFontSize } }
    // XLS: write XLSX, then LibreOffice converts it (Simple Cell has no native BIFF8 writer).
    const viaLo = format === 'xls'
    if (viaLo && !(await api().libreOfficeAvailable())) throw new Error('Saving as .xls needs LibreOffice, which isn’t installed. Save as .xlsx instead.')
    const result = await book.api.save(viaLo ? 'xlsx' : format, saveOpts)
    if (viaLo) {
      result.data = await api().libreOfficeConvert(result.data, 'xlsx', 'xls')
      result.warnings.push({ level: 'loss', message: 'Excel 97-2003 (.xls) is written by LibreOffice from an .xlsx copy: newer features (dynamic arrays, more than 65,536 rows, newer functions) may be lost.' })
    }
    const losses = result.warnings.filter((w) => w.level === 'loss')
    if (losses.length && !opts.auto && (asNew || format !== tab.format || !FULL_FORMATS.includes(format))) {
      const choice = await confirmWarnings(id, format, losses)
      if (choice !== 'save') {
        updateTab(id, { status: tab.dirty ? 'unsaved' : 'saved' })
        return choice === 'xlsx' ? save(id, { as: true, format: 'xlsx' }) : false
      }
    } else if (losses.length && opts.auto && !FULL_FORMATS.includes(format)) {
      updateTab(id, { status: 'unsaved', autosaveBlocked: true })
      return false
    }
    await api().writeFile(path!, result.data)
    const title = baseName(path!)
    updateTab(id, { path, format, title, dirty: false, status: 'saved', savedAt: Date.now(), autosaveBlocked: false, error: undefined, warnings: format === tab.format ? tab.warnings : [] })
    clearTimers(id)
    void api().recoveryRemove(id)
    if (!opts.auto) await addRecent(path!, `${title}.${EXT[format]}`)
    return true
  } catch (e) {
    updateTab(id, { status: 'error', error: (e as Error).message })
    toast(`Save failed: ${(e as Error).message}`, 'error')
    return false
  }
}

// ---------- closing ----------

export async function closeTab(id: string, force = false) {
  // An edit still on its way to the core marks the workbook unsaved when it lands: wait for it, so closing right
  // after typing still asks to save instead of discarding the edit.
  if (!force) await books.get(id)?.idle()
  const tab = get().tabs.find((t) => t.id === id)
  if (!tab) return
  if (tab.dirty && !force) {
    set({ dialog: { kind: 'confirm-close', tabIds: [id] } })
    return
  }
  clearTimers(id)
  void api().recoveryRemove(id)
  const tabs = get().tabs
  const idx = tabs.findIndex((t) => t.id === id)
  const rest = tabs.filter((t) => t.id !== id)
  const next = get().active === id ? ((rest[idx] ?? rest[idx - 1])?.id ?? HOME) : get().active
  set({ tabs: rest, active: next })
  requestAnimationFrame(() => {
    books.get(id)?.close()
    books.delete(id)
    dropView(id)
  })
}

export async function requestQuit() {
  await Promise.all(get().tabs.map((t) => books.get(t.id)?.idle()))
  const dirty = get().tabs.filter((t) => t.dirty).map((t) => t.id)
  if (dirty.length) set({ dialog: { kind: 'confirm-close', tabIds: dirty, then: 'quit' } })
  else quitNow()
}

export function quitNow() {
  for (const t of get().tabs) void api().recoveryRemove(t.id)
  api().confirmClose()
}

// ---------- recovery ----------

export async function restoreRecovery(entry: RecoveryEntry) {
  try {
    const data = JSON.parse(entry.payload) as { title: string; path: string | null; format: FormatId | null; snapshot: unknown }
    const id = newId()
    const book = await Book.restore(id, data.snapshot)
    addTab({ id, title: `${data.title} (recovered)`, path: data.path, format: data.format, dirty: true, status: 'unsaved', autosaveBlocked: true }, book)
    void writeRecovery(id)
  } catch {
    toast('That recovery file is damaged and could not be restored.', 'error')
  }
  void discardRecovery(entry.id)
}

export async function discardRecovery(id: string) {
  await api().recoveryRemove(id)
  set((s) => ({ recovery: s.recovery.filter((r) => r.id !== id) }))
}

export function switchSheet(tab: string, sheet: number) {
  setView(tab, (v) => (v.sheet === sheet ? {} : { sheet, editing: null, sel: { ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], active: { r: 0, c: 0 }, focus: { r: 0, c: 0 } } }))
}
