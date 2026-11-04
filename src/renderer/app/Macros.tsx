// Macros: JavaScript run in a sandboxed Web Worker with a workbook-only API, their editor, trust, and the VBA viewer.
//
// Sandbox: the worker has no DOM, no Node and no access to the app's IPC. The page's Content-Security-Policy
// (connect-src 'self') blocks network requests from it, and the prelude also removes fetch, XMLHttpRequest,
// WebSocket, EventSource, importScripts and storage APIs. A macro can only read and change the open workbook,
// through the messages below, and is stopped after 30 seconds.
import { useEffect, useRef, useState } from 'react'
import { EditorView, basicSetup } from 'codemirror'
import { javascript } from '@codemirror/lang-javascript'
import { HighlightStyle, syntaxHighlighting } from '@codemirror/language'
import { tags as t } from '@lezer/highlight'
import { Play, Plus, Square, Trash2 } from 'lucide-react'
import type { Area } from '@/core/engine/values'
import { areaName, parseArea } from '@/core/refs'
import type { Book } from '../state/client'
import { get, set, updateSettings } from '../state/store'
import { getView, setView } from '../state/view'
import { Btn, Modal, cx, inputCls } from '../ui'
import * as A from './actions'

export const MACRO_TIMEOUT_MS = 30_000

// Syntax colours with at least 4.5:1 contrast in both themes (CodeMirror's default style is for light backgrounds only).
const highlight = HighlightStyle.define([
  { tag: t.keyword, color: 'light-dark(#7c3aed, #c4b5fd)' },
  { tag: [t.string, t.regexp], color: 'light-dark(#b42318, #fca5a5)' },
  { tag: [t.number, t.bool, t.null], color: 'light-dark(#15803d, #86efac)' },
  { tag: [t.definition(t.variableName), t.function(t.variableName), t.function(t.propertyName)], color: 'light-dark(#1d4ed8, #93c5fd)' },
  { tag: [t.propertyName, t.typeName, t.className], color: 'light-dark(#9a3412, #fdba74)' },
  // Brighter than --text-3 in dark mode: it must also pass on the active line's highlight.
  { tag: t.comment, color: 'light-dark(#696973, #9d9da8)', fontStyle: 'italic' }
])
const editorTheme = EditorView.theme({
  '&': { color: 'var(--text)', backgroundColor: 'var(--panel)' },
  '.cm-gutters': { color: 'var(--text-3)', backgroundColor: 'var(--panel-2)', borderColor: 'var(--border)' },
  '.cm-activeLine, .cm-activeLineGutter': { backgroundColor: 'var(--hover)' },
  '.cm-cursor': { borderLeftColor: 'var(--text)' }
})

const PRELUDE = `"use strict";
(() => {
  for (const k of ['fetch', 'XMLHttpRequest', 'WebSocket', 'EventSource', 'importScripts', 'indexedDB', 'caches', 'BroadcastChannel', 'Request', 'Response', 'Worker', 'SharedWorker', 'navigator']) {
    try { Object.defineProperty(self, k, { value: undefined, writable: false, configurable: false }) } catch (e) {}
  }
})();
const __pending = new Map(); let __seq = 0;
self.onmessage = (e) => { const m = e.data; const p = __pending.get(m.id); if (!p) return; __pending.delete(m.id); m.error ? p.reject(new Error(m.error)) : p.resolve(m.result) };
const __call = (op, args) => new Promise((resolve, reject) => { const id = ++__seq; __pending.set(id, { resolve, reject }); postMessage({ call: op, args: args || [], id }) });
const __fmt = (x) => { try { return typeof x === 'string' ? x : JSON.stringify(x) } catch (e) { return String(x) } };
const console = { log: (...a) => postMessage({ log: a.map(__fmt).join(' ') }), info: (...a) => postMessage({ log: a.map(__fmt).join(' ') }), warn: (...a) => postMessage({ log: '⚠ ' + a.map(__fmt).join(' ') }), error: (...a) => postMessage({ log: '✖ ' + a.map(__fmt).join(' ') }) };
class Range {
  constructor(sheet, address) { this.sheet = sheet; this.address = address }
  getValue() { return __call('getValues', [this.sheet, this.address]).then((v) => (v[0] && v[0].length ? v[0][0] : null)) }
  getValues() { return __call('getValues', [this.sheet, this.address]) }
  getFormulas() { return __call('getFormulas', [this.sheet, this.address]) }
  setValue(v) { return __call('setValues', [this.sheet, this.address, [[v]]]) }
  setValues(v) { return __call('setValues', [this.sheet, this.address, v]) }
  setFormula(f) { return __call('setValues', [this.sheet, this.address, [[f]]]) }
  setFormulas(f) { return __call('setValues', [this.sheet, this.address, f]) }
  format(style) { return __call('format', [this.sheet, this.address, style]) }
  clear() { return __call('clear', [this.sheet, this.address]) }
  select() { return __call('select', [this.sheet, this.address]) }
}
class Sheet {
  constructor(name) { this.name = name }
  range(address) { return new Range(this.name, address) }
  getUsedRange() { return __call('usedRange', [this.name]).then((a) => new Range(this.name, a)) }
}
const workbook = {
  get activeSheet() { return new Sheet(null) },
  sheet(name) { return new Sheet(name) },
  sheets() { return __call('sheets') },
  addSheet(name) { return __call('addSheet', [name]).then((n) => new Sheet(n)) },
  getSelection() { return __call('getSelection').then((s) => new Range(s.sheet, s.address)) },
  calculate() { return __call('calculate') }
};
`

export interface RunHandle {
  done: Promise<{ ok: boolean; error?: string }>
  stop: () => void
}

/** Run macro code against the active workbook. `log` receives console output. */
export function runMacro(book: Book, tab: string, code: string, log: (line: string) => void): RunHandle {
  const src = `${PRELUDE}\n;(async () => {\n${code}\n;if (typeof main === 'function') await main(workbook)\n})().then(() => postMessage({ done: true }), (e) => postMessage({ error: (e && e.message) || String(e) }))`
  const url = URL.createObjectURL(new Blob([src], { type: 'text/javascript' }))
  const worker = new Worker(url, { type: 'classic' })
  let finish: (r: { ok: boolean; error?: string }) => void = () => {}
  const done = new Promise<{ ok: boolean; error?: string }>((r) => (finish = r))
  const end = (r: { ok: boolean; error?: string }) => {
    clearTimeout(timer)
    worker.terminate()
    URL.revokeObjectURL(url)
    finish(r)
  }
  const timer = setTimeout(() => end({ ok: false, error: `The macro ran longer than ${MACRO_TIMEOUT_MS / 1000} seconds and was stopped.` }), MACRO_TIMEOUT_MS)
  const sheetId = (name: unknown): number => {
    if (name === null || name === undefined) return getView(tab).sheet
    const s = book.meta.sheets.find((x) => x.name.toLowerCase() === String(name).toLowerCase())
    if (!s) throw new Error(`There is no sheet called “${String(name)}”.`)
    return s.id
  }
  const area = (sheet: number, ref: unknown): Area => {
    const a = typeof ref === 'string' ? parseArea(ref.replace(/\$/g, ''), sheet) : null
    if (!a) throw new Error(`“${String(ref)}” isn’t a range address such as A1 or B2:D10.`)
    return a
  }
  const ops: Record<string, (...a: unknown[]) => Promise<unknown>> = {
    sheets: async () => book.meta.sheets.map((s) => s.name),
    getValues: async (s, r) => {
      const id = sheetId(s)
      return book.api.rangeValues(id, area(id, r))
    },
    getFormulas: async (s, r) => {
      const id = sheetId(s)
      return book.api.rangeValues(id, area(id, r), true)
    },
    setValues: async (s, r, v) => {
      const id = sheetId(s)
      if (!Array.isArray(v) || !v.every(Array.isArray)) throw new Error('setValues needs rows of values: [[1, 2], [3, 4]].')
      const res = await book.api.setRangeValues(id, area(id, r), v as unknown[][])
      if (!res.ok) throw new Error(res.error)
      A.markDirtyTab(tab)
    },
    format: async (s, r, style) => {
      const id = sheetId(s)
      const allowed = ['bold', 'italic', 'underline', 'strike', 'color', 'fill', 'numFmt', 'size', 'font', 'hAlign', 'wrap']
      const patch = Object.fromEntries(Object.entries((style ?? {}) as Record<string, unknown>).filter(([k]) => allowed.includes(k)))
      const res = await book.api.formatCells([area(id, r)], patch)
      if (!res.ok) throw new Error(res.error)
      A.markDirtyTab(tab)
    },
    clear: async (s, r) => {
      const id = sheetId(s)
      await book.api.clear([area(id, r)], 'contents')
      A.markDirtyTab(tab)
    },
    select: async (s, r) => {
      const id = sheetId(s)
      const a = area(id, r)
      setView(tab, { sheet: id, sel: { ranges: [{ r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }], active: { r: a.r0, c: a.c0 }, focus: { r: a.r1, c: a.c1 } } })
    },
    getSelection: async () => {
      const v = getView(tab)
      const r = v.sel.ranges[v.sel.ranges.length - 1]
      return { sheet: book.sheetMeta(v.sheet)?.name, address: areaName(r) }
    },
    usedRange: async (s) => {
      const m = book.sheetMeta(sheetId(s))!
      return areaName({ r0: 0, c0: 0, r1: Math.max(0, m.used.r), c1: Math.max(0, m.used.c) })
    },
    addSheet: async (name) => {
      const res = await book.api.addSheet(typeof name === 'string' ? name : undefined)
      if (!res.ok) throw new Error(res.error)
      await book.refreshMeta()
      A.markDirtyTab(tab)
      return book.meta.sheets[book.meta.sheets.length - 1]?.name
    },
    calculate: async () => {
      await book.api.calculate(false)
    }
  }
  worker.onmessage = async (e) => {
    const m = e.data as { call?: string; args?: unknown[]; id?: number; log?: string; done?: boolean; error?: string }
    if (m.log !== undefined) return log(String(m.log).slice(0, 2000))
    if (m.done) return end({ ok: true })
    if (m.error !== undefined && m.call === undefined) return end({ ok: false, error: m.error })
    if (!m.call || typeof m.id !== 'number') return
    const op = ops[m.call]
    try {
      if (!op) throw new Error(`Unknown operation “${m.call}”.`)
      const result = await op(...(Array.isArray(m.args) ? m.args : []))
      worker.postMessage({ id: m.id, result })
    } catch (err) {
      worker.postMessage({ id: m.id, error: (err as Error).message })
    }
  }
  worker.onerror = (e) => {
    e.preventDefault()
    // Line numbers count from the user's code (the prelude and wrapper come first).
    const offset = PRELUDE.split('\n').length + 1
    end({ ok: false, error: `${e.message}${e.lineno ? ` (line ${Math.max(1, e.lineno - offset)})` : ''}` })
  }
  return { done, stop: () => end({ ok: false, error: 'Stopped.' }) }
}

async function hashOf(list: { name: string; code: string }[]): Promise<string> {
  const data = new TextEncoder().encode(JSON.stringify(list))
  const d = await crypto.subtle.digest('SHA-256', data)
  return [...new Uint8Array(d)].map((b) => b.toString(16).padStart(2, '0')).join('')
}

/** Macros saved in a file run only after the user trusts that exact code (once per version of the macros). */
async function trusted(list: { name: string; code: string }[]): Promise<boolean> {
  const h = await hashOf(list)
  return (get().settings.trustedMacros ?? []).includes(h)
}

async function trust(list: { name: string; code: string }[]) {
  const h = await hashOf(list)
  const cur = get().settings.trustedMacros ?? []
  if (!cur.includes(h)) await updateSettings({ trustedMacros: [...cur, h].slice(-200) })
}

const EXAMPLE = `// Macros use the workbook API and run in a sandbox: they can read and change this workbook,
// but can't reach the network or your files.
async function main(workbook) {
  const sheet = workbook.activeSheet
  const [[a]] = await sheet.range('A1').getValues()
  await sheet.range('B1').setValue(typeof a === 'number' ? a * 2 : 'A1 isn’t a number')
  await sheet.range('B1').format({ bold: true })
  console.log('Done')
}
`

export function MacrosDialog() {
  const c = A.ctx()
  const [list, setList] = useState<{ name: string; code: string }[] | null>(null)
  const [sel, setSel] = useState(0)
  const [log, setLog] = useState<string[]>([])
  const [running, setRunning] = useState<RunHandle | null>(null)
  const [needTrust, setNeedTrust] = useState(false)
  // The dialog's content mounts after the first render: keep the editor's container in state.
  const [hostEl, setHostEl] = useState<HTMLDivElement | null>(null)
  const view = useRef<EditorView | null>(null)
  useEffect(() => {
    if (!c) return
    void c.book.api.macros().then(async (m) => {
      setList(m.length ? m : [{ name: 'Macro1', code: EXAMPLE }])
      setNeedTrust(m.length > 0 && !(await trusted(m)))
    })
  }, [])
  useEffect(() => {
    if (!hostEl || !list) return
    view.current?.destroy()
    view.current = new EditorView({
      doc: list[sel]?.code ?? '',
      extensions: [basicSetup, javascript(), syntaxHighlighting(highlight), editorTheme, EditorView.contentAttributes.of({ 'aria-label': 'Macro code' }), EditorView.updateListener.of((u) => u.docChanged && setList((l) => l && l.map((m, i) => (i === sel ? { ...m, code: u.state.doc.toString() } : m))))],
      parent: hostEl
    })
    return () => view.current?.destroy()
  }, [hostEl, sel, list === null])
  if (!c || !list) return null
  const save = async (l = list) => {
    await c.book.api.setMacros(l)
    // Edited now, before the trust store's round trip: a save started meanwhile would otherwise finish first and
    // then be marked edited again.
    A.markDirtyTab(c.tab)
    // Code the user writes or edits here is theirs: trusted.
    await trust(l)
    setNeedTrust(false)
  }
  const run = async () => {
    if (needTrust) return
    setLog([`▶ ${list[sel].name}`])
    await save()
    const h = runMacro(c.book, c.tab, list[sel].code, (line) => setLog((x) => [...x.slice(-500), line]))
    setRunning(h)
    const r = await h.done
    setRunning(null)
    setLog((x) => [...x, r.ok ? '✔ Finished' : `✖ ${r.error}`])
  }
  const close = async () => {
    running?.stop()
    if (!needTrust) await save()
    set({ dialog: null })
  }
  return (
    <Modal
      open
      onOpenChange={(o) => !o && void close()}
      title="Macros"
      width={980}
      footer={
        <>
          <span className="mr-auto text-[11.5px] text-text-3">Macros run in a sandbox: this workbook only, no network, no files, stopped after {MACRO_TIMEOUT_MS / 1000} s.</span>
          {running ? (
            <Btn onClick={() => running.stop()}>
              <Square className="size-3.5" /> Stop
            </Btn>
          ) : (
            <Btn variant="primary" disabled={needTrust} onClick={() => void run()}>
              <Play className="size-3.5" /> Run
            </Btn>
          )}
          <Btn onClick={() => void close()}>Close</Btn>
        </>
      }
    >
      {needTrust && (
        <div role="alert" className="mb-3 flex items-center gap-3 rounded-lg border border-warn/40 bg-warn/10 px-3 py-2 text-[12.5px]">
          <span className="flex-1">This workbook’s macros came from a file. They can read and change everything in this workbook (not your files or the network). Only enable them if you trust where the file came from.</span>
          <Btn size="sm" variant="primary" onClick={() => void trust(list).then(() => setNeedTrust(false))}>
            Enable macros
          </Btn>
        </div>
      )}
      <div className="flex h-[460px] gap-3">
        <div className="flex w-[170px] shrink-0 flex-col gap-1">
          {list.map((m, i) => (
            <div key={i} className={cx('flex items-center gap-1 rounded px-1', i === sel ? 'bg-accent-soft' : 'hover:bg-hover')}>
              <input aria-label="Macro name" className={cx(inputCls, 'h-7 border-0 bg-transparent px-1')} value={m.name} onFocus={() => setSel(i)} onChange={(e) => setList(list.map((x, k) => (k === i ? { ...x, name: e.target.value } : x)))} />
              <button type="button" aria-label={`Delete ${m.name}`} className="rounded p-1 text-text-3 hover:text-text" onClick={() => (setList(list.filter((_, k) => k !== i)), setSel(0))}>
                <Trash2 className="size-3.5" />
              </button>
            </div>
          ))}
          <Btn size="sm" onClick={() => (setList([...list, { name: `Macro${list.length + 1}`, code: EXAMPLE }]), setSel(list.length))}>
            <Plus className="size-3.5" /> New macro
          </Btn>
        </div>
        <div className="flex min-w-0 flex-1 flex-col gap-2">
          <div ref={setHostEl} data-testid="macro-editor" className="min-h-0 flex-1 overflow-auto rounded-lg border border-border font-mono text-[12.5px] [&_.cm-editor]:h-full" />
          <pre data-testid="macro-log" aria-live="polite" className="h-[110px] overflow-auto rounded-lg bg-hover p-2 font-mono text-[11.5px] text-text-2">
            {log.join('\n')}
          </pre>
        </div>
      </div>
    </Modal>
  )
}

export function VbaDialog() {
  const c = A.ctx()
  const [mods, setMods] = useState<{ name: string; code: string }[] | null>(null)
  const [sel, setSel] = useState(0)
  useEffect(() => {
    if (c) void c.book.api.vba().then((m) => setMods(m ?? []))
  }, [])
  if (!c || !mods) return null
  return (
    <Modal open onOpenChange={(o) => !o && set({ dialog: null })} title="VBA macros (read only)" description="Simple Cell keeps VBA in .xlsm files but never runs it." width={900} footer={<Btn variant="primary" onClick={() => set({ dialog: null })}>Close</Btn>}>
      {!mods.length ? (
        <p className="text-[12.5px] text-text-2">The VBA project couldn’t be read, or has no modules. It is still kept when you save as .xlsm.</p>
      ) : (
        <div className="flex h-[440px] gap-3">
          <div className="w-[170px] shrink-0 space-y-0.5">
            {mods.map((m, i) => (
              <button key={m.name} type="button" onClick={() => setSel(i)} className={cx('block w-full truncate rounded px-2 py-1 text-left text-[12.5px]', i === sel ? 'bg-accent-soft' : 'hover:bg-hover')}>
                {m.name}
              </button>
            ))}
          </div>
          <pre data-testid="vba-code" className="min-w-0 flex-1 overflow-auto rounded-lg border border-border p-3 font-mono text-[12px]">
            {mods[sel]?.code}
          </pre>
        </div>
      )}
    </Modal>
  )
}

/** Run a macro by name without opening the editor (palette, buttons). */
export async function runNamedMacro(name: string) {
  const c = A.ctx()
  if (!c) return
  const list = await c.book.api.macros()
  const m = list.find((x) => x.name === name)
  if (!m) return
  if (!(await trusted(list))) return set({ dialog: { kind: 'macros' } })
  const r = await runMacro(c.book, c.tab, m.code, () => {}).done
  if (!r.ok) set({ dialog: { kind: 'message', title: `Macro “${name}” stopped`, body: r.error ?? '' } })
}
