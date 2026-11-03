import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { AlertTriangle, Info, Monitor, Moon, Sun, Search } from 'lucide-react'
import { useApp, set, get, books, updateSettings, toast, type Dialog } from '../state/store'
import * as docs from '../state/documents'
import { allCommands, eventToShortcut, formatShortcut, getCommand, runCommand, shortcutFor, shortcutsFor } from './commands'
import { BUILTIN, RESERVED } from './shortcuts'
import { Btn, Field, Kbd, Modal, Segmented, Switch, cx, inputCls } from '../ui'
import { FORMAT_LABELS } from '@/core/formats/types'
import type { Settings } from '@/shared/ipc-types'
import * as A from './actions'
import { getView, setView, areasOf } from '../state/view'
import { CATALOG } from '@/core/engine/catalog'
import { formatValue } from '@/core/numfmt/format'
import type { Style, BorderStyle } from '@/core/model/styles'
import { colName, parseCell } from '@/core/refs'
import { guessHeader, currentRegion } from './commands'

const close = () => set({ dialog: null })

export function Dialogs() {
  const d = useApp((s) => s.dialog)
  if (!d) return null
  switch (d.kind) {
    case 'settings':
      return <SettingsDialog initial={d.section} />
    case 'confirm-close':
      return <ConfirmClose d={d} />
    case 'compat':
      return <CompatDialog tabId={d.tabId} />
    case 'shortcuts':
      return <ShortcutsDialog />
    case 'export-warnings':
      return <ExportWarnings d={d} />
    case 'palette':
      return <CommandPalette />
    case 'goto':
      return <GotoDialog />
    case 'find':
      return <FindDialog replace={d.replace} />
    case 'names':
      return <NamesDialog />
    case 'insert-function':
      return <InsertFunctionDialog />
    case 'format-cells':
      return <FormatCellsDialog initial={d.tab} />
    case 'sort':
      return <SortDialog />
    case 'row-height':
    case 'col-width':
      return <SizeDialog axis={d.kind === 'row-height' ? 'row' : 'col'} />
    case 'message':
      return (
        <Modal open onOpenChange={(o) => !o && close()} title={d.title} width={460} footer={<Btn variant="primary" autoFocus onClick={close}>OK</Btn>}>
          <p className="whitespace-pre-line text-[13px] leading-relaxed text-text-2">{d.body}</p>
        </Modal>
      )
    case 'about':
      return <AboutDialog />
    default:
      return null
  }
}

function afterDialog() {
  requestAnimationFrame(() => document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus())
}

function ConfirmClose({ d }: { d: Extract<Dialog, { kind: 'confirm-close' }> }) {
  const tabs = useApp((s) => s.tabs.filter((t) => d.tabIds.includes(t.id)))
  const many = tabs.length > 1
  const finish = () => {
    close()
    if (d.then === 'quit') docs.quitNow()
  }
  return (
    <Modal
      open
      onOpenChange={(o) => !o && close()}
      title={many ? `Save changes to ${tabs.length} workbooks?` : `Save changes to “${tabs[0]?.title}”?`}
      description="Your changes will be lost if you don’t save them."
      width={440}
      footer={
        <>
          <Btn
            variant="ghost"
            onClick={() => {
              tabs.forEach((t) => docs.closeTab(t.id, true))
              finish()
            }}
          >
            Don’t save
          </Btn>
          <Btn onClick={close}>Cancel</Btn>
          <Btn
            variant="primary"
            autoFocus
            onClick={async () => {
              for (const t of tabs) {
                docs.activate(t.id)
                if (!(await docs.save(t.id))) return
                docs.closeTab(t.id, true)
              }
              finish()
            }}
          >
            Save
          </Btn>
        </>
      }
    />
  )
}

function CompatDialog({ tabId }: { tabId: string }) {
  const tab = useApp((s) => s.tabs.find((t) => t.id === tabId))
  if (!tab) return null
  const loss = tab.warnings.filter((w) => w.level === 'loss')
  const info = tab.warnings.filter((w) => w.level === 'info')
  return (
    <Modal open onOpenChange={(o) => !o && close()} title="Compatibility report" description={tab.format ? `${tab.title} · ${FORMAT_LABELS[tab.format]}` : tab.title} width={560} footer={<Btn variant="primary" onClick={close}>Done</Btn>}>
      {!tab.warnings.length && <p className="text-[13px] text-text-2">No compatibility issues were found when this file was opened.</p>}
      {loss.length > 0 && (
        <WarningList title="Preserved or changed" icon={<AlertTriangle className="size-4 text-warn" />} items={loss}>
          Some content in this file is kept as-is but can’t be edited yet, or is shown differently. Saving back to the same format keeps preserved content.
          {tab.autosaveBlocked && ' AutoSave is paused for this workbook until you save it yourself.'}
        </WarningList>
      )}
      {info.length > 0 && <WarningList title="Notes" icon={<Info className="size-4 text-text-3" />} items={info} />}
    </Modal>
  )
}

function WarningList({ title, icon, items, children }: { title: string; icon: ReactNode; items: { message: string; count?: number }[]; children?: ReactNode }) {
  return (
    <section className="mb-4">
      <h4 className="mb-1.5 flex items-center gap-2 text-[13px] font-semibold">
        {icon}
        {title}
      </h4>
      {children && <p className="mb-2 text-[12.5px] leading-relaxed text-text-2">{children}</p>}
      <ul className="divide-y divide-border rounded-lg border border-border">
        {items.map((w) => (
          <li key={w.message} className="flex items-center justify-between gap-3 px-3 py-2 text-[12.5px]">
            <span>{w.message}</span>
            {w.count && w.count > 1 && <span className="shrink-0 rounded-full bg-hover px-2 text-[11px] text-text-3">×{w.count}</span>}
          </li>
        ))}
      </ul>
    </section>
  )
}

function ExportWarnings({ d }: { d: Extract<Dialog, { kind: 'export-warnings' }> }) {
  const done = (ok: boolean) => {
    close()
    d.resolve(ok)
  }
  return (
    <Modal
      open
      onOpenChange={(o) => !o && done(false)}
      title={`Some content can’t be saved as ${FORMAT_LABELS[d.format]}`}
      description="The file will be written, but these parts of your workbook will be simplified or left out."
      width={540}
      footer={
        <>
          <Btn onClick={() => done(false)}>Cancel</Btn>
          <Btn variant="primary" autoFocus onClick={() => done(true)}>
            Save anyway
          </Btn>
        </>
      }
    >
      <WarningList title="What changes" icon={<AlertTriangle className="size-4 text-warn" />} items={d.warnings} />
    </Modal>
  )
}

// ---------- command palette ----------

function CommandPalette() {
  const [q, setQ] = useState('')
  const [i, setI] = useState(0)
  const results = useMemo(() => {
    const words = q.toLowerCase().split(/\s+/).filter(Boolean)
    const cmds = allCommands().filter((c) => words.every((w) => `${c.label} ${c.id} ${c.keywords ?? ''}`.toLowerCase().includes(w)))
    const out: { id: string; label: string; hint: string; run: () => void }[] = cmds.slice(0, 40).map((c) => ({ id: c.id, label: c.label, hint: formatShortcut(shortcutFor(c.id)), run: () => void runCommand(c.id) }))
    // Function names: insert them into the active cell.
    if (q.length >= 2)
      for (const f of CATALOG.filter((f) => f.n.toLowerCase().startsWith(q.toLowerCase())).slice(0, 6))
        out.push({ id: 'fn:' + f.n, label: `Insert function ${f.n}`, hint: f.cat, run: () => insertFunctionText(f.n) })
    if (parseCell(q.trim()) || /^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test(q.trim())) out.unshift({ id: 'goto', label: `Go to ${q.trim().toUpperCase()}`, hint: 'Enter', run: () => withCtx((c) => A.goTo(c, q)) })
    return out
  }, [q])
  const choose = (k: number) => {
    const r = results[k]
    if (!r) return
    close()
    afterDialog()
    setTimeout(r.run, 0)
  }
  return (
    <Modal open onOpenChange={(o) => !o && (close(), afterDialog())} title="Search commands" width={560}>
      <div className="relative">
        <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-text-3" />
        <input
          autoFocus
          aria-label="Search commands"
          className={cx(inputCls, 'h-10 pl-9 text-[14px]')}
          placeholder="Type a command, function name or cell reference"
          value={q}
          onChange={(e) => {
            setQ(e.target.value)
            setI(0)
          }}
          onKeyDown={(e) => {
            if (e.key === 'ArrowDown') {
              e.preventDefault()
              setI((x) => Math.min(results.length - 1, x + 1))
            } else if (e.key === 'ArrowUp') {
              e.preventDefault()
              setI((x) => Math.max(0, x - 1))
            } else if (e.key === 'Enter') {
              e.preventDefault()
              choose(i)
            }
          }}
        />
      </div>
      <ul className="mt-2 max-h-[50vh] overflow-y-auto" role="listbox" aria-label="Commands" data-testid="palette-results">
        {results.map((r, k) => (
          <li key={r.id} role="option" aria-selected={k === i} onMouseEnter={() => setI(k)} onClick={() => choose(k)} className={cx('flex h-9 cursor-default items-center justify-between rounded-lg px-3 text-[13px]', k === i && 'bg-accent-soft text-accent-ink')}>
            <span>{r.label}</span>
            <span className="text-[11.5px] text-text-3">{r.hint}</span>
          </li>
        ))}
        {!results.length && <li className="px-3 py-6 text-center text-[13px] text-text-3">No matching commands</li>}
      </ul>
    </Modal>
  )
}

function withCtx(fn: (c: A.Ctx) => unknown) {
  const c = A.ctx()
  if (c) void fn(c)
}

function insertFunctionText(name: string) {
  const c = A.ctx()
  if (!c) return
  const v = getView(c.tab)
  if (v.editing) {
    const e = v.editing
    const ins = (e.text.startsWith('=') ? '' : '=') + name + '('
    setView(c.tab, { editing: { ...e, text: e.text.slice(0, e.caret) + ins + e.text.slice(e.caret), caret: e.caret + ins.length } })
  } else setView(c.tab, { editing: { text: `=${name}(`, where: 'cell', enterMode: false, caret: name.length + 2 } })
  afterDialog()
}

// ---------- go to ----------

function GotoDialog() {
  const c = A.ctx()
  const [ref, setRef] = useState('')
  const names = c?.book.meta.names ?? []
  const go = async () => {
    if (!c) return
    if (await A.goTo(c, ref)) {
      close()
      afterDialog()
    } else toast('Reference isn’t valid.', 'error')
  }
  return (
    <Modal open onOpenChange={(o) => !o && (close(), afterDialog())} title="Go to" width={400} footer={<><Btn onClick={close}>Cancel</Btn><Btn variant="primary" onClick={go}>OK</Btn></>}>
      {names.length > 0 && (
        <ul className="mb-3 max-h-48 overflow-y-auto rounded-lg border border-border">
          {names.map((n) => (
            <li key={n.name + n.scope}>
              <button className="w-full px-3 py-1.5 text-left text-[13px] hover:bg-hover" onClick={() => setRef(n.name)} onDoubleClick={() => c && void A.goTo(c, n.name).then(() => (close(), afterDialog()))}>
                {n.name}
              </button>
            </li>
          ))}
        </ul>
      )}
      <Field label="Reference">
        <input autoFocus className={inputCls} value={ref} onChange={(e) => setRef(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && go()} placeholder="A1, Sheet2!B5:D9 or a name" aria-label="Reference" />
      </Field>
    </Modal>
  )
}

// ---------- find & replace ----------

function FindDialog({ replace: initialReplace }: { replace: boolean }) {
  const c = A.ctx()
  const [replace, setReplace] = useState(initialReplace)
  const [text, setText] = useState('')
  const [rep, setRep] = useState('')
  const [matchCase, setMatchCase] = useState(false)
  const [whole, setWhole] = useState(false)
  const [scope, setScope] = useState<'sheet' | 'workbook'>('sheet')
  const [lookIn, setLookIn] = useState<'formulas' | 'values'>('formulas')
  const [hits, setHits] = useState<{ sheet: number; r: number; c: number; text: string }[] | null>(null)
  const idx = useRef(-1)
  if (!c) return null
  const opts = () => ({ text, sheet: scope === 'sheet' ? c.view.sheet : null, lookIn: replace ? ('formulas' as const) : lookIn, matchCase, wholeCell: whole })
  const show = (h: { sheet: number; r: number; c: number }) => {
    if (h.sheet !== getView(c.tab).sheet) setView(c.tab, { sheet: h.sheet })
    setView(c.tab, { sel: { ranges: [{ r0: h.r, c0: h.c, r1: h.r, c1: h.c }], active: { r: h.r, c: h.c }, focus: { r: h.r, c: h.c } } })
  }
  const findAll = async () => {
    const list = await c.book.api.find(opts())
    setHits(list)
    idx.current = -1
    if (!list.length) toast(`Simple Cell couldn’t find “${text}”.`, 'info')
    return list
  }
  const next = async () => {
    const list = hits ?? (await findAll())
    if (!list.length) return
    const a = getView(c.tab).sel.active
    const cur = getView(c.tab).sheet
    // Next match after the active cell (row by row)
    let k = list.findIndex((h) => (h.sheet === cur && (h.r > a.r || (h.r === a.r && h.c > a.c))) || c.book.meta.sheets.findIndex((s) => s.id === h.sheet) > c.book.meta.sheets.findIndex((s) => s.id === cur))
    if (k < 0) k = 0
    idx.current = k
    show(list[k])
  }
  const replaceOne = async () => {
    const a = getView(c.tab).sel.active
    const r = await c.book.api.replace({ ...opts(), replacement: rep, all: false, at: { sheet: getView(c.tab).sheet, ...a } })
    if (r.ok && r.count) docs.markDirty(c.tab)
    setHits(null)
    await next()
  }
  const replaceAll = async () => {
    const r = await c.book.api.replace({ ...opts(), replacement: rep, all: true })
    if (!r.ok) return toast(r.error ?? 'Replace failed', 'error')
    if (r.count) docs.markDirty(c.tab)
    setHits(null)
    toast(r.count ? `Done. Simple Cell made ${r.count} replacement${r.count === 1 ? '' : 's'}.` : `Simple Cell couldn’t find “${text}”.`, 'info')
  }
  return (
    <Modal
      open
      onOpenChange={(o) => !o && (close(), afterDialog())}
      title="Find and replace"
      width={560}
      footer={
        <>
          {replace && <Btn onClick={replaceAll} disabled={!text}>Replace all</Btn>}
          {replace && <Btn onClick={replaceOne} disabled={!text}>Replace</Btn>}
          <Btn onClick={findAll} disabled={!text}>Find all</Btn>
          <Btn variant="primary" onClick={next} disabled={!text}>
            Find next
          </Btn>
        </>
      }
    >
      <div className="mb-3">
        <Segmented value={replace ? 'replace' : 'find'} onChange={(v) => setReplace(v === 'replace')} label="Mode" options={[{ value: 'find', label: 'Find' }, { value: 'replace', label: 'Replace' }]} />
      </div>
      <div className="space-y-3">
        <Field label="Find what">
          <input autoFocus aria-label="Find what" className={inputCls} value={text} onChange={(e) => (setText(e.target.value), setHits(null))} onKeyDown={(e) => e.key === 'Enter' && next()} />
        </Field>
        {replace && (
          <Field label="Replace with">
            <input aria-label="Replace with" className={inputCls} value={rep} onChange={(e) => setRep(e.target.value)} />
          </Field>
        )}
        <div className="flex flex-wrap items-center gap-x-5 gap-y-2 text-[12.5px] text-text-2">
          <label className="flex items-center gap-2">
            <Switch checked={matchCase} onChange={setMatchCase} label="Match case" /> Match case
          </label>
          <label className="flex items-center gap-2">
            <Switch checked={whole} onChange={setWhole} label="Match entire cell contents" /> Entire cell
          </label>
          <Segmented value={scope} onChange={setScope} label="Within" options={[{ value: 'sheet', label: 'Sheet' }, { value: 'workbook', label: 'Workbook' }]} />
          {!replace && <Segmented value={lookIn} onChange={setLookIn} label="Look in" options={[{ value: 'formulas', label: 'Formulas' }, { value: 'values', label: 'Values' }]} />}
        </div>
      </div>
      {hits && hits.length > 0 && (
        <div className="mt-3 max-h-56 overflow-y-auto rounded-lg border border-border" data-testid="find-results">
          <div className="sticky top-0 bg-panel-2 px-3 py-1 text-[11.5px] text-text-3">{hits.length} cell(s) found</div>
          {hits.slice(0, 500).map((h) => (
            <button key={`${h.sheet}:${h.r}:${h.c}`} className="flex w-full gap-3 px-3 py-1 text-left text-[12.5px] hover:bg-hover" onClick={() => show(h)}>
              <span className="w-24 shrink-0 truncate text-text-3">{c.book.meta.sheets.find((s) => s.id === h.sheet)?.name}</span>
              <span className="w-16 shrink-0 font-medium">
                {colName(h.c)}
                {h.r + 1}
              </span>
              <span className="truncate">{h.text}</span>
            </button>
          ))}
        </div>
      )}
    </Modal>
  )
}

// ---------- names ----------

function NamesDialog() {
  const c = A.ctx()
  const [name, setName] = useState('')
  const [ref, setRef] = useState(() => {
    if (!c) return ''
    const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
    const sh = c.book.sheetMeta(c.view.sheet)!.name
    const abs = (rr: number, cc: number) => `$${colName(cc)}$${rr + 1}`
    return `='${sh.replace(/'/g, "''")}'!${abs(r.r0, r.c0)}${r.r0 !== r.r1 || r.c0 !== r.c1 ? ':' + abs(r.r1, r.c1) : ''}`
  })
  const [scope, setScope] = useState<number | null>(null)
  const [, force] = useState(0)
  if (!c) return null
  const names = c.book.meta.names
  const save = async () => {
    const r = await c.book.api.defineName(name, ref, scope)
    if (!r.ok) return toast(r.error ?? 'Couldn’t define the name', 'error')
    docs.markDirty(c.tab)
    await c.book.refreshMeta()
    setName('')
    force((n) => n + 1)
  }
  return (
    <Modal open onOpenChange={(o) => !o && (close(), afterDialog())} title="Name manager" width={620} footer={<Btn variant="primary" onClick={() => (close(), afterDialog())}>Close</Btn>}>
      <table className="mb-4 w-full text-[12.5px]">
        <thead>
          <tr className="text-left text-text-3">
            <th className="py-1 font-medium">Name</th>
            <th className="py-1 font-medium">Refers to</th>
            <th className="py-1 font-medium">Scope</th>
            <th />
          </tr>
        </thead>
        <tbody>
          {names.map((n) => (
            <tr key={n.name + n.scope} className="border-t border-border">
              <td className="py-1.5 font-medium">{n.name}</td>
              <td className="max-w-[260px] truncate py-1.5 font-mono text-[12px]">={n.text}</td>
              <td className="py-1.5 text-text-3">{n.scope === null ? 'Workbook' : c.book.meta.sheets.find((s) => s.id === n.scope)?.name}</td>
              <td className="py-1.5 text-right">
                <Btn size="sm" variant="ghost" onClick={() => (setName(n.name), setRef('=' + n.text), setScope(n.scope))}>
                  Edit
                </Btn>
                <Btn
                  size="sm"
                  variant="ghost"
                  onClick={async () => {
                    await c.book.api.deleteName(n.name, n.scope)
                    docs.markDirty(c.tab)
                    await c.book.refreshMeta()
                    force((x) => x + 1)
                  }}
                >
                  Delete
                </Btn>
              </td>
            </tr>
          ))}
          {!names.length && (
            <tr>
              <td colSpan={4} className="py-3 text-text-3">
                No names yet.
              </td>
            </tr>
          )}
        </tbody>
      </table>
      <div className="grid grid-cols-[1fr_2fr_auto] items-end gap-2">
        <Field label="Name">
          <input className={inputCls} value={name} onChange={(e) => setName(e.target.value)} aria-label="Name" />
        </Field>
        <Field label="Refers to">
          <input className={cx(inputCls, 'font-mono')} value={ref} onChange={(e) => setRef(e.target.value)} aria-label="Refers to" />
        </Field>
        <Btn variant="primary" onClick={save} disabled={!name}>
          Save name
        </Btn>
      </div>
      <div className="mt-2 flex items-center gap-2 text-[12.5px] text-text-2">
        Scope
        <select className={cx(inputCls, 'h-7 w-48')} value={scope ?? ''} onChange={(e) => setScope(e.target.value === '' ? null : Number(e.target.value))} aria-label="Scope">
          <option value="">Workbook</option>
          {c.book.meta.sheets.map((s) => (
            <option key={s.id} value={s.id}>
              {s.name}
            </option>
          ))}
        </select>
      </div>
    </Modal>
  )
}

// ---------- insert function ----------

function InsertFunctionDialog() {
  const [q, setQ] = useState('')
  const [cat, setCat] = useState('All')
  const [sel, setSel] = useState<string | null>(null)
  const [fns, setFns] = useState<{ name: string; category: string; description: string; implemented: boolean }[]>([])
  useEffect(() => {
    const c = A.ctx()
    if (c) void c.book.api.functionList().then(setFns)
  }, [])
  const cats = ['All', ...[...new Set(fns.map((f) => f.category))].sort()]
  const list = fns.filter((f) => (cat === 'All' || f.category === cat) && (!q || f.name.toLowerCase().includes(q.toLowerCase()) || f.description.toLowerCase().includes(q.toLowerCase())))
  const chosen = fns.find((f) => f.name === sel)
  const ok = () => {
    if (!sel) return
    close()
    insertFunctionText(sel)
  }
  return (
    <Modal open onOpenChange={(o) => !o && (close(), afterDialog())} title="Insert function" width={640} footer={<><Btn onClick={close}>Cancel</Btn><Btn variant="primary" disabled={!sel} onClick={ok}>Insert</Btn></>}>
      <div className="flex gap-2">
        <input autoFocus aria-label="Search for a function" className={inputCls} placeholder="Search for a function" value={q} onChange={(e) => setQ(e.target.value)} />
        <select aria-label="Category" className={cx(inputCls, 'w-56')} value={cat} onChange={(e) => setCat(e.target.value)}>
          {cats.map((c) => (
            <option key={c}>{c}</option>
          ))}
        </select>
      </div>
      <ul className="mt-2 h-64 overflow-y-auto rounded-lg border border-border" role="listbox" aria-label="Functions">
        {list.map((f) => (
          <li key={f.name} role="option" aria-selected={sel === f.name} onClick={() => setSel(f.name)} onDoubleClick={() => (setSel(f.name), setTimeout(ok))} className={cx('flex cursor-default items-center justify-between px-3 py-1 text-[12.5px]', sel === f.name ? 'bg-accent-soft text-accent-ink' : 'hover:bg-hover')}>
            <span className="font-medium">{f.name}</span>
            {!f.implemented && <span className="rounded bg-warn/15 px-1.5 text-[10.5px] text-warn">not yet calculated</span>}
          </li>
        ))}
      </ul>
      {chosen && (
        <p className="mt-2 text-[12.5px] text-text-2">
          <span className="font-semibold text-text">{chosen.name}</span> · {chosen.category} — {chosen.description}
          {!chosen.implemented && ' Simple Cell keeps formulas using it, and shows the result saved in the file, but can’t calculate it yet.'}
        </p>
      )}
    </Modal>
  )
}

// ---------- format cells ----------

const CATEGORIES: { id: string; label: string; codes: string[] }[] = [
  { id: 'general', label: 'General', codes: ['General'] },
  { id: 'number', label: 'Number', codes: ['0', '0.00', '#,##0', '#,##0.00', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)'] },
  { id: 'currency', label: 'Currency', codes: ['"$"#,##0.00', '"$"#,##0.00_);\\("$"#,##0.00\\)', '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', '"€"#,##0.00', '"£"#,##0.00'] },
  { id: 'accounting', label: 'Accounting', codes: ['_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'] },
  { id: 'date', label: 'Date', codes: ['m/d/yyyy', 'dddd, mmmm d, yyyy', 'yyyy-mm-dd', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'mmmm d, yyyy', 'm/d/yy h:mm'] },
  { id: 'time', label: 'Time', codes: ['h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', '[h]:mm:ss', 'mm:ss.0'] },
  { id: 'percentage', label: 'Percentage', codes: ['0%', '0.00%'] },
  { id: 'fraction', label: 'Fraction', codes: ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16'] },
  { id: 'scientific', label: 'Scientific', codes: ['0.00E+00', '##0.0E+0'] },
  { id: 'text', label: 'Text', codes: ['@'] },
  { id: 'special', label: 'Special', codes: ['00000', '00000-0000', '[<=9999999]###-####;(###) ###-####', '000-00-0000'] }
]

function FormatCellsDialog({ initial }: { initial?: string }) {
  const c = A.ctx()
  const [tab, setTab] = useState(initial ?? 'number')
  const [style, setStyle] = useState<Style | null>(null)
  const [patch, setPatch] = useState<Partial<Style>>({})
  const [sample, setSample] = useState<number | string | boolean | null>(1234.5678)
  useEffect(() => {
    if (!c) return
    void A.activeStyle(c).then(setStyle)
    const v = c.book.cell(c.view.sheet, c.view.sel.active.r, c.view.sel.active.c)
    void c.book.api.value(c.view.sheet, c.view.sel.active.r, c.view.sel.active.c).then((x) => v && x !== null && typeof x !== 'object' && setSample(x as never))
  }, [])
  if (!c || !style) return null
  const cur: Style = { ...style, ...patch }
  const up = (p: Partial<Style>) => setPatch((x) => ({ ...x, ...p }))
  const code = cur.numFmt ?? 'General'
  const preview = formatValue(sample as never, code)
  const apply = async () => {
    if (Object.keys(patch).length) await A.edit(c, (b) => b.api.formatCells(areasOf(c.view), patch))
    close()
    afterDialog()
  }
  const border = (side: 'borderTop' | 'borderBottom' | 'borderLeft' | 'borderRight', on: boolean, s: BorderStyle = 'thin') => up({ [side]: on ? { style: s } : undefined })
  return (
    <Modal open onOpenChange={(o) => !o && (close(), afterDialog())} title="Format cells" width={640} footer={<><Btn onClick={close}>Cancel</Btn><Btn variant="primary" onClick={apply}>OK</Btn></>}>
      <div className="mb-3">
        <Segmented value={tab} onChange={setTab} label="Format cells" options={['number', 'alignment', 'font', 'border', 'fill', 'protection'].map((t) => ({ value: t, label: t[0].toUpperCase() + t.slice(1) }))} />
      </div>
      {tab === 'number' && (
        <div className="grid grid-cols-[160px_1fr] gap-4">
          <ul className="rounded-lg border border-border py-1 text-[12.5px]" role="listbox" aria-label="Category">
            {CATEGORIES.map((cat) => (
              <li key={cat.id} role="option" aria-selected={cat.codes.includes(code)} onClick={() => up({ numFmt: cat.codes[0] === 'General' ? undefined : cat.codes[0] })} className={cx('cursor-default px-3 py-1', cat.codes.includes(code) ? 'bg-accent-soft text-accent-ink' : 'hover:bg-hover')}>
                {cat.label}
              </li>
            ))}
          </ul>
          <div className="space-y-3">
            <div className="rounded-lg border border-border bg-panel-2 px-3 py-2 text-[13px]" data-testid="format-preview">
              <span className="text-text-3">Sample: </span>
              <span style={{ color: preview.color }}>{preview.text}</span>
            </div>
            <Field label="Type (format code)">
              <input aria-label="Format code" className={cx(inputCls, 'font-mono')} value={code} onChange={(e) => up({ numFmt: e.target.value === 'General' || !e.target.value ? undefined : e.target.value })} />
            </Field>
            <ul className="max-h-40 overflow-y-auto rounded-lg border border-border text-[12px]">
              {(CATEGORIES.find((cat) => cat.codes.includes(code))?.codes ?? CATEGORIES.flatMap((x) => x.codes)).map((k) => (
                <li key={k} onClick={() => up({ numFmt: k === 'General' ? undefined : k })} className={cx('flex cursor-default justify-between gap-3 px-3 py-1 font-mono', k === code ? 'bg-accent-soft' : 'hover:bg-hover')}>
                  <span className="truncate">{k}</span>
                  <span className="font-sans text-text-3">{formatValue(typeof sample === 'number' ? sample : 1234.5678, k).text}</span>
                </li>
              ))}
            </ul>
          </div>
        </div>
      )}
      {tab === 'alignment' && (
        <div className="grid grid-cols-2 gap-4 text-[12.5px]">
          <Field label="Horizontal">
            <select aria-label="Horizontal" className={inputCls} value={cur.hAlign ?? 'general'} onChange={(e) => up({ hAlign: e.target.value === 'general' ? undefined : (e.target.value as Style['hAlign']) })}>
              {['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'].map((v) => (
                <option key={v} value={v}>
                  {v === 'centerContinuous' ? 'Center across selection' : v[0].toUpperCase() + v.slice(1)}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Vertical">
            <select aria-label="Vertical" className={inputCls} value={cur.vAlign ?? 'bottom'} onChange={(e) => up({ vAlign: e.target.value === 'bottom' ? undefined : (e.target.value as Style['vAlign']) })}>
              {['top', 'center', 'bottom', 'justify', 'distributed'].map((v) => (
                <option key={v} value={v}>
                  {v[0].toUpperCase() + v.slice(1)}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Indent">
            <input aria-label="Indent" type="number" min={0} max={15} className={inputCls} value={cur.indent ?? 0} onChange={(e) => up({ indent: Number(e.target.value) || undefined })} />
          </Field>
          <Field label="Orientation (degrees, -90 to 90)">
            <input aria-label="Orientation" type="number" min={-90} max={90} className={inputCls} value={cur.rotation ?? 0} onChange={(e) => up({ rotation: Number(e.target.value) || undefined })} />
          </Field>
          <label className="flex items-center gap-2">
            <Switch checked={!!cur.wrap} onChange={(v) => up({ wrap: v || undefined })} label="Wrap text" /> Wrap text
          </label>
          <label className="flex items-center gap-2">
            <Switch checked={!!cur.shrink} onChange={(v) => up({ shrink: v || undefined })} label="Shrink to fit" /> Shrink to fit
          </label>
        </div>
      )}
      {tab === 'font' && (
        <div className="grid grid-cols-2 gap-4 text-[12.5px]">
          <Field label="Font">
            <input aria-label="Font" className={inputCls} value={cur.font ?? get().settings.defaultFont} onChange={(e) => up({ font: e.target.value })} />
          </Field>
          <Field label="Size">
            <input aria-label="Size" type="number" min={1} max={409} className={inputCls} value={cur.size ?? get().settings.defaultFontSize} onChange={(e) => up({ size: Number(e.target.value) })} />
          </Field>
          <div className="col-span-2 flex flex-wrap gap-5">
            {(['bold', 'italic', 'strike'] as const).map((k) => (
              <label key={k} className="flex items-center gap-2">
                <Switch checked={!!cur[k]} onChange={(v) => up({ [k]: v || undefined })} label={k} /> {k === 'strike' ? 'Strikethrough' : k[0].toUpperCase() + k.slice(1)}
              </label>
            ))}
            <label className="flex items-center gap-2">
              Underline
              <select aria-label="Underline" className={cx(inputCls, 'h-7 w-32')} value={cur.underline ?? 'none'} onChange={(e) => up({ underline: e.target.value === 'none' ? undefined : (e.target.value as Style['underline']) })}>
                <option value="none">None</option>
                <option value="single">Single</option>
                <option value="double">Double</option>
                <option value="singleAccounting">Single accounting</option>
                <option value="doubleAccounting">Double accounting</option>
              </select>
            </label>
            <label className="flex items-center gap-2">
              Color <input aria-label="Font color" type="color" value={cur.color ?? '#000000'} onChange={(e) => up({ color: e.target.value })} />
            </label>
          </div>
        </div>
      )}
      {tab === 'border' && (
        <div className="space-y-3 text-[12.5px]">
          <div className="flex flex-wrap gap-2">
            <Btn size="sm" onClick={() => up({ borderTop: undefined, borderBottom: undefined, borderLeft: undefined, borderRight: undefined })}>
              None
            </Btn>
            <Btn size="sm" onClick={() => up({ borderTop: { style: 'thin' }, borderBottom: { style: 'thin' }, borderLeft: { style: 'thin' }, borderRight: { style: 'thin' } })}>
              All sides
            </Btn>
          </div>
          {(['borderTop', 'borderBottom', 'borderLeft', 'borderRight'] as const).map((side) => (
            <div key={side} className="flex items-center gap-3">
              <span className="w-16">{side.replace('border', '')}</span>
              <select aria-label={`${side} style`} className={cx(inputCls, 'h-7 w-40')} value={cur[side]?.style ?? 'none'} onChange={(e) => border(side, e.target.value !== 'none', e.target.value as BorderStyle)}>
                {['none', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair'].map((v) => (
                  <option key={v}>{v}</option>
                ))}
              </select>
            </div>
          ))}
        </div>
      )}
      {tab === 'fill' && (
        <div className="flex items-center gap-3 text-[12.5px]">
          Background color <input aria-label="Fill color" type="color" value={cur.fill ?? '#ffffff'} onChange={(e) => up({ fill: e.target.value })} />
          <Btn size="sm" onClick={() => up({ fill: undefined, pattern: undefined })}>
            No color
          </Btn>
        </div>
      )}
      {tab === 'protection' && (
        <div className="space-y-2 text-[12.5px]">
          <label className="flex items-center gap-2">
            <Switch checked={cur.locked !== false} onChange={(v) => up({ locked: v ? undefined : false })} label="Locked" /> Locked
          </label>
          <label className="flex items-center gap-2">
            <Switch checked={!!cur.hiddenFormula} onChange={(v) => up({ hiddenFormula: v || undefined })} label="Hidden" /> Hidden (formula not shown when the sheet is protected)
          </label>
          <p className="text-text-3">Locking takes effect when the sheet is protected (Review → Protect sheet arrives with the next milestone).</p>
        </div>
      )}
    </Modal>
  )
}

// ---------- sort ----------

function SortDialog() {
  const c = A.ctx()
  const [range, setRange] = useState<{ r0: number; c0: number; r1: number; c1: number } | null>(null)
  const [header, setHeader] = useState(true)
  const [levels, setLevels] = useState<{ col: number; order: 1 | -1 }[]>([])
  useEffect(() => {
    if (!c) return
    const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
    void (async () => {
      const rr = r.r0 === r.r1 && r.c0 === r.c1 ? await currentRegion(c) : r
      setRange(rr)
      setHeader(guessHeader(c, rr))
      setLevels([{ col: c.view.sel.active.c, order: 1 }])
    })()
  }, [])
  if (!c || !range) return null
  const label = (col: number) => (header ? (c.book.cell(c.view.sheet, range.r0, col)?.t || `Column ${colName(col)}`) : `Column ${colName(col)}`)
  const cols = Array.from({ length: Math.min(range.c1, range.c0 + 200) - range.c0 + 1 }, (_, i) => range.c0 + i)
  const apply = async () => {
    await A.edit(c, (b) => b.api.sort({ sheet: c.view.sheet, ...range }, levels.map((l) => ({ index: l.col, order: l.order })), header))
    close()
    afterDialog()
  }
  return (
    <Modal open onOpenChange={(o) => !o && close()} title="Sort" width={520} footer={<><Btn onClick={close}>Cancel</Btn><Btn variant="primary" onClick={apply}>OK</Btn></>}>
      <div className="mb-3 flex items-center justify-between text-[12.5px]">
        <Btn size="sm" onClick={() => setLevels((l) => [...l, { col: range.c0, order: 1 }])}>
          Add level
        </Btn>
        <label className="flex items-center gap-2">
          <Switch checked={header} onChange={setHeader} label="My data has headers" /> My data has headers
        </label>
      </div>
      {levels.map((l, i) => (
        <div key={i} className="mb-2 flex items-center gap-2 text-[12.5px]">
          <span className="w-16 text-text-3">{i ? 'Then by' : 'Sort by'}</span>
          <select aria-label="Column" className={cx(inputCls, 'h-8 flex-1')} value={l.col} onChange={(e) => setLevels((ls) => ls.map((x, k) => (k === i ? { ...x, col: Number(e.target.value) } : x)))}>
            {cols.map((col) => (
              <option key={col} value={col}>
                {label(col)}
              </option>
            ))}
          </select>
          <select aria-label="Order" className={cx(inputCls, 'h-8 w-40')} value={l.order} onChange={(e) => setLevels((ls) => ls.map((x, k) => (k === i ? { ...x, order: Number(e.target.value) as 1 | -1 } : x)))}>
            <option value={1}>A to Z / smallest first</option>
            <option value={-1}>Z to A / largest first</option>
          </select>
          {levels.length > 1 && (
            <Btn size="sm" variant="ghost" onClick={() => setLevels((ls) => ls.filter((_, k) => k !== i))}>
              Remove
            </Btn>
          )}
        </div>
      ))}
    </Modal>
  )
}

// ---------- row height / column width ----------

function SizeDialog({ axis }: { axis: 'row' | 'col' }) {
  const c = A.ctx()
  const [v, setV] = useState('')
  useEffect(() => {
    if (!c) return
    const m = c.book.sheetMeta(c.view.sheet)!
    const i = axis === 'row' ? c.view.sel.active.r : c.view.sel.active.c
    const list = axis === 'row' ? m.rows : m.cols
    const px = list.find(([k]) => k === i)?.[1] ?? (axis === 'row' ? m.defaultRowHeight : m.defaultColWidth)
    setV(String(axis === 'row' ? Math.round(px * 0.75 * 100) / 100 : Math.round(((px - 5) / 7) * 100) / 100))
  }, [])
  if (!c) return null
  // Excel units: row height in points, column width in characters of the default font.
  const apply = async () => {
    const n = Number(v)
    if (!Number.isFinite(n) || n < 0 || n > (axis === 'row' ? 409 : 255)) return toast(axis === 'row' ? 'Row height must be between 0 and 409.' : 'Column width must be between 0 and 255.', 'error')
    const px = axis === 'row' ? Math.round(n / 0.75) : n === 0 ? 0 : Math.round(n * 7 + 5)
    const idx: number[] = []
    for (const r of c.view.sel.ranges) for (let i = axis === 'row' ? r.r0 : r.c0; i <= Math.min(axis === 'row' ? r.r1 : r.c1, (axis === 'row' ? r.r0 : r.c0) + 100_000); i++) idx.push(i)
    await A.edit(c, (b) => (axis === 'row' ? b.api.setRowHeights(c.view.sheet, idx, px) : b.api.setColWidths(c.view.sheet, idx, px)))
    close()
    afterDialog()
  }
  return (
    <Modal open onOpenChange={(o) => !o && close()} title={axis === 'row' ? 'Row height' : 'Column width'} width={320} footer={<><Btn onClick={close}>Cancel</Btn><Btn variant="primary" onClick={apply}>OK</Btn></>}>
      <Field label={axis === 'row' ? 'Row height (points)' : 'Column width (characters)'}>
        <input autoFocus aria-label={axis === 'row' ? 'Row height' : 'Column width'} className={inputCls} value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && apply()} />
      </Field>
    </Modal>
  )
}

// ---------- shortcuts ----------

function ShortcutsDialog() {
  const [q, setQ] = useState('')
  const [recording, setRecording] = useState<string | null>(null)
  const settings = useApp((s) => s.settings)
  const cmds = allCommands().filter((c) => !q || `${c.label} ${c.id}`.toLowerCase().includes(q.toLowerCase()) || shortcutsFor(c.id).some((k) => formatShortcut(k).toLowerCase().includes(q.toLowerCase())))
  // Conflict detection: the same key on two commands.
  const owners = new Map<string, string[]>()
  for (const c of allCommands()) for (const k of shortcutsFor(c.id)) owners.set(k, [...(owners.get(k) ?? []), c.id])
  useEffect(() => {
    if (!recording) return
    const onKey = (e: KeyboardEvent) => {
      e.preventDefault()
      e.stopPropagation()
      if (e.key === 'Escape') return setRecording(null)
      if (['Shift', 'Control', 'Alt', 'Meta'].includes(e.key)) return
      const combo = eventToShortcut(e)
      void updateSettings({ shortcuts: { ...get().settings.shortcuts, [recording]: combo } })
      setRecording(null)
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [recording])
  const exportJson = async () => {
    const path = await window.simpleCell.saveDialog('Simple Cell shortcuts.json', ['json'])
    if (path) await window.simpleCell.writeFile(path, new TextEncoder().encode(JSON.stringify({ preset: settings.shortcutPreset, shortcuts: settings.shortcuts }, null, 2)))
  }
  const importJson = async () => {
    const [f] = await window.simpleCell.openDialog({ name: 'Shortcuts', extensions: ['json'] })
    if (!f) return
    try {
      const data = JSON.parse(new TextDecoder().decode(f.data)) as { preset?: Settings['shortcutPreset']; shortcuts?: Record<string, string> }
      const clean = Object.fromEntries(Object.entries(data.shortcuts ?? {}).filter(([k, v]) => getCommand(k) && typeof v === 'string'))
      await updateSettings({ shortcuts: clean, ...(data.preset === 'excel' || data.preset === 'sheets' ? { shortcutPreset: data.preset } : {}) })
      toast('Shortcuts imported', 'success')
    } catch {
      toast('That file isn’t a Simple Cell shortcut file.', 'error')
    }
  }
  return (
    <Modal open onOpenChange={(o) => !o && close()} title="Keyboard shortcuts" description="Click a shortcut to change it. Shortcuts follow Microsoft Excel; the Google Sheets preset changes the ones that differ." width={720} footer={<><Btn onClick={importJson}>Import…</Btn><Btn onClick={exportJson}>Export…</Btn><Btn onClick={() => updateSettings({ shortcuts: {} })}>Reset all</Btn><Btn variant="primary" onClick={close}>Done</Btn></>}>
      <div className="mb-3 flex items-center gap-3">
        <input autoFocus aria-label="Search shortcuts" placeholder="Search by action or key" className={inputCls} value={q} onChange={(e) => setQ(e.target.value)} />
        <Segmented value={settings.shortcutPreset} onChange={(v) => updateSettings({ shortcutPreset: v })} label="Preset" options={[{ value: 'excel', label: 'Excel' }, { value: 'sheets', label: 'Google Sheets' }]} />
      </div>
      <table className="w-full text-[12.5px]" data-testid="shortcut-table">
        <tbody>
          {cmds.map((c) => {
            const keys = shortcutsFor(c.id)
            const conflict = keys.some((k) => (owners.get(k)?.length ?? 0) > 1)
            return (
              <tr key={c.id} className="border-t border-border">
                <td className="py-1.5">{c.label}</td>
                <td className="py-1.5 text-right">
                  <button className={cx('inline-flex min-w-[96px] justify-end gap-1 rounded px-2 py-0.5 hover:bg-hover', conflict && 'text-danger')} onClick={() => setRecording(c.id)} aria-label={`Change shortcut for ${c.label}`}>
                    {recording === c.id ? <span className="text-accent-ink">Press keys…</span> : keys.length ? keys.map((k) => <Kbd key={k}>{formatShortcut(k)}</Kbd>) : <span className="text-text-3">—</span>}
                  </button>
                  {conflict && <span className="ml-1 text-[11px] text-danger">conflict</span>}
                </td>
              </tr>
            )
          })}
          <tr>
            <td colSpan={2} className="pb-1 pt-4 text-[11px] font-semibold uppercase tracking-wide text-text-3">
              Navigation and editing
            </td>
          </tr>
          {BUILTIN.filter((b) => !q || `${b.action} ${b.keys}`.toLowerCase().includes(q.toLowerCase())).map((b) => (
            <tr key={b.keys + b.action} className="border-t border-border">
              <td className="py-1.5">
                {b.action}
                {b.status === 'planned' && <span className="ml-2 rounded bg-hover px-1.5 text-[10.5px] text-text-3">coming</span>}
              </td>
              <td className="py-1.5 text-right">
                <Kbd>{formatShortcut(b.keys)}</Kbd>
              </td>
            </tr>
          ))}
          <tr>
            <td colSpan={2} className="pb-1 pt-4 text-[11px] font-semibold uppercase tracking-wide text-text-3">
              Reserved by the operating system
            </td>
          </tr>
          {RESERVED.map((r) => (
            <tr key={r.keys + r.os} className="border-t border-border text-text-3">
              <td className="py-1.5">
                {r.note} ({r.os})
              </td>
              <td className="py-1.5 text-right">
                <Kbd>{formatShortcut(r.keys)}</Kbd>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </Modal>
  )
}

// ---------- settings ----------

const ACCENTS = ['#1f8a5b', '#5b5bd6', '#2563eb', '#0891b2', '#c2410c', '#be185d', '#7c3aed', '#3f3f46']

function Row({ label, children, hint }: { label: string; children: ReactNode; hint?: string }) {
  return (
    <div className="flex items-center justify-between gap-6 py-2.5">
      <div>
        <div className="text-[13px] font-medium">{label}</div>
        {hint && <div className="text-[12px] text-text-3">{hint}</div>}
      </div>
      {children}
    </div>
  )
}

function SettingsDialog({ initial }: { initial?: string }) {
  const s = useApp((x) => x.settings)
  const [section, setSection] = useState(initial ?? 'general')
  const [version, setVersion] = useState('')
  useEffect(() => {
    void window.simpleCell.appVersion().then(setVersion)
  }, [])
  const sections = [
    ['general', 'General'],
    ['editing', 'Editing'],
    ['keyboard', 'Keyboard'],
    ['updates', 'Updates']
  ]
  return (
    <Modal open onOpenChange={(o) => !o && close()} title="Settings" width={680} footer={<Btn variant="primary" onClick={close}>Done</Btn>}>
      <div className="flex gap-5">
        <nav className="w-36 shrink-0 space-y-0.5">
          {sections.map(([id, label]) => (
            <button key={id} onClick={() => setSection(id)} className={cx('h-8 w-full rounded-lg px-3 text-left text-[13px]', section === id ? 'bg-accent-soft font-medium text-accent-ink' : 'text-text-2 hover:bg-hover')}>
              {label}
            </button>
          ))}
        </nav>
        <div className="min-h-[300px] flex-1 divide-y divide-border">
          {section === 'general' && (
            <>
              <Row label="Theme">
                <Segmented value={s.theme} onChange={(v) => updateSettings({ theme: v })} label="Theme" options={[{ value: 'system', label: <Monitor className="size-3.5" />, title: 'System' }, { value: 'light', label: <Sun className="size-3.5" />, title: 'Light' }, { value: 'dark', label: <Moon className="size-3.5" />, title: 'Dark' }]} />
              </Row>
              <Row label="Accent color">
                <div className="flex gap-1.5">
                  {ACCENTS.map((a) => (
                    <button key={a} aria-label={`Accent ${a}`} onClick={() => updateSettings({ accent: a })} className={cx('size-6 rounded-full', s.accent === a && 'ring-2 ring-text ring-offset-2 ring-offset-elev')} style={{ background: a }} />
                  ))}
                </div>
              </Row>
              <Row label="Ribbon" hint="The full ribbon with group labels, or a single compact toolbar row">
                <Segmented value={s.ribbonMode} onChange={(v) => updateSettings({ ribbonMode: v })} label="Ribbon layout" options={[{ value: 'expanded', label: 'Full' }, { value: 'compact', label: 'Compact' }]} />
              </Row>
              <Row label="Interface size">
                <Segmented value={String(s.uiScale)} onChange={(v) => updateSettings({ uiScale: Number(v) })} label="Interface size" options={['0.9', '1', '1.1', '1.25'].map((v) => ({ value: v, label: `${Math.round(Number(v) * 100)}%` }))} />
              </Row>
              <Row label="Reduce motion">
                <Switch checked={s.reduceMotion} onChange={(v) => updateSettings({ reduceMotion: v })} label="Reduce motion" />
              </Row>
            </>
          )}
          {section === 'editing' && (
            <>
              <Row label="Default font">
                <input aria-label="Default font" className={cx(inputCls, 'w-48')} value={s.defaultFont} onChange={(e) => updateSettings({ defaultFont: e.target.value })} />
              </Row>
              <Row label="Default font size">
                <input aria-label="Default font size" type="number" min={6} max={72} className={cx(inputCls, 'w-24')} value={s.defaultFontSize} onChange={(e) => updateSettings({ defaultFontSize: Number(e.target.value) || 11 })} />
              </Row>
              <Row label="AutoSave" hint="Saves workbooks that already have a file, a few seconds after you stop typing">
                <Switch checked={s.autosave} onChange={(v) => updateSettings({ autosave: v })} label="AutoSave" />
              </Row>
              <Row label="Recovery snapshots" hint={`Unsaved changes are kept every ${s.recoveryIntervalSec} s in case the app closes unexpectedly`}>
                <input aria-label="Recovery interval" type="number" min={2} max={300} className={cx(inputCls, 'w-24')} value={s.recoveryIntervalSec} onChange={(e) => updateSettings({ recoveryIntervalSec: Math.max(2, Number(e.target.value) || 5) })} />
              </Row>
              <Row label="Protect CSV exports" hint="Text starting with = + - @ gets a leading apostrophe, so other apps don’t run it as a formula">
                <Switch checked={s.csvFormulaGuard} onChange={(v) => updateSettings({ csvFormulaGuard: v })} label="Protect CSV exports" />
              </Row>
            </>
          )}
          {section === 'keyboard' && (
            <>
              <Row label="Shortcut preset" hint="Excel shortcuts, or Google Sheets where they differ">
                <Segmented value={s.shortcutPreset} onChange={(v) => updateSettings({ shortcutPreset: v })} label="Preset" options={[{ value: 'excel', label: 'Excel' }, { value: 'sheets', label: 'Google Sheets' }]} />
              </Row>
              <Row label="Customize shortcuts">
                <Btn onClick={() => set({ dialog: { kind: 'shortcuts' } })}>Open shortcut list</Btn>
              </Row>
            </>
          )}
          {section === 'updates' && (
            <>
              <Row label="Check for updates automatically" hint="Once a day Simple Cell asks GitHub whether a newer version exists. Nothing downloads until you click Install.">
                <Switch checked={s.updateCheck} onChange={(v) => updateSettings({ updateCheck: v })} label="Check for updates automatically" />
              </Row>
              <Row label={`Version ${version}`}>
                <Btn onClick={() => window.simpleCell.checkForUpdates()}>Check now</Btn>
              </Row>
            </>
          )}
        </div>
      </div>
    </Modal>
  )
}

function AboutDialog() {
  const [version, setVersion] = useState('')
  useEffect(() => {
    void window.simpleCell.appVersion().then(setVersion)
  }, [])
  return (
    <Modal open onOpenChange={(o) => !o && close()} title="Simple Cell" description={`Version ${version}`} width={440} footer={<Btn variant="primary" onClick={close}>OK</Btn>}>
      <p className="text-[13px] leading-relaxed text-text-2">An open-source spreadsheet that works offline. MIT licence. No account, no telemetry: the only network request is the daily update check, which you can turn off.</p>
    </Modal>
  )
}

export { books }
