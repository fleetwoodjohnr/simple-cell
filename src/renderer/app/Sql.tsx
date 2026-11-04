// Automate ▸ SQL: query tables and selected data; return ordinary ranges. Each run is one undo step.
import { useEffect, useRef, useState } from 'react'
import type { SqlOut } from '@/core/sql'
import { areaName, cellName, parseCell, parseArea } from '@/core/refs'
import { toast } from '../state/store'
import { setView } from '../state/view'
import { Btn, Modal, Segmented, cx, inputCls } from '../ui'
import * as A from './actions'
import { closeDialog, focusGrid } from './focus'
import { currentRegion, guessHeader } from './commands'

type Preview = { columns: string[] | null; rows: (string | number | boolean | null)[][]; rowCount: number; changed: { table: string; rows: number }[] }

const done = () => (closeDialog(), focusGrid())
/** A name as SQL needs it: bare when it is a plain identifier, otherwise double-quoted. */
const ident = (n: string) => (/^[A-Za-z_]\w*$/.test(n) ? n : `"${n.replace(/"/g, '""')}"`)

export function SqlDialog() {
  const c = A.ctx()
  const [tables, setTables] = useState<{ name: string; sheet: number; columns: string[] }[] | null>(null)
  const [sql, setSql] = useState('')
  const [where, setWhere] = useState<'sheet' | 'cell'>('sheet')
  const [outCell, setOutCell] = useState('')
  const [sourceRef, setSourceRef] = useState('')
  const [headers, setHeaders] = useState(true)
  const [useRange, setUseRange] = useState(false)
  const [rangeInfo, setRangeInfo] = useState<{ name: string; sheet: number; columns: string[] } | null>(null)
  const [outSheet, setOutSheet] = useState(c?.view.sheet ?? 0)
  const [preview, setPreview] = useState<Preview | null>(null)
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const returnCleanup = useRef<() => void>(() => {})
  useEffect(() => () => returnCleanup.current(), [])
  const editor = useRef<HTMLTextAreaElement>(null)
  useEffect(() => {
    if (!c) return
    const { r, c: col } = c.view.sel.active
    setOutCell(cellName(r, col))
    const selected = c.view.sel.ranges.at(-1)!
    void (async () => {
      const range = selected.r0 === selected.r1 && selected.c0 === selected.c1 ? await currentRegion(c) : selected
      setSourceRef(areaName(range)); setHeaders(guessHeader(c, range))
      const here = c.book.meta.tables.some((t) => t.sheet === c.view.sheet && r >= t.area.r0 && r <= t.area.r1 && col >= t.area.c0 && col <= t.area.c1)
      if (!here) {
        let alias = 'Selection'
        for (let n = 2; c.book.meta.tables.some((t) => t.name.toLowerCase() === alias.toLowerCase()); n++) alias = `Selection${n}`
        setUseRange(true); setSql(`SELECT *\nFROM ${alias}`)
      }
    })()
    returnCleanup.current = () => { void c.book.api.cancelSql() }
    void c.book.api.sqlTables().then((ts) => {
      setTables(ts)
      const here = c.book.meta.tables.find((t) => t.sheet === c.view.sheet && r >= t.area.r0 && r <= t.area.r1 && col >= t.area.c0 && col <= t.area.c1)
      const first = here?.name
      if (first) {
        setSql(`SELECT *\nFROM ${ident(first)}`)
      }
    })
  }, [])
  useEffect(() => {
    if (!c || !sourceRef || !useRange) { setRangeInfo(null); return }
    const area = parseArea(sourceRef.replace(/\$/g, ''), c.view.sheet)
    let live = true
    if (area) void c.book.api.sqlRange({ area, headers }).then((info) => { if (live) setRangeInfo(info) }).catch(() => { if (live) setRangeInfo(null) })
    else setRangeInfo(null)
    return () => { live = false }
  }, [sourceRef, headers, useRange])
  if (!c || !tables) return null
  const source = () => {
    if (!useRange) return undefined
    const area = parseArea(sourceRef.replace(/\$/g, ''), c.view.sheet)
    if (!area || !rangeInfo) throw new Error('Enter a valid source range.')
    return { area, headers }
  }
  const available = [...tables, ...(useRange && rangeInfo ? [rangeInfo] : [])]

  const insert = (text: string) => {
    const el = editor.current
    if (!el) return
    el.focus()
    el.setRangeText(text, el.selectionStart, el.selectionEnd, 'end')
    setSql(el.value)
  }
  const runPreview = async () => {
    setBusy(true)
    let selected
    try { selected = source() } catch (e) { setBusy(false); return setError((e as Error).message) }
    const r = await c.book.api.sqlPreview(sql, selected)
    setBusy(false)
    setError(r.ok ? '' : r.error)
    setPreview(r.ok ? r : null)
  }
  const run = async () => {
    let out: SqlOut = { kind: 'sheet' }
    if (where === 'cell') {
      const o = parseCell(outCell.replace(/\$/g, ''))
      if (!o) return setError('Enter the top-left cell for the result, like H1.')
      out = { kind: 'cell', sheet: outSheet, ...o }
    }
    setBusy(true)
    let selected
    try { selected = source() } catch (e) { setBusy(false); return setError((e as Error).message) }
    const r = await c.book.api.runSql(sql, out, selected)
    setBusy(false)
    if (!r.ok) return setError(r.error)
    A.markDirtyTab(c.tab)
    await c.book.refreshMeta()
    if (r.sheet !== undefined && r.sheet !== c.view.sheet) setView(c.tab, { sheet: r.sheet })
    done()
    const parts = [r.changed.length ? `Changed ${r.changed.join(', ')}` : '', r.rows !== null ? `${r.rows} row${r.rows === 1 ? '' : 's'} in ${r.table}` : '']
    toast(parts.filter(Boolean).join('; ') + '.', 'success')
  }

  return (
    <Modal
      open
      onOpenChange={(o) => !o && (void c.book.api.cancelSql(), done())}
      title="SQL query"
      description="SQLite over tables and selected data. Ctrl+Enter previews. Results are limited to one million cells."
      width={780}
      footer={
        <>
          <Btn onClick={() => { void c.book.api.cancelSql(); done() }}>Cancel</Btn>
          <Btn disabled={busy} onClick={() => void runPreview()}>
            Preview
          </Btn>
          <Btn variant="primary" disabled={busy} onClick={() => void run()}>
            Run
          </Btn>
        </>
      }
    >
      <div className="mb-3 space-y-2 text-[12.5px]">
        <label className="flex items-center gap-2"><input type="checkbox" checked={useRange} onChange={(e) => { setUseRange(e.target.checked); setPreview(null) }} /> Include selected dataset as {rangeInfo?.name ?? 'Selection'}</label>
        {useRange && <div className="flex items-center gap-2"><label>Source range <input aria-label="Source range" className={cx(inputCls, 'w-32')} value={sourceRef} onChange={(e) => { setSourceRef(e.target.value); setPreview(null) }} /></label><label className="flex items-center gap-1"><input type="checkbox" checked={headers} onChange={(e) => { setHeaders(e.target.checked); setPreview(null) }} /> First row contains headers</label></div>}
      </div>
      <div className="flex gap-3">
        <div role="list" aria-label="Tables and columns" className="max-h-[320px] w-44 shrink-0 overflow-auto rounded-lg border border-border p-1 text-[12.5px]">
          {available.map((t) => (
            <div role="listitem" key={t.name}>
              <button type="button" title="Insert" className="min-h-6 w-full truncate rounded px-1.5 py-0.5 text-left font-medium hover:bg-hover" onClick={() => insert(ident(t.name))}>
                {t.name}
              </button>
              {t.columns.map((col) => (
                <button type="button" title="Insert" key={col} className="min-h-6 w-full truncate rounded py-0.5 pl-4 pr-1.5 text-left text-text-2 hover:bg-hover" onClick={() => insert(`"${col.replace(/"/g, '""')}"`)}>
                  {col}
                </button>
              ))}
            </div>
          ))}
        </div>
        <div className="min-w-0 flex-1 space-y-2.5">
          <textarea
            ref={editor}
            autoFocus
            aria-label="SQL query"
            spellCheck={false}
            value={sql}
            onChange={(e) => { setSql(e.target.value); setPreview(null); setError('') }}
            onKeyDown={(e) => {
              if (!busy && (e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                e.preventDefault()
                void runPreview()
              }
            }}
            className={cx(inputCls, 'h-36 resize-y py-1.5 font-mono text-[12.5px] leading-snug')}
          />
          <p className="text-[12px] text-text-3">UPDATE, DELETE and INSERT edit their source data in place. Preview changes nothing; Run is one undo step. The rows a query returns go to:</p>
          <div className="flex items-center gap-2">
            <Segmented value={where} onChange={setWhere} label="Put the result" options={[{ value: 'sheet', label: 'New sheet' }, { value: 'cell', label: 'Output range' }]} />
            {where === 'cell' && <input aria-label="Top-left cell" className={cx(inputCls, 'w-24')} value={outCell} onChange={(e) => setOutCell(e.target.value)} />}
            {where === 'cell' && <select aria-label="Output sheet" className={cx(inputCls, 'w-32')} value={outSheet} onChange={(e) => setOutSheet(Number(e.target.value))}>{c.book.meta.sheets.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select>}
          </div>
          {error && (
            <p role="alert" className="text-[12px] text-danger">
              {error}
            </p>
          )}
          {preview && <PreviewTable p={preview} />}
        </div>
      </div>
    </Modal>
  )
}

function PreviewTable({ p }: { p: Preview }) {
  const changed = p.changed.map((x) => `${x.table} (${x.rows} row${x.rows === 1 ? '' : 's'})`).join(', ')
  return (
    <div className="space-y-1.5">
      <p className="text-[12px] text-text-2" data-testid="sql-summary">
        {p.columns ? `${p.rowCount} row${p.rowCount === 1 ? '' : 's'}${p.rowCount > p.rows.length ? ` (first ${p.rows.length} shown)` : ''}` : 'No rows returned'}
        {changed && ` · Will change ${changed}`}
      </p>
      {p.columns && (
        <div className="max-h-[220px] overflow-auto rounded-lg border border-border">
          <table className="w-full border-collapse text-[12px] tabular-nums">
            <thead className="sticky top-0 bg-panel">
              <tr>
                {p.columns.map((col, i) => (
                  <th key={i} className="border-b border-border px-2 py-1 text-left font-medium">
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {p.rows.map((row, j) => (
                <tr key={j}>
                  {row.map((v, i) => (
                    <td key={i} className={cx('border-b border-border px-2 py-0.5', typeof v === 'number' && 'text-right')}>
                      {v === null ? '' : String(v)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
