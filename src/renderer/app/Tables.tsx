// Tables and filters in the UI: Create Table, the AutoFilter header buttons and menu, Table Design options.
import { useEffect, useMemo, useRef, useState } from 'react'
import { ArrowDownAZ, ArrowUpAZ, ChevronDown, Filter } from 'lucide-react'
import type { FilterColumn, FilterOp } from '@/core/filter'
import { TABLE_STYLES, tableCellStyle } from '@/core/tablestyle'
import type { Area } from '@/core/engine/values'
import { areaName, parseArea } from '@/core/refs'
import { colX, rowY, type Layout } from '@/grid/layout'
import type { Book } from '../state/client'
import { set, toast, type Dialog } from '../state/store'
import { getView, setView, useViews } from '../state/view'
import { Btn, Field, Modal, cx, inputCls } from '../ui'
import * as A from './actions'
import { markDirty } from '../state/documents'
import { currentRegion, guessHeader } from './commands'
import { closeDialog, focusGrid } from './focus'

const close = closeDialog
const refocus = focusGrid

// ---------- Create table ----------

export async function openCreateTable(style?: string) {
  const c = A.ctx()
  if (!c) return
  const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  const range = r.r0 === r.r1 && r.c0 === r.c1 ? await currentRegion(c) : r
  if (range.r0 === range.r1 && range.c0 === range.c1 && !c.book.cell(c.view.sheet, range.r0, range.c0)?.t) {
    toast('Select the data for the table first (a cell inside it is enough).', 'info')
    return
  }
  set({ dialog: { kind: 'create-table', area: range, headers: guessHeader(c, range), style } })
}

export function CreateTableDialog({ d }: { d: Extract<Dialog, { kind: 'create-table' }> }) {
  const c = A.ctx()
  const [ref, setRef] = useState(areaName(d.area).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'))
  const [headers, setHeaders] = useState(d.headers)
  const [error, setError] = useState('')
  if (!c) return null
  const ok = async () => {
    const a = parseArea(ref.replace(/\$/g, ''), c.view.sheet)
    if (!a) return setError('Enter a range such as $A$1:$D$20.')
    const r = await c.book.api.createTable(c.view.sheet, a, headers, d.style)
    if (!r.ok) return setError(r.error ?? 'The table could not be created.')
    markDirty(c.tab)
    close()
    refocus()
  }
  return (
    <Modal open onOpenChange={(o) => !o && (close(), refocus())} title="Create table" width={380} footer={<><Btn onClick={() => (close(), refocus())}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
      <Field label="Where is the data for your table?">
        <input aria-label="Table range" autoFocus className={inputCls} value={ref} onChange={(e) => setRef(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && void ok()} />
      </Field>
      <label className="mt-3 flex items-center gap-2 text-[12.5px]">
        <input type="checkbox" checked={headers} onChange={(e) => setHeaders(e.target.checked)} /> My table has headers
      </label>
      {error && <p role="alert" className="mt-2 text-[12px] text-danger">{error}</p>}
    </Modal>
  )
}

/** A small preview of a table style (header, two stripes). */
export function StyleSwatch({ name }: { name: string }) {
  const t = { id: 0, name: 'p', displayName: 'p', sheet: 0, area: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 }, headerRow: true, totalsRow: false, columns: [], style: name, showRowStripes: true }
  return (
    <span aria-hidden className="grid h-[30px] w-[46px] grid-cols-3 grid-rows-4 overflow-hidden rounded-sm border border-border bg-white">
      {Array.from({ length: 12 }, (_, i) => {
        const s = tableCellStyle(t, Math.floor(i / 3), i % 3)
        return <span key={i} style={{ background: s?.fill ?? '#fff', borderBottom: s?.borderBottom ? `1px solid ${s.borderBottom.color}` : undefined }} />
      })}
    </span>
  )
}

/** The Format as Table gallery (Home) and the style picker (Table Design). */
export function StyleGallery({ onPick, current }: { onPick: (style: string) => void; current?: string }) {
  const groups: [string, string][] = [
    ['Light', 'TableStyleLight'],
    ['Medium', 'TableStyleMedium'],
    ['Dark', 'TableStyleDark']
  ]
  return (
    <div className="max-h-[420px] w-[392px] overflow-y-auto p-2">
      {groups.map(([label, prefix]) => (
        <div key={label} className="mb-2">
          <div className="mb-1 px-1 text-[11px] font-semibold text-text-3">{label}</div>
          <div className="grid grid-cols-7 gap-1">
            {TABLE_STYLES.filter((s) => s.startsWith(prefix) && /\d+$/.test(s.slice(prefix.length))).map((s) => (
              <button key={s} type="button" title={s.replace('TableStyle', '')} aria-label={s.replace('TableStyle', 'Table style ')} onClick={() => onPick(s)} className={cx('rounded p-0.5 hover:bg-hover', current === s && 'ring-2 ring-accent')}>
                <StyleSwatch name={s} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

// ---------- filter buttons and menu ----------

interface FilterRange {
  key: string
  area: Area
  active: number[]
}

/** Header buttons of every filter range on the sheet; the open filter menu. */
export function FilterOverlay({ tab, book, L }: { tab: string; book: Book; L: Layout }) {
  const view = useViews((s) => s.views[tab])
  const meta = book.sheetMeta(view.sheet)
  const filters = (meta?.filters ?? []) as FilterRange[]
  const open = view.filterMenu
  return (
    <>
      {filters.flatMap((f) =>
        Array.from({ length: f.area.c1 - f.area.c0 + 1 }, (_, i) => {
          const c = f.area.c0 + i
          const r = f.area.r0
          const w = L.cols.size(c) * L.zoom
          const h = L.rows.size(r) * L.zoom
          const x = colX(L, c)
          const y = rowY(L, r)
          if (!w || !h || x + w < L.headerW || y < L.headerH - 1 || x > L.width || y > L.height) return []
          const s = Math.min(17, h - 2)
          const on = f.active.includes(i)
          return [
            <button
              key={`${f.key}:${c}`}
              type="button"
              tabIndex={-1}
              data-testid={`filter-button-${areaName({ r0: r, c0: c, r1: r, c1: c })}`}
              aria-label={`Filter ${book.cell(view.sheet, r, c)?.t || `column ${i + 1}`}${on ? ' (filtered)' : ''}`}
              aria-haspopup="dialog"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setView(tab, (v) => ({ filterMenu: v.filterMenu?.key === f.key && v.filterMenu.col === i ? undefined : { key: f.key, col: i } }))}
              className={cx('absolute z-[4] flex items-center justify-center rounded-sm border border-border bg-panel text-text-2 shadow-sm hover:bg-hover', on && 'text-accent')}
              style={{ left: x + w - s - 2, top: y + h - s - 1, width: s, height: s }}
            >
              {on ? <Filter className="size-3" /> : <ChevronDown className="size-3.5" />}
            </button>
          ]
        })
      )}
      {open && filters.some((f) => f.key === open.key) && <FilterMenu tab={tab} book={book} L={L} range={filters.find((f) => f.key === open.key)!} col={open.col} />}
    </>
  )
}

const OPS: [string, string, (v: string) => { op: FilterOp; val: string }][] = [
  ['equal', 'equals', (v) => ({ op: 'equal', val: v })],
  ['notEqual', 'does not equal', (v) => ({ op: 'notEqual', val: v })],
  ['contains', 'contains', (v) => ({ op: 'equal', val: `*${v}*` })],
  ['notContains', 'does not contain', (v) => ({ op: 'notEqual', val: `*${v}*` })],
  ['begins', 'begins with', (v) => ({ op: 'equal', val: `${v}*` })],
  ['ends', 'ends with', (v) => ({ op: 'equal', val: `*${v}` })],
  ['greaterThan', 'is greater than', (v) => ({ op: 'greaterThan', val: v })],
  ['greaterThanOrEqual', 'is greater than or equal to', (v) => ({ op: 'greaterThanOrEqual', val: v })],
  ['lessThan', 'is less than', (v) => ({ op: 'lessThan', val: v })],
  ['lessThanOrEqual', 'is less than or equal to', (v) => ({ op: 'lessThanOrEqual', val: v })]
]

function FilterMenu({ tab, book, L, range, col }: { tab: string; book: Book; L: Layout; range: FilterRange; col: number }) {
  const view = getView(tab)
  const [data, setData] = useState<{ values: string[]; blanks: boolean; current: FilterColumn | null } | null>(null)
  const [checked, setChecked] = useState<Set<string>>(new Set())
  const [blank, setBlank] = useState(true)
  const [q, setQ] = useState('')
  const [mode, setMode] = useState<'list' | 'custom'>('list')
  const [op, setOp] = useState('contains')
  const [val, setVal] = useState('')
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    void book.api.filterValues(view.sheet, range.key, col).then((d) => {
      if (!d) return
      setData(d)
      const cur = d.current
      setChecked(new Set(cur?.values ? cur.values.map((v) => v.toLowerCase()) : d.values.map((v) => v.toLowerCase())))
      setBlank(cur?.values ? !!cur.blanks : true)
      if (cur?.custom) {
        setMode('custom')
        const cd = cur.custom.conds[0]
        const found = OPS.find(([, , f]) => {
          const inner = cd.val.replace(/^\*|\*$/g, '')
          const x = f(inner)
          return x.op === cd.op && x.val === cd.val
        })
        if (found) {
          setOp(found[0])
          setVal(cd.val.replace(/^\*|\*$/g, ''))
        }
      }
    })
    // Now, not a frame later (keys typed meanwhile were lost).
    ref.current?.querySelector<HTMLElement>('input')?.focus()
  }, [range.key, col])
  const shown = useMemo(() => (data?.values ?? []).filter((v) => v.toLowerCase().includes(q.toLowerCase())), [data, q])
  const done = () => {
    setView(tab, { filterMenu: undefined })
    refocus()
  }
  const apply = async (fc: Omit<FilterColumn, 'col'> | null) => {
    const c = A.ctx()
    if (!c) return
    await A.edit(c, (b) => b.api.setFilter(view.sheet, range.key, col, fc))
    done()
  }
  const ok = () => {
    if (mode === 'custom') {
      if (!val.trim()) return void apply(null)
      return void apply({ custom: { conds: [OPS.find(([k]) => k === op)![2](val.trim())] } })
    }
    // With a search, only the matching values that are ticked are kept (as in Excel).
    if (q) return void apply({ values: shown.filter((v) => checked.has(v.toLowerCase())) })
    const all = (data?.values ?? []).every((v) => checked.has(v.toLowerCase())) && (blank || !data?.blanks)
    if (all) return void apply(null)
    void apply({ values: (data?.values ?? []).filter((v) => checked.has(v.toLowerCase())), blanks: blank || undefined })
  }
  const sort = async (order: 1 | -1) => {
    const c = A.ctx()
    if (!c) return
    const a = range.area
    await A.edit(c, (b) => b.api.sort({ sheet: view.sheet, r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }, [{ index: a.c0 + col, order }], true))
    done()
  }
  const c = range.area.c0 + col
  const x = Math.max(L.headerW, Math.min(colX(L, c), L.width - 270))
  const y = rowY(L, range.area.r0) + L.rows.size(range.area.r0) * L.zoom + 2
  const allOn = shown.every((v) => checked.has(v.toLowerCase())) && (!data?.blanks || blank)
  return (
    <div
      ref={ref}
      role="dialog"
      aria-label="Filter"
      data-testid="filter-menu"
      className="absolute z-[8] w-[260px] rounded-lg border border-border bg-panel p-2 text-[12.5px] shadow-xl"
      style={{ left: x, top: y }}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          done()
        }
      }}
    >
      <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1 text-left hover:bg-hover" onClick={() => void sort(1)}>
        <ArrowDownAZ className="size-4" /> Sort A to Z
      </button>
      <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1 text-left hover:bg-hover" onClick={() => void sort(-1)}>
        <ArrowUpAZ className="size-4" /> Sort Z to A
      </button>
      <button type="button" disabled={!range.active.includes(col)} className="flex w-full items-center gap-2 rounded px-2 py-1 text-left hover:bg-hover disabled:opacity-40" onClick={() => void apply(null)}>
        <Filter className="size-4" /> Clear filter
      </button>
      <div className="my-1.5 h-px bg-border" />
      <div className="mb-1.5 flex gap-1">
        {(['list', 'custom'] as const).map((m) => (
          <button key={m} type="button" aria-pressed={mode === m} onClick={() => setMode(m)} className={cx('flex-1 rounded px-2 py-1', mode === m ? 'bg-accent-soft text-accent-ink' : 'hover:bg-hover')}>
            {m === 'list' ? 'Values' : 'Condition'}
          </button>
        ))}
      </div>
      {mode === 'list' ? (
        <>
          <input aria-label="Search values" className={cx(inputCls, 'mb-1.5 h-7')} placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && ok()} />
          <div className="max-h-[200px] overflow-y-auto rounded border border-border p-1">
            <label className="flex items-center gap-2 px-1 py-0.5">
              <input
                type="checkbox"
                checked={allOn}
                onChange={(e) => {
                  const next = new Set(checked)
                  for (const v of shown) e.target.checked ? next.add(v.toLowerCase()) : next.delete(v.toLowerCase())
                  setChecked(next)
                  if (data?.blanks && !q) setBlank(e.target.checked)
                }}
              />
              (Select all)
            </label>
            {shown.map((v) => (
              <label key={v} className="flex items-center gap-2 px-1 py-0.5">
                <input
                  type="checkbox"
                  checked={checked.has(v.toLowerCase())}
                  onChange={(e) => {
                    const next = new Set(checked)
                    if (e.target.checked) next.add(v.toLowerCase())
                    else next.delete(v.toLowerCase())
                    setChecked(next)
                  }}
                />
                <span className="truncate">{v}</span>
              </label>
            ))}
            {data?.blanks && !q && (
              <label className="flex items-center gap-2 px-1 py-0.5">
                <input type="checkbox" checked={blank} onChange={(e) => setBlank(e.target.checked)} /> (Blanks)
              </label>
            )}
          </div>
        </>
      ) : (
        <div className="space-y-1.5">
          <select aria-label="Condition" className={cx(inputCls, 'h-7')} value={op} onChange={(e) => setOp(e.target.value)}>
            {OPS.map(([k, l]) => (
              <option key={k} value={k}>
                {l}
              </option>
            ))}
          </select>
          <input aria-label="Condition value" className={cx(inputCls, 'h-7')} value={val} onChange={(e) => setVal(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && ok()} />
          <p className="text-[11px] text-text-3">Use ? for any one character and * for any run of characters.</p>
        </div>
      )}
      <div className="mt-2 flex justify-end gap-1.5">
        <Btn size="sm" onClick={done}>
          Cancel
        </Btn>
        <Btn size="sm" variant="primary" onClick={ok}>
          OK
        </Btn>
      </div>
    </div>
  )
}

/** Data ▸ Filter (Ctrl+Shift+L): on over the current region, or off. */
export async function toggleFilter() {
  const c = A.ctx()
  if (!c) return
  const meta = c.book.sheetMeta(c.view.sheet)
  const has = meta?.filters.some((f) => f.key === 'sheet')
  const { r, c: col } = c.view.sel.active
  const inTable = c.book.meta.tables.find((t) => t.sheet === c.view.sheet && r >= t.area.r0 && r <= t.area.r1 && col >= t.area.c0 && col <= t.area.c1)
  if (inTable && !has) {
    await A.edit(c, (b) => b.api.setTableOptions(inTable.name, { autoFilter: !inTable.autoFilter }))
    return
  }
  let area: Area | null = null
  if (!has) {
    const sel = c.view.sel.ranges[c.view.sel.ranges.length - 1]
    const reg = sel.r0 === sel.r1 && sel.c0 === sel.c1 ? await currentRegion(c) : sel
    area = { sheet: c.view.sheet, ...reg }
  }
  await A.edit(c, (b) => b.api.toggleAutoFilter(c.view.sheet, area))
}

/** The table at the active cell (for Table Design), or null. */
export function useActiveTable(tab: string, book: Book | undefined) {
  const view = useViews((s) => s.views[tab])
  if (!book || !view) return null
  const { r, c } = view.sel.active
  return book.meta.tables.find((t) => t.sheet === view.sheet && r >= t.area.r0 && r <= t.area.r1 && c >= t.area.c0 && c <= t.area.c1) ?? null
}
