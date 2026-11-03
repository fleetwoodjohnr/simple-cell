// PivotTables in the UI: Insert ▸ PivotTable, and the PivotTable Fields panel beside the grid.
import { useEffect, useState } from 'react'
import { ArrowDown, ArrowUp, RefreshCw, Table2, X } from 'lucide-react'
import type { Agg, PivotConfig, PivotValue } from '@/core/pivot'
import { AGG_LABEL, valueName } from '@/core/pivot'
import { areaName, parseArea } from '@/core/refs'
import type { Book } from '../state/client'
import { set, toast, type Dialog } from '../state/store'
import { setView, useViews } from '../state/view'
import { Btn, Field, Modal, Segmented, inputCls } from '../ui'
import * as A from './actions'
import { currentRegion } from './commands'
import { useBookVersion } from './hooks'

const close = () => set({ dialog: null })

export async function openInsertPivot() {
  const c = A.ctx()
  if (!c) return
  const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  const area = r.r0 === r.r1 && r.c0 === r.c1 ? await currentRegion(c) : r
  set({ dialog: { kind: 'pivot-insert', area } })
}

export function InsertPivotDialog({ d }: { d: Extract<Dialog, { kind: 'pivot-insert' }> }) {
  const c = A.ctx()
  const sheetName = c?.book.sheetMeta(c.view.sheet)?.name ?? ''
  const [src, setSrc] = useState(`${sheetName}!${areaName(d.area).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`)
  const [where, setWhere] = useState<'new' | 'existing'>('new')
  const [dest, setDest] = useState('')
  const [error, setError] = useState('')
  if (!c) return null
  const ok = async () => {
    const ref = src.replace(/^.*!/, '').replace(/\$/g, '')
    const a = parseArea(ref, c.view.sheet)
    if (!a) return setError('Enter the data range, such as Sheet1!$A$1:$D$100.')
    let target: { sheet: number; r: number; c: number } | null = null
    if (where === 'existing') {
      const p = parseArea(dest.replace(/^.*!/, '').replace(/\$/g, ''), c.view.sheet)
      if (!p) return setError('Enter the cell where the PivotTable should start, such as H2.')
      target = { sheet: c.view.sheet, r: p.r0, c: p.c0 }
    }
    const res = await c.book.api.createPivot(c.view.sheet, a, target)
    if (!res.ok) return setError(res.error)
    A.markDirtyTab(c.tab)
    close()
    setView(c.tab, { sheet: res.sheet, sel: { ranges: [{ r0: target?.r ?? 2, c0: target?.c ?? 0, r1: target?.r ?? 2, c1: target?.c ?? 0 }], active: { r: target?.r ?? 2, c: target?.c ?? 0 }, focus: { r: target?.r ?? 2, c: target?.c ?? 0 } } })
  }
  return (
    <Modal open onOpenChange={(o) => !o && close()} title="Create PivotTable" width={420} footer={<><Btn onClick={close}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
      <div className="space-y-3 text-[12.5px]">
        <Field label="Table or range">
          <input aria-label="PivotTable source" autoFocus className={inputCls} value={src} onChange={(e) => setSrc(e.target.value)} />
        </Field>
        <Field label="Put the PivotTable">
          <Segmented label="Put the PivotTable" value={where} onChange={setWhere} options={[{ value: 'new', label: 'New worksheet' }, { value: 'existing', label: 'Existing worksheet' }]} />
        </Field>
        {where === 'existing' && (
          <Field label="Location">
            <input aria-label="PivotTable location" className={inputCls} value={dest} placeholder="H2" onChange={(e) => setDest(e.target.value)} />
          </Field>
        )}
        {error && <p role="alert" className="text-[12px] text-danger">{error}</p>}
      </div>
    </Modal>
  )
}

type Info = NonNullable<Awaited<ReturnType<Book['api']['pivotAt']>>>
type Area = 'filters' | 'cols' | 'rows' | 'values'

/** The PivotTable at the active cell, refreshed as the workbook changes. */
export function usePivotAt(tab: string, book: Book | undefined) {
  const view = useViews((s) => s.views[tab])
  const [info, setInfo] = useState<Info | null>(null)
  const v = useBookVersion(book)
  const r = view?.sel.active.r
  const c = view?.sel.active.c
  const inArea = !!book?.sheetMeta(view?.sheet ?? -1)?.pivots.some((p) => r! >= p.area.r0 && r! <= p.area.r1 && c! >= p.area.c0 && c! <= p.area.c1)
  useEffect(() => {
    let live = true
    if (book && view && inArea) void book.api.pivotAt(view.sheet, r!, c!).then((x) => live && setInfo(x))
    else setInfo(null)
    return () => {
      live = false
    }
  }, [book, view?.sheet, r, c, v, inArea])
  return info
}

/** PivotTable Fields: tick fields, arrange them in Filters / Columns / Rows / Values, choose how values summarize. */
export function PivotPanel({ tab, book }: { tab: string; book: Book }) {
  const view = useViews((s) => s.views[tab])
  const info = usePivotAt(tab, book)
  if (!info || !view) return null
  const cfg = info.config
  const apply = (next: PivotConfig) => {
    const c = A.ctx()
    if (c) void A.edit(c, (b) => b.api.setPivotConfig(view.sheet, info.id, next))
  }
  const inUse = (f: string) => cfg.rows.includes(f) || cfg.cols.includes(f) || cfg.filters.includes(f) || cfg.values.some((v) => v.field === f)
  const toggle = (f: string, i: number) => {
    if (inUse(f)) return apply({ ...cfg, rows: cfg.rows.filter((x) => x !== f), cols: cfg.cols.filter((x) => x !== f), filters: cfg.filters.filter((x) => x !== f), values: cfg.values.filter((v) => v.field !== f) })
    // As Excel: numbers go to Values (Sum), other fields to Rows.
    if (info.numeric[i] && !info.dates[i]) apply({ ...cfg, values: [...cfg.values, { field: f, agg: 'sum' }] })
    else apply({ ...cfg, rows: [...cfg.rows, f] })
  }
  const list = (a: Area): string[] => (a === 'values' ? cfg.values.map((v) => v.field) : cfg[a])
  const without = (a: Area, i: number): PivotConfig => (a === 'values' ? { ...cfg, values: cfg.values.filter((_, k) => k !== i) } : { ...cfg, [a]: cfg[a].filter((_, k) => k !== i) })
  const moveTo = (from: Area, i: number, to: Area) => {
    const f = list(from)[i]
    const base = without(from, i)
    if (to === 'values') apply({ ...base, values: [...base.values, { field: f, agg: info.numeric[info.fields.indexOf(f)] ? 'sum' : 'count' }] })
    else apply({ ...base, [to]: [...base[to].filter((x) => x !== f), f] })
  }
  const reorder = (a: Area, i: number, d: -1 | 1) => {
    const arr = a === 'values' ? [...cfg.values] : [...cfg[a]]
    const j = i + d
    if (j < 0 || j >= arr.length) return
    ;[arr[i], arr[j]] = [arr[j], arr[i]]
    apply(a === 'values' ? { ...cfg, values: arr as PivotValue[] } : { ...cfg, [a]: arr })
  }
  const setValue = (i: number, patch: Partial<PivotValue>) => apply({ ...cfg, values: cfg.values.map((v, k) => (k === i ? { ...v, ...patch } : v)) })
  const AREAS: [Area, string][] = [
    ['filters', 'Filters'],
    ['cols', 'Columns'],
    ['rows', 'Rows'],
    ['values', 'Values']
  ]
  return (
    <aside aria-label="PivotTable fields" data-testid="pivot-panel" className="flex w-[280px] shrink-0 flex-col overflow-y-auto border-l border-border bg-panel p-3 text-[12.5px]">
      <div className="mb-2 flex items-center gap-2">
        <Table2 className="size-4 text-accent" />
        <span className="flex-1 font-semibold">{info.name}</span>
        <button type="button" aria-label="Refresh PivotTable" className="rounded p-1 text-text-2 hover:bg-hover" onClick={() => A.ctx() && void A.edit(A.ctx()!, (b) => b.api.refreshPivots(view.sheet, info.id))}>
          <RefreshCw className="size-3.5" />
        </button>
      </div>
      <div className="mb-1 text-[11px] font-semibold text-text-3">Fields</div>
      <div className="mb-3 max-h-[200px] overflow-y-auto rounded border border-border p-1">
        {info.fields.map((f, i) => (
          <label key={f} className="flex items-center gap-2 rounded px-1 py-0.5 hover:bg-hover">
            <input type="checkbox" checked={inUse(f)} onChange={() => toggle(f, i)} /> {f}
          </label>
        ))}
      </div>
      {AREAS.map(([a, label]) => (
        <div key={a} className="mb-3">
          <div className="mb-1 text-[11px] font-semibold text-text-3">{label}</div>
          <div className="min-h-[30px] space-y-1 rounded border border-dashed border-border p-1" data-testid={`pivot-area-${a}`}>
            {list(a).map((f, i) => (
              <div key={`${f}${i}`} className="rounded bg-hover px-1.5 py-1">
                <div className="flex items-center gap-1">
                  <span className="flex-1 truncate">{a === 'values' ? valueName(cfg.values[i]) : f}</span>
                  <select aria-label={`Move ${f}`} className="h-6 rounded border border-border bg-panel text-[11px]" value={a} onChange={(e) => moveTo(a, i, e.target.value as Area)}>
                    {AREAS.map(([v, l]) => (
                      <option key={v} value={v}>
                        {l}
                      </option>
                    ))}
                  </select>
                  <button type="button" aria-label={`Move ${f} up`} className="rounded p-0.5 hover:bg-panel" onClick={() => reorder(a, i, -1)}>
                    <ArrowUp className="size-3" />
                  </button>
                  <button type="button" aria-label={`Move ${f} down`} className="rounded p-0.5 hover:bg-panel" onClick={() => reorder(a, i, 1)}>
                    <ArrowDown className="size-3" />
                  </button>
                  <button type="button" aria-label={`Remove ${f}`} className="rounded p-0.5 hover:bg-panel" onClick={() => apply(without(a, i))}>
                    <X className="size-3" />
                  </button>
                </div>
                {a === 'values' && (
                  <div className="mt-1 flex gap-1">
                    <select aria-label={`Summarize ${f} by`} className="h-6 flex-1 rounded border border-border bg-panel text-[11px]" value={cfg.values[i].agg} onChange={(e) => setValue(i, { agg: e.target.value as Agg, name: undefined })}>
                      {(Object.keys(AGG_LABEL) as Agg[]).map((g) => (
                        <option key={g} value={g}>
                          {AGG_LABEL[g]}
                        </option>
                      ))}
                    </select>
                    <select aria-label={`Show ${f} as`} className="h-6 flex-1 rounded border border-border bg-panel text-[11px]" value={cfg.values[i].showAs ?? 'normal'} onChange={(e) => setValue(i, { showAs: e.target.value as PivotValue['showAs'] })}>
                      <option value="normal">No calculation</option>
                      <option value="percentOfTotal">% of grand total</option>
                      <option value="percentOfRow">% of row total</option>
                      <option value="percentOfCol">% of column total</option>
                    </select>
                  </div>
                )}
                {(a === 'rows' || a === 'cols') && info.dates[info.fields.indexOf(f)] && (
                  <select aria-label={`Group ${f} by`} className="mt-1 h-6 w-full rounded border border-border bg-panel text-[11px]" value={cfg.dateGroups?.[f] ?? ''} onChange={(e) => apply({ ...cfg, dateGroups: { ...cfg.dateGroups, [f]: (e.target.value || undefined) as 'years' } })}>
                    <option value="">Dates (no grouping)</option>
                    <option value="years">Years</option>
                    <option value="quarters">Quarters</option>
                    <option value="months">Months</option>
                  </select>
                )}
                {a === 'filters' && (
                  <select
                    aria-label={`Filter ${f}`}
                    className="mt-1 h-6 w-full rounded border border-border bg-panel text-[11px]"
                    value={cfg.include?.[f]?.length === 1 ? cfg.include[f][0] : ''}
                    onChange={(e) => {
                      const inc = { ...cfg.include }
                      if (e.target.value) inc[f] = [e.target.value]
                      else delete inc[f]
                      apply({ ...cfg, include: inc })
                    }}
                  >
                    <option value="">(All)</option>
                    {(info.items[f] ?? []).map((it) => (
                      <option key={it} value={it}>
                        {it}
                      </option>
                    ))}
                  </select>
                )}
              </div>
            ))}
          </div>
        </div>
      ))}
      <div className="mt-auto flex flex-col gap-1 border-t border-border pt-2">
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={cfg.grandRows !== false} onChange={(e) => apply({ ...cfg, grandRows: e.target.checked })} /> Grand total row
        </label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={cfg.grandCols !== false} onChange={(e) => apply({ ...cfg, grandCols: e.target.checked })} /> Grand total column
        </label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={cfg.subtotals !== false} onChange={(e) => apply({ ...cfg, subtotals: e.target.checked })} /> Subtotals
        </label>
        <Btn size="sm" className="mt-1" onClick={() => A.ctx() && void A.edit(A.ctx()!, (b) => b.api.deletePivot(view.sheet, info.id))}>
          Delete PivotTable
        </Btn>
      </div>
    </aside>
  )
}

export function refreshAll() {
  const c = A.ctx()
  if (!c) return
  void A.edit(c, (b) => b.api.refreshPivots()).then((ok) => ok && toast('Refreshed.', 'success'))
}

