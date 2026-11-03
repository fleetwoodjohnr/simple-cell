// Charts in the UI: ECharts drawing, the overlay that places charts on the grid (select, move, resize), the editor.
import { useEffect, useRef, useState } from 'react'
import * as echarts from 'echarts/core'
import { BarChart, LineChart, PieChart, ScatterChart } from 'echarts/charts'
import { GridComponent, LegendComponent, TitleComponent, TooltipComponent } from 'echarts/components'
import { CanvasRenderer, SVGRenderer } from 'echarts/renderers'
import { ChartArea as AreaIcon, ChartBar, ChartColumn, ChartLine, ChartPie, ChartScatter, Donut, Plus, Trash2 } from 'lucide-react'
import type { Chart, ChartAnchor, ChartData, ChartSeries, ChartType } from '@/core/charts'
import { colX, rowY, type Layout } from '@/grid/layout'
import type { Book } from '../state/client'
import { get, set, toast, type Dialog } from '../state/store'
import { getView, setView, useViews } from '../state/view'
import { Btn, Field, Modal, cx, inputCls } from '../ui'
import * as A from './actions'

echarts.use([BarChart, LineChart, PieChart, ScatterChart, GridComponent, LegendComponent, TitleComponent, TooltipComponent, CanvasRenderer, SVGRenderer])

/** Office theme series colours (accents, then darker accents). */
export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300', '#255E91', '#43682B']

type ChartShape = Omit<Chart, 'carry'>

/** ECharts option for a chart and its data. */
export function chartOption(c: ChartShape, d: ChartData, opts: { dark?: boolean; animate?: boolean } = {}) {
  const text = opts.dark ? '#d6d6dc' : '#404040'
  const line = opts.dark ? '#3a3a40' : '#d9d9d9'
  const title = c.title ?? (d.series.length === 1 ? d.series[0].name : undefined)
  const legend = c.legend === 'none' ? { show: false } : { show: true, type: 'scroll' as const, textStyle: { color: text }, ...(c.legend === 'right' ? { right: 8, top: 'middle', orient: 'vertical' as const } : c.legend === 'left' ? { left: 8, top: 'middle', orient: 'vertical' as const } : c.legend === 'top' ? { top: title ? 30 : 8 } : { bottom: 6 }) }
  const base = {
    animation: !!opts.animate,
    color: d.series.map((s, i) => s.color ?? PALETTE[i % PALETTE.length]),
    textStyle: { fontFamily: "Calibri, Carlito, 'Inter Variable', sans-serif" },
    title: title ? { text: title, left: 'center', top: 6, textStyle: { fontSize: 14, fontWeight: 'normal' as const, color: text } } : undefined,
    legend,
    tooltip: { trigger: c.type === 'pie' || c.type === 'doughnut' || c.type === 'scatter' ? ('item' as const) : ('axis' as const) }
  }
  if (c.type === 'pie' || c.type === 'doughnut') {
    const s = d.series[0]
    return {
      ...base,
      color: PALETTE,
      series: [
        {
          type: 'pie' as const,
          radius: c.type === 'doughnut' ? ['45%', '70%'] : '70%',
          center: [c.legend === 'right' ? '42%' : c.legend === 'left' ? '58%' : '50%', '55%'],
          data: d.categories.map((name, i) => ({ name, value: s?.values[i] ?? null })),
          label: { show: !!c.dataLabels, color: text },
          labelLine: { show: !!c.dataLabels }
        }
      ]
    }
  }
  const pad = { left: 48, right: c.legend === 'right' ? 120 : 20, top: title ? 40 : 18, bottom: c.legend === 'bottom' ? 56 : 32, containLabel: true }
  const valueAxis = { type: 'value' as const, name: c.yTitle, nameLocation: 'middle' as const, nameGap: 40, axisLine: { show: false }, splitLine: { lineStyle: { color: line } }, axisLabel: { color: text, formatter: c.grouping === 'percent' ? '{value}%' : undefined }, max: c.grouping === 'percent' ? 100 : undefined }
  if (c.type === 'scatter') {
    return {
      ...base,
      grid: pad,
      xAxis: { ...valueAxis, name: c.xTitle, nameGap: 28, splitLine: { show: false } },
      yAxis: valueAxis,
      series: d.series.map((s) => ({ type: 'scatter' as const, name: s.name, symbolSize: 7, data: s.values.map((v, i) => [s.x?.[i] ?? i + 1, v]) }))
    }
  }
  const catAxis = { type: 'category' as const, data: d.categories, name: c.xTitle, nameLocation: 'middle' as const, nameGap: 28, axisLine: { lineStyle: { color: line } }, axisTick: { show: false }, axisLabel: { color: text } }
  const horizontal = c.type === 'bar'
  const stack = c.grouping !== 'clustered' ? 'total' : undefined
  // 100% stacked: each category's values as shares of its total.
  const totals = d.categories.map((_, i) => d.series.reduce((t, s) => t + Math.abs(s.values[i] ?? 0), 0))
  const values = (s: ChartData['series'][number]) => (c.grouping === 'percent' ? s.values.map((v, i) => (v === null || !totals[i] ? null : (v / totals[i]) * 100)) : s.values)
  return {
    ...base,
    grid: pad,
    xAxis: horizontal ? valueAxis : catAxis,
    yAxis: horizontal ? { ...catAxis, inverse: true } : valueAxis,
    series: d.series.map((s) => ({
      type: c.type === 'line' || c.type === 'area' ? ('line' as const) : ('bar' as const),
      name: s.name,
      data: values(s),
      stack,
      areaStyle: c.type === 'area' ? { opacity: 0.85 } : undefined,
      symbol: c.type === 'line' ? 'circle' : 'none',
      symbolSize: 5,
      barMaxWidth: 48,
      label: { show: !!c.dataLabels, position: c.type === 'bar' ? ('right' as const) : ('top' as const), color: text }
    }))
  }
}

/** Render a chart to SVG text (printing, PDF). */
export function chartSvg(c: ChartShape, d: ChartData, width: number, height: number): string {
  const chart = echarts.init(null, null, { renderer: 'svg', ssr: true, width, height })
  chart.setOption(chartOption(c, d) as never)
  const svg = chart.renderToSVGString()
  chart.dispose()
  return svg
}

export function ChartView({ chart, data, width, height, dark }: { chart: ChartShape; data: ChartData | null; width: number; height: number; dark?: boolean }) {
  const ref = useRef<HTMLDivElement>(null)
  const inst = useRef<echarts.ECharts | null>(null)
  useEffect(() => {
    if (!ref.current) return
    inst.current = echarts.init(ref.current, null, { renderer: 'canvas' })
    return () => {
      inst.current?.dispose()
      inst.current = null
    }
  }, [])
  useEffect(() => {
    inst.current?.resize({ width, height })
  }, [width, height])
  useEffect(() => {
    if (inst.current && data) inst.current.setOption(chartOption(chart, data, { dark }) as never, true)
  }, [chart, data, dark])
  return <div ref={ref} style={{ width, height }} />
}

/** Screen point → anchor corner (cell + unzoomed offset). */
function pointToCorner(L: Layout, x: number, y: number) {
  const ux = Math.max(0, (x - L.headerW) / L.zoom)
  const uy = Math.max(0, (y - L.headerH) / L.zoom)
  const ax = ux < L.frozenW ? ux : ux + L.sx
  const ay = uy < L.frozenH ? uy : uy + L.sy
  const c = L.cols.indexAt(ax)
  const r = L.rows.indexAt(ay)
  return { r, c, dx: Math.round(ax - L.cols.offset(c)), dy: Math.round(ay - L.rows.offset(r)) }
}

const rectOf = (L: Layout, a: ChartAnchor) => {
  const x0 = colX(L, a.c0) + a.dx0 * L.zoom
  const y0 = rowY(L, a.r0) + a.dy0 * L.zoom
  const x1 = colX(L, a.c1) + a.dx1 * L.zoom
  const y1 = rowY(L, a.r1) + a.dy1 * L.zoom
  return { x: x0, y: y0, w: Math.max(40, x1 - x0), h: Math.max(30, y1 - y0) }
}

export function ChartsOverlay({ tab, book, L }: { tab: string; book: Book; L: Layout }) {
  const view = useViews((s) => s.views[tab])
  const meta = book.sheetMeta(view.sheet)
  const charts = meta?.charts ?? []
  const [data, setData] = useState<Record<number, ChartData | null>>({})
  const [drag, setDrag] = useState<{ id: number; kind: 'move' | 'size'; x: number; y: number; dx: number; dy: number } | null>(null)
  const dark = document.documentElement.classList.contains('dark')
  // Chart data follows the workbook.
  useEffect(() => {
    let live = true
    void Promise.all(charts.map((c) => book.api.chartData(view.sheet, c.id).then((d) => [c.id, d] as const))).then((list) => live && setData(Object.fromEntries(list)))
    return () => {
      live = false
    }
  }, [book.version, view.sheet, charts.length])
  if (!charts.length) return null
  const selected = view.chart
  const finish = async (c: ChartShape, r: { x: number; y: number; w: number; h: number }) => {
    const a = pointToCorner(L, r.x, r.y)
    const b = pointToCorner(L, r.x + r.w, r.y + r.h)
    const ctx = A.ctx()
    if (ctx) await A.edit(ctx, (bk) => bk.api.updateChart(view.sheet, c.id, { anchor: { r0: a.r, c0: a.c, dx0: a.dx, dy0: a.dy, r1: b.r, c1: b.c, dx1: b.dx, dy1: b.dy } }))
  }
  return (
    <div className="pointer-events-none absolute overflow-hidden" style={{ left: L.headerW, top: L.headerH, width: L.width - L.headerW, height: L.height - L.headerH }}>
      {charts.map((c) => {
        const base = rectOf(L, c.anchor)
        const r = { ...base, x: base.x - L.headerW, y: base.y - L.headerH }
        if (drag?.id === c.id) {
          if (drag.kind === 'move') {
            r.x += drag.dx
            r.y += drag.dy
          } else {
            r.w = Math.max(80, r.w + drag.dx)
            r.h = Math.max(60, r.h + drag.dy)
          }
        }
        if (r.x > L.width || r.y > L.height || r.x + r.w < 0 || r.y + r.h < 0) return null
        const sel = selected === c.id
        return (
          <div
            key={c.id}
            data-testid={`chart-${c.id}`}
            role="img"
            aria-label={`Chart: ${c.title ?? c.type}`}
            className={cx('pointer-events-auto absolute rounded-sm bg-white shadow-sm', dark && 'bg-[#1b1b1f]', sel ? 'ring-2 ring-accent' : 'ring-1 ring-border')}
            style={{ left: r.x, top: r.y, width: r.w, height: r.h, cursor: sel ? 'move' : 'default' }}
            onPointerDown={(e) => {
              e.stopPropagation()
              setView(tab, { chart: c.id })
              ;(e.currentTarget as HTMLElement).setPointerCapture(e.pointerId)
              setDrag({ id: c.id, kind: (e.target as HTMLElement).dataset.handle ? 'size' : 'move', x: e.clientX, y: e.clientY, dx: 0, dy: 0 })
            }}
            onPointerMove={(e) => drag?.id === c.id && setDrag({ ...drag, dx: e.clientX - drag.x, dy: e.clientY - drag.y })}
            onPointerUp={() => {
              if (drag?.id === c.id && (Math.abs(drag.dx) > 2 || Math.abs(drag.dy) > 2)) void finish(c, { x: r.x + L.headerW, y: r.y + L.headerH, w: r.w, h: r.h })
              setDrag(null)
              requestAnimationFrame(() => document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus())
            }}
            onDoubleClick={() => set({ dialog: { kind: 'chart', id: c.id } })}
          >
            <div className="pointer-events-none">
              <ChartView chart={c} data={data[c.id] ?? null} width={Math.round(r.w)} height={Math.round(r.h)} dark={dark} />
            </div>
            {sel && <span data-handle="size" data-testid="chart-resize" className="absolute -bottom-1.5 -right-1.5 size-3 cursor-nwse-resize rounded-sm border border-accent bg-panel" />}
          </div>
        )
      })}
    </div>
  )
}

/** Insert ▸ Chart (Alt+F1): a chart of the selection, or of the data region around the active cell. */
export async function insertChart(type: ChartType = 'column') {
  const c = A.ctx()
  if (!c) return
  const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  const { currentRegion } = await import('./commands')
  const area = r.r0 === r.r1 && r.c0 === r.c1 ? await currentRegion(c) : r
  const res = await c.book.api.addChartFromRange(c.view.sheet, { sheet: c.view.sheet, ...area }, type)
  if (!res.ok) return toast(res.error, 'info')
  A.markDirtyTab(c.tab)
  setView(c.tab, { chart: res.id })
}

const TYPES: [ChartType, string, typeof ChartColumn][] = [
  ['column', 'Column', ChartColumn],
  ['bar', 'Bar', ChartBar],
  ['line', 'Line', ChartLine],
  ['area', 'Area', AreaIcon],
  ['pie', 'Pie', ChartPie],
  ['doughnut', 'Doughnut', Donut],
  ['scatter', 'Scatter', ChartScatter]
]

export function ChartTypeIcon({ type }: { type: ChartType }) {
  const I = TYPES.find((t) => t[0] === type)?.[2] ?? ChartColumn
  return <I className="size-4" />
}

export { TYPES as CHART_TYPES }

export function ChartDialog({ d }: { d: Extract<Dialog, { kind: 'chart' }> }) {
  const c = A.ctx()
  const meta = c?.book.sheetMeta(c.view.sheet)
  const orig = meta?.charts.find((x) => x.id === d.id)
  const [chart, setChart] = useState<ChartShape | null>(orig ?? null)
  const [error, setError] = useState('')
  if (!c || !chart) return null
  const up = (p: Partial<ChartShape>) => setChart({ ...chart, ...p })
  const upSeries = (i: number, p: Partial<ChartSeries>) => up({ series: chart.series.map((s, k) => (k === i ? { ...s, ...p } : s)) })
  const close = () => set({ dialog: null })
  const ok = async () => {
    const { id: _i, anchor: _a, ...patch } = chart
    void _i
    void _a
    const clean = { ...patch, series: patch.series.map((s) => ({ ...s, name: s.name?.trim() || undefined, cat: s.cat?.trim() || undefined, val: s.val.trim() })) }
    const r = await c.book.api.updateChart(c.view.sheet, chart.id, clean)
    if (!r.ok) return setError(r.error ?? 'That chart can’t be drawn.')
    A.markDirtyTab(c.tab)
    close()
  }
  const stackable = chart.type === 'column' || chart.type === 'bar' || chart.type === 'area' || chart.type === 'line'
  return (
    <Modal
      open
      onOpenChange={(o) => !o && close()}
      title="Chart"
      width={620}
      footer={
        <>
          <Btn className="mr-auto" onClick={() => void A.edit(c, (b) => b.api.deleteChart(c.view.sheet, chart.id)).then(close)}>
            Delete chart
          </Btn>
          <Btn onClick={close}>Cancel</Btn>
          <Btn variant="primary" onClick={() => void ok()}>
            OK
          </Btn>
        </>
      }
    >
      <div className="space-y-3 text-[12.5px]">
        <div role="radiogroup" aria-label="Chart type" className="flex flex-wrap gap-1">
          {TYPES.map(([t, label, I]) => (
            <button key={t} type="button" role="radio" aria-checked={chart.type === t} onClick={() => up({ type: t })} className={cx('flex h-14 w-[76px] flex-col items-center justify-center gap-1 rounded-md border text-[11.5px]', chart.type === t ? 'border-accent bg-accent-soft text-accent-ink' : 'border-border hover:bg-hover')}>
              <I className="size-5" />
              {label}
            </button>
          ))}
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Title">
            <input aria-label="Chart title" className={inputCls} value={chart.title ?? ''} placeholder="Automatic" onChange={(e) => up({ title: e.target.value || undefined })} />
          </Field>
          <Field label="Legend">
            <select aria-label="Legend" className={inputCls} value={chart.legend} onChange={(e) => up({ legend: e.target.value as ChartShape['legend'] })}>
              {['right', 'bottom', 'top', 'left', 'none'].map((v) => (
                <option key={v} value={v}>
                  {v === 'none' ? 'None' : v[0].toUpperCase() + v.slice(1)}
                </option>
              ))}
            </select>
          </Field>
          {stackable && (
            <Field label="Grouping">
              <select aria-label="Grouping" className={inputCls} value={chart.grouping} onChange={(e) => up({ grouping: e.target.value as ChartShape['grouping'] })}>
                <option value="clustered">{chart.type === 'line' || chart.type === 'area' ? 'Standard' : 'Clustered'}</option>
                <option value="stacked">Stacked</option>
                <option value="percent">100% stacked</option>
              </select>
            </Field>
          )}
          <label className="flex items-center gap-2 self-end pb-2">
            <input type="checkbox" checked={!!chart.dataLabels} onChange={(e) => up({ dataLabels: e.target.checked })} /> Data labels
          </label>
          {chart.type !== 'pie' && chart.type !== 'doughnut' && (
            <>
              <Field label="Horizontal axis title">
                <input aria-label="Horizontal axis title" className={inputCls} value={chart.xTitle ?? ''} onChange={(e) => up({ xTitle: e.target.value || undefined })} />
              </Field>
              <Field label="Vertical axis title">
                <input aria-label="Vertical axis title" className={inputCls} value={chart.yTitle ?? ''} onChange={(e) => up({ yTitle: e.target.value || undefined })} />
              </Field>
            </>
          )}
        </div>
        <div>
          <div className="mb-1 flex items-center justify-between">
            <span className="font-medium text-text-2">Series</span>
            <Btn size="sm" onClick={() => up({ series: [...chart.series, { val: chart.series[chart.series.length - 1]?.val ?? '', cat: chart.series[0]?.cat }] })}>
              <Plus className="size-3.5" /> Add series
            </Btn>
          </div>
          <div className="grid grid-cols-[1fr_1.3fr_1.3fr_28px] gap-1.5 text-[11.5px] text-text-3">
            <span>Name</span>
            <span>Values</span>
            <span>{chart.type === 'scatter' ? 'X values' : 'Categories'}</span>
            <span />
            {chart.series.map((s, i) => (
              <div key={i} className="contents">
                <input aria-label={`Series ${i + 1} name`} className={cx(inputCls, 'h-7 font-mono text-[11.5px]')} value={s.name ?? ''} onChange={(e) => upSeries(i, { name: e.target.value })} />
                <input aria-label={`Series ${i + 1} values`} className={cx(inputCls, 'h-7 font-mono text-[11.5px]')} value={s.val} onChange={(e) => upSeries(i, { val: e.target.value })} />
                <input aria-label={`Series ${i + 1} categories`} className={cx(inputCls, 'h-7 font-mono text-[11.5px]')} value={s.cat ?? ''} onChange={(e) => upSeries(i, { cat: e.target.value })} />
                <button type="button" aria-label={`Remove series ${i + 1}`} disabled={chart.series.length < 2} className="rounded p-1 text-text-2 hover:bg-hover disabled:opacity-30" onClick={() => up({ series: chart.series.filter((_, k) => k !== i) })}>
                  <Trash2 className="size-3.5" />
                </button>
              </div>
            ))}
          </div>
        </div>
        {error && (
          <p role="alert" className="text-[12px] text-danger">
            {error}
          </p>
        )}
      </div>
    </Modal>
  )
}

/** Keys while a chart is selected: Delete removes it, Escape returns to the cells. Returns true when handled. */
export function chartKey(e: KeyboardEvent | React.KeyboardEvent, tab: string): boolean {
  const v = getView(tab)
  if (v.chart === undefined || v.editing) return false
  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault()
    const c = A.ctx()
    const id = v.chart
    setView(tab, { chart: undefined })
    if (c) void A.edit(c, (b) => b.api.deleteChart(v.sheet, id))
    return true
  }
  if (e.key === 'Escape') {
    setView(tab, { chart: undefined })
    return true
  }
  if (e.key === 'Enter') {
    set({ dialog: { kind: 'chart', id: v.chart } })
    return true
  }
  void get
  return false
}
