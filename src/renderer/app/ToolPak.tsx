// Data ▸ Data Analysis: the list of Analysis ToolPak tools, and one form for each built from its fields.
import { useEffect, useState } from 'react'
import { areaName, cellName, parseArea, parseCell } from '@/core/refs'
import type { ToolId, ToolParams } from '@/core/toolpak'
import { set, toast } from '../state/store'
import { setView } from '../state/view'
import { Btn, Field, Modal, cx, inputCls } from '../ui'
import * as A from './actions'
import { closeDialog, focusGrid } from './focus'
import { currentRegion } from './commands'

const done = () => (closeDialog(), focusGrid())

type Kind = 'range' | 'range?' | 'number' | 'check' | 'select'
interface FieldDef {
  key: string
  label: string
  kind: Kind
  value?: string | boolean
  options?: [string, string][]
  /** Shown only when another field has this value. */
  when?: [string, string]
}

const input: FieldDef = { key: 'input', label: 'Input range', kind: 'range' }
const byRows: FieldDef = { key: 'byRows', label: 'Grouped by', kind: 'select', value: 'columns', options: [['columns', 'Columns'], ['rows', 'Rows']] }
const labels: FieldDef = { key: 'labels', label: 'Labels in first row', kind: 'check', value: true }
const alpha: FieldDef = { key: 'alpha', label: 'Alpha', kind: 'number', value: '0.05' }
const var12: FieldDef[] = [
  { key: 'input', label: 'Variable 1 range', kind: 'range' },
  { key: 'input2', label: 'Variable 2 range', kind: 'range' }
]
const diff: FieldDef = { key: 'n', label: 'Hypothesized mean difference', kind: 'number', value: '0' }

export const TOOLS: { id: ToolId | 'descriptive'; title: string; fields: FieldDef[] }[] = [
  { id: 'anova1', title: 'Anova: Single Factor', fields: [input, byRows, labels, alpha] },
  { id: 'anova2r', title: 'Anova: Two-Factor With Replication', fields: [input, { key: 'n', label: 'Rows per sample', kind: 'number', value: '2' }, alpha] },
  { id: 'anova2', title: 'Anova: Two-Factor Without Replication', fields: [input, { ...labels, label: 'Labels' }, alpha] },
  { id: 'correlation', title: 'Correlation', fields: [input, byRows, labels] },
  { id: 'covariance', title: 'Covariance', fields: [input, byRows, labels] },
  { id: 'descriptive', title: 'Descriptive Statistics', fields: [] },
  { id: 'expsmooth', title: 'Exponential Smoothing', fields: [input, { key: 'n', label: 'Damping factor', kind: 'number', value: '0.3' }, labels, { key: 'stdErrors', label: 'Standard errors', kind: 'check' }] },
  { id: 'ftest', title: 'F-Test Two-Sample for Variances', fields: [...var12, labels, alpha] },
  { id: 'fourier', title: 'Fourier Analysis', fields: [input, labels, { key: 'inverse', label: 'Inverse', kind: 'check' }] },
  { id: 'histogram', title: 'Histogram', fields: [input, { key: 'extra', label: 'Bin range (optional)', kind: 'range?' }, labels, { key: 'pareto', label: 'Pareto (sorted histogram)', kind: 'check' }, { key: 'cumulative', label: 'Cumulative percentage', kind: 'check' }] },
  { id: 'movavg', title: 'Moving Average', fields: [input, labels, { key: 'n', label: 'Interval', kind: 'number', value: '3' }, { key: 'stdErrors', label: 'Standard errors', kind: 'check' }] },
  {
    id: 'random',
    title: 'Random Number Generation',
    fields: [
      { key: 'variables', label: 'Number of variables', kind: 'number', value: '1' },
      { key: 'count', label: 'Number of random numbers', kind: 'number', value: '10' },
      { key: 'distribution', label: 'Distribution', kind: 'select', value: 'uniform', options: [['uniform', 'Uniform'], ['normal', 'Normal'], ['bernoulli', 'Bernoulli'], ['binomial', 'Binomial'], ['poisson', 'Poisson']] },
      { key: 'p1', label: 'Between', kind: 'number', value: '0', when: ['distribution', 'uniform'] },
      { key: 'p2', label: 'and', kind: 'number', value: '1', when: ['distribution', 'uniform'] },
      { key: 'p1', label: 'Mean', kind: 'number', value: '0', when: ['distribution', 'normal'] },
      { key: 'p2', label: 'Standard deviation', kind: 'number', value: '1', when: ['distribution', 'normal'] },
      { key: 'p1', label: 'p value', kind: 'number', value: '0.5', when: ['distribution', 'bernoulli'] },
      { key: 'p1', label: 'p value', kind: 'number', value: '0.5', when: ['distribution', 'binomial'] },
      { key: 'p2', label: 'Number of trials', kind: 'number', value: '10', when: ['distribution', 'binomial'] },
      { key: 'p1', label: 'Lambda', kind: 'number', value: '1', when: ['distribution', 'poisson'] },
      { key: 'seed', label: 'Random seed (optional)', kind: 'number', value: '' }
    ]
  },
  { id: 'rank', title: 'Rank and Percentile', fields: [input, byRows, labels] },
  {
    id: 'regression',
    title: 'Regression',
    fields: [
      { key: 'input', label: 'Input Y range', kind: 'range' },
      { key: 'extra', label: 'Input X range', kind: 'range' },
      { ...labels, label: 'Labels' },
      { key: 'zeroConstant', label: 'Constant is zero', kind: 'check' },
      { key: 'confidence', label: 'Confidence level (%)', kind: 'number', value: '95' },
      { key: 'residuals', label: 'Residuals', kind: 'check' }
    ]
  },
  {
    id: 'sampling',
    title: 'Sampling',
    fields: [
      input,
      { ...labels, label: 'Labels', value: false },
      { key: 'method', label: 'Sampling method', kind: 'select', value: 'periodic', options: [['periodic', 'Periodic'], ['random', 'Random']] },
      { key: 'n', label: 'Period', kind: 'number', value: '2', when: ['method', 'periodic'] },
      { key: 'n', label: 'Number of samples', kind: 'number', value: '5', when: ['method', 'random'] }
    ]
  },
  { id: 'ttest-paired', title: 't-Test: Paired Two Sample for Means', fields: [...var12, diff, labels, alpha] },
  { id: 'ttest-equal', title: 't-Test: Two-Sample Assuming Equal Variances', fields: [...var12, diff, labels, alpha] },
  { id: 'ttest-unequal', title: 't-Test: Two-Sample Assuming Unequal Variances', fields: [...var12, diff, labels, alpha] },
  { id: 'ztest', title: 'z-Test: Two Sample for Means', fields: [...var12, diff, { key: 'var1', label: 'Variable 1 variance (known)', kind: 'number', value: '' }, { key: 'var2', label: 'Variable 2 variance (known)', kind: 'number', value: '' }, labels, alpha] }
]

export function DataAnalysis() {
  const [pick, setPick] = useState(0)
  const [tool, setTool] = useState<(typeof TOOLS)[number] | null>(null)
  if (tool) return <ToolForm tool={tool} onBack={() => setTool(null)} />
  const go = () => {
    const t = TOOLS[pick]
    if (t.id === 'descriptive') return set({ dialog: { kind: 'tool', tool: 'descriptive' } })
    setTool(t)
  }
  return (
    <Modal open onOpenChange={(o) => !o && done()} title="Data analysis" width={420} footer={<><Btn onClick={done}>Cancel</Btn><Btn variant="primary" onClick={go}>OK</Btn></>}>
      <span id="tool-list" className="mb-1.5 block text-[12px] font-medium text-text-2">
        Analysis tools
      </span>
      <select aria-labelledby="tool-list" autoFocus size={12} className={cx(inputCls, 'h-auto w-full py-1 text-[12.5px]')} value={String(pick)} onChange={(e) => setPick(Number(e.target.value))} onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), go())} onDoubleClick={go}>
        {TOOLS.map((t, i) => (
          <option key={t.id} value={i}>
            {t.title}
          </option>
        ))}
      </select>
    </Modal>
  )
}

function ToolForm({ tool, onBack }: { tool: (typeof TOOLS)[number]; onBack: () => void }) {
  const c = A.ctx()
  const [v, setV] = useState<Record<string, string | boolean>>(() => Object.fromEntries(tool.fields.map((f) => [f.key, f.value ?? (f.kind === 'check' ? false : '')])))
  const [where, setWhere] = useState<'sheet' | 'range'>('sheet')
  const [outCell, setOutCell] = useState('')
  const [outSheet, setOutSheet] = useState(tool.title.split(':')[0])
  const [error, setError] = useState('')
  useEffect(() => {
    if (!c) return
    const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
    const fill = (a: { r0: number; c0: number; r1: number; c1: number }) => {
      const t = areaName(a).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')
      setV((x) => ({ ...x, input: x.input || (tool.fields.some((f) => f.key === 'input') ? t : '') }))
      setOutCell(cellName(a.r0, a.c1 + 2))
    }
    if (r.r0 === r.r1 && r.c0 === r.c1) void currentRegion(c).then((a) => a && fill(a))
    else fill(r)
  }, [])
  if (!c) return null
  const shown = tool.fields.filter((f) => !f.when || v[f.when[0]] === f.when[1])
  const ok = async () => {
    const p: Record<string, unknown> = {}
    for (const f of shown) {
      const raw = v[f.key]
      if (f.kind === 'check') p[f.key] = raw === true
      else if (f.kind === 'select') p[f.key] = f.key === 'byRows' ? raw === 'rows' : raw
      else if (f.kind === 'number') {
        if (raw === '') continue
        const n = Number(raw)
        if (!Number.isFinite(n)) return setError(`${f.label} must be a number.`)
        p[f.key] = f.key === 'confidence' ? n / 100 : n
      } else {
        const t = String(raw).trim()
        if (!t) {
          if (f.kind === 'range?') continue
          return setError(`Enter the ${f.label.toLowerCase()}.`)
        }
        const a = parseArea(t.replace(/^.*!/, ''), c.view.sheet)
        if (!a) return setError(`${f.label}: “${t}” isn’t a range.`)
        p[f.key] = a
      }
    }
    if (where === 'range') {
      const o = parseCell(outCell.replace(/\$/g, ''))
      if (!o) return setError('Enter the output range’s top-left cell.')
      p.out = { sheet: c.view.sheet, ...o }
    } else p.out = { newSheet: outSheet }
    const r = await c.book.api.runTool(tool.id as ToolId, p as unknown as ToolParams)
    if (!r.ok) return setError(r.error)
    A.markDirtyTab(c.tab)
    await c.book.refreshMeta()
    if (r.sheet !== c.view.sheet) setView(c.tab, { sheet: r.sheet })
    done()
    toast(`${tool.title}: done.`, 'success')
  }
  return (
    <Modal open onOpenChange={(o) => !o && done()} title={tool.title} width={440} footer={<><Btn onClick={onBack}>Back</Btn><Btn onClick={done}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
      <div className="space-y-2.5 text-[12.5px]">
        {shown.map((f, i) =>
          f.kind === 'check' ? (
            <label key={f.key + f.label} className="flex items-center gap-2">
              <input type="checkbox" checked={v[f.key] === true} onChange={(e) => setV({ ...v, [f.key]: e.target.checked })} /> {f.label}
            </label>
          ) : f.kind === 'select' ? (
            <Field key={f.key + f.label} label={f.label}>
              <select aria-label={f.label} className={inputCls} value={String(v[f.key])} onChange={(e) => setV({ ...v, [f.key]: e.target.value })}>
                {f.options!.map(([val, lab]) => (
                  <option key={val} value={val}>
                    {lab}
                  </option>
                ))}
              </select>
            </Field>
          ) : (
            <Field key={f.key + f.label} label={f.label}>
              <input aria-label={f.label} autoFocus={i === 0} className={inputCls} value={String(v[f.key])} onChange={(e) => setV({ ...v, [f.key]: e.target.value })} />
            </Field>
          )
        )}
        <fieldset className="space-y-1.5 rounded-lg border border-border p-2">
          <legend className="px-1 text-[12px] font-medium text-text-2">Output</legend>
          <label className="flex items-center gap-2">
            <input type="radio" name="tool-out" checked={where === 'sheet'} onChange={() => setWhere('sheet')} /> New worksheet
            <input aria-label="New worksheet name" className={cx(inputCls, 'ml-auto w-44')} disabled={where !== 'sheet'} value={outSheet} onChange={(e) => setOutSheet(e.target.value)} />
          </label>
          <label className="flex items-center gap-2">
            <input type="radio" name="tool-out" checked={where === 'range'} onChange={() => setWhere('range')} /> Output range
            <input aria-label="Output range" className={cx(inputCls, 'ml-auto w-44')} disabled={where !== 'range'} value={outCell} onChange={(e) => setOutCell(e.target.value)} />
          </label>
        </fieldset>
        {error && (
          <p role="alert" className="text-[12px] text-danger">
            {error}
          </p>
        )}
      </div>
    </Modal>
  )
}
