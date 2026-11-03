// Conditional formatting: the ribbon menu (Excel's presets), the rule dialog and Manage Rules.
import { useEffect, useRef, useState } from 'react'
import { ArrowDown, ArrowUp, Trash2 } from 'lucide-react'
import type { CfListItem, CfSpec } from '@/core/doc'
import type { Cfvo } from '@/core/condfmt'
import { ICON_COUNTS } from '@/core/condfmt'
import type { Style } from '@/core/model/styles'
import { cellName } from '@/core/refs'
import { drawIcon } from '@/grid/icons'
import { set, type Dialog } from '../state/store'
import { areasOf } from '../state/view'
import { Btn, Field, MenuItem, MenuSep, Modal, SubMenu, cx, inputCls } from '../ui'
import * as A from './actions'
import { markDirty } from '../state/documents'

export type CfPreset = 'greaterThan' | 'lessThan' | 'between' | 'equal' | 'text' | 'date' | 'duplicate' | 'top' | 'topPercent' | 'bottom' | 'bottomPercent' | 'formula'

const close = () => set({ dialog: null })
const refocus = () => requestAnimationFrame(() => document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus())

const border = { style: 'thin' as const, color: '#9C0006' }
export const FORMAT_PRESETS: { label: string; style: Style }[] = [
  { label: 'Light red fill with dark red text', style: { fill: '#FFC7CE', color: '#9C0006' } },
  { label: 'Yellow fill with dark yellow text', style: { fill: '#FFEB9C', color: '#9C5700' } },
  { label: 'Green fill with dark green text', style: { fill: '#C6EFCE', color: '#006100' } },
  { label: 'Light red fill', style: { fill: '#FFC7CE' } },
  { label: 'Red text', style: { color: '#9C0006' } },
  { label: 'Red border', style: { borderTop: border, borderBottom: border, borderLeft: border, borderRight: border } },
  { label: 'Bold', style: { bold: true } }
]

const COLOR_SCALES: { label: string; colors: string[] }[] = [
  { label: 'Green – yellow – red', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  { label: 'Red – yellow – green', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
  { label: 'Green – white – red', colors: ['#F8696B', '#FCFCFF', '#63BE7B'] },
  { label: 'Red – white – green', colors: ['#63BE7B', '#FCFCFF', '#F8696B'] },
  { label: 'Blue – white – red', colors: ['#F8696B', '#FCFCFF', '#5A8AC6'] },
  { label: 'Red – white – blue', colors: ['#5A8AC6', '#FCFCFF', '#F8696B'] },
  { label: 'White – red', colors: ['#F8696B', '#FCFCFF'] },
  { label: 'Red – white', colors: ['#FCFCFF', '#F8696B'] },
  { label: 'Green – white', colors: ['#FCFCFF', '#63BE7B'] },
  { label: 'White – green', colors: ['#63BE7B', '#FCFCFF'] },
  { label: 'Green – yellow', colors: ['#FFEF9C', '#63BE7B'] },
  { label: 'Yellow – green', colors: ['#63BE7B', '#FFEF9C'] }
]

const DATA_BARS = [
  { label: 'Blue', color: '#638EC6' },
  { label: 'Green', color: '#63C384' },
  { label: 'Red', color: '#FF555A' },
  { label: 'Orange', color: '#FFB628' },
  { label: 'Light blue', color: '#008AEF' },
  { label: 'Purple', color: '#D6007B' }
]

const ICON_SETS: { label: string; set: string }[] = [
  { label: '3 arrows', set: '3Arrows' },
  { label: '3 arrows (gray)', set: '3ArrowsGray' },
  { label: '4 arrows', set: '4Arrows' },
  { label: '5 arrows', set: '5Arrows' },
  { label: '3 triangles', set: '3Triangles' },
  { label: '3 traffic lights', set: '3TrafficLights1' },
  { label: '3 traffic lights (rimmed)', set: '3TrafficLights2' },
  { label: '3 signs', set: '3Signs' },
  { label: '4 traffic lights', set: '4TrafficLights' },
  { label: 'Red to black', set: '4RedToBlack' },
  { label: '3 symbols (circled)', set: '3Symbols' },
  { label: '3 symbols', set: '3Symbols2' },
  { label: '3 flags', set: '3Flags' },
  { label: '3 stars', set: '3Stars' },
  { label: '4 ratings', set: '4Rating' },
  { label: '5 ratings', set: '5Rating' },
  { label: '5 quarters', set: '5Quarters' },
  { label: '5 boxes', set: '5Boxes' }
]

/** Add a rule over the selection. */
async function addRule(spec: CfSpec) {
  const c = A.ctx()
  if (!c) return
  await A.edit(c, (b) => b.api.addCondFormat(c.view.sheet, areasOf(c.view), spec))
}

const percentCfvo = (n: number): Cfvo[] => Array.from({ length: n }, (_, i) => ({ type: 'percent', val: String(Math.round((i * 100) / n)) }))

function IconStrip({ set: name }: { set: string }) {
  const ref = useRef<HTMLCanvasElement>(null)
  const n = ICON_COUNTS[name] ?? 3
  useEffect(() => {
    const cv = ref.current
    const ctx = cv?.getContext('2d')
    if (!cv || !ctx) return
    const dpr = window.devicePixelRatio || 1
    cv.width = n * 16 * dpr
    cv.height = 16 * dpr
    ctx.scale(dpr, dpr)
    for (let i = 0; i < n; i++) drawIcon(ctx, name, n - 1 - i, i * 16 + 1, 1, 14)
  }, [name, n])
  return <canvas ref={ref} className="h-4" style={{ width: n * 16 }} aria-hidden />
}

function Swatch({ colors }: { colors: string[] }) {
  return <span aria-hidden className="inline-block h-4 w-10 rounded-sm border border-border" style={{ background: `linear-gradient(to right, ${colors.join(', ')})` }} />
}

const open = (dialog: Dialog) => set({ dialog })

/** Items of the Home ▸ Conditional Formatting menu. */
export function CondFormatMenuItems() {
  const preset = (p: CfPreset) => open({ kind: 'cf-rule', preset: p })
  return (
    <>
      <SubMenu label="Highlight cells rules">
        <MenuItem onSelect={() => preset('greaterThan')}>Greater than…</MenuItem>
        <MenuItem onSelect={() => preset('lessThan')}>Less than…</MenuItem>
        <MenuItem onSelect={() => preset('between')}>Between…</MenuItem>
        <MenuItem onSelect={() => preset('equal')}>Equal to…</MenuItem>
        <MenuItem onSelect={() => preset('text')}>Text that contains…</MenuItem>
        <MenuItem onSelect={() => preset('date')}>A date occurring…</MenuItem>
        <MenuItem onSelect={() => preset('duplicate')}>Duplicate values…</MenuItem>
      </SubMenu>
      <SubMenu label="Top/bottom rules">
        <MenuItem onSelect={() => preset('top')}>Top 10 items…</MenuItem>
        <MenuItem onSelect={() => preset('topPercent')}>Top 10%…</MenuItem>
        <MenuItem onSelect={() => preset('bottom')}>Bottom 10 items…</MenuItem>
        <MenuItem onSelect={() => preset('bottomPercent')}>Bottom 10%…</MenuItem>
        <MenuItem onSelect={() => void addRule({ type: 'aboveAverage', dxf: FORMAT_PRESETS[0].style })}>Above average</MenuItem>
        <MenuItem onSelect={() => void addRule({ type: 'aboveAverage', aboveAverage: false, dxf: FORMAT_PRESETS[0].style })}>Below average</MenuItem>
      </SubMenu>
      <SubMenu label="Data bars">
        {DATA_BARS.map((d) => (
          <MenuItem key={d.color} icon={<Swatch colors={[d.color, '#ffffff']} />} onSelect={() => void addRule({ type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: d.color } })}>
            {d.label}
          </MenuItem>
        ))}
      </SubMenu>
      <SubMenu label="Color scales">
        {COLOR_SCALES.map((s) => (
          <MenuItem
            key={s.label}
            icon={<Swatch colors={[...s.colors].reverse()} />}
            onSelect={() => void addRule({ type: 'colorScale', colorScale: { cfvo: s.colors.length === 3 ? [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }], colors: s.colors } })}
          >
            {s.label}
          </MenuItem>
        ))}
      </SubMenu>
      <SubMenu label="Icon sets">
        {ICON_SETS.map((s) => (
          <MenuItem key={s.set} icon={<IconStrip set={s.set} />} onSelect={() => void addRule({ type: 'iconSet', iconSet: { set: s.set, cfvo: percentCfvo(ICON_COUNTS[s.set]) } })}>
            {s.label}
          </MenuItem>
        ))}
      </SubMenu>
      <MenuSep />
      <MenuItem onSelect={() => preset('formula')}>New rule with a formula…</MenuItem>
      <SubMenu label="Clear rules">
        <MenuItem
          onSelect={() => {
            const c = A.ctx()
            if (c) void A.edit(c, (b) => b.api.clearCondFormats(c.view.sheet, areasOf(c.view)))
          }}
        >
          From selected cells
        </MenuItem>
        <MenuItem
          onSelect={() => {
            const c = A.ctx()
            if (c) void A.edit(c, (b) => b.api.clearCondFormats(c.view.sheet, null))
          }}
        >
          From entire sheet
        </MenuItem>
      </SubMenu>
      <MenuItem onSelect={() => open({ kind: 'cf-manage' })}>Manage rules…</MenuItem>
    </>
  )
}

const TITLES: Record<CfPreset, [string, string]> = {
  greaterThan: ['Greater than', 'Format cells that are greater than:'],
  lessThan: ['Less than', 'Format cells that are less than:'],
  between: ['Between', 'Format cells that are between:'],
  equal: ['Equal to', 'Format cells that are equal to:'],
  text: ['Text that contains', 'Format cells that contain the text:'],
  date: ['A date occurring', 'Format cells that contain a date occurring:'],
  duplicate: ['Duplicate values', 'Format cells that contain:'],
  top: ['Top 10 items', 'Format cells that rank in the top:'],
  topPercent: ['Top 10%', 'Format cells that rank in the top (percent):'],
  bottom: ['Bottom 10 items', 'Format cells that rank in the bottom:'],
  bottomPercent: ['Bottom 10%', 'Format cells that rank in the bottom (percent):'],
  formula: ['New formatting rule', 'Format cells where this formula is true:']
}

const PERIODS: [string, string][] = [
  ['yesterday', 'Yesterday'],
  ['today', 'Today'],
  ['tomorrow', 'Tomorrow'],
  ['last7Days', 'In the last 7 days'],
  ['lastWeek', 'Last week'],
  ['thisWeek', 'This week'],
  ['nextWeek', 'Next week'],
  ['lastMonth', 'Last month'],
  ['thisMonth', 'This month'],
  ['nextMonth', 'Next month']
]

/** Excel's formula for each date period (`a` is the first cell of the range). */
export function periodFormula(p: string, a: string): string {
  const d = `FLOOR(${a},1)`
  const r = `ROUNDDOWN(${a},0)`
  switch (p) {
    case 'today':
      return `${d}=TODAY()`
    case 'yesterday':
      return `${d}=TODAY()-1`
    case 'tomorrow':
      return `${d}=TODAY()+1`
    case 'last7Days':
      return `AND(TODAY()-${d}<=6,${d}<=TODAY())`
    case 'thisWeek':
      return `AND(TODAY()-${r}<=WEEKDAY(TODAY())-1,${r}-TODAY()<=7-WEEKDAY(TODAY()))`
    case 'lastWeek':
      return `AND(TODAY()-${r}>=(WEEKDAY(TODAY())),TODAY()-${r}<(WEEKDAY(TODAY())+7))`
    case 'nextWeek':
      return `AND(${r}-TODAY()>(7-WEEKDAY(TODAY())),${r}-TODAY()<(15-WEEKDAY(TODAY())))`
    case 'thisMonth':
      return `AND(MONTH(${a})=MONTH(TODAY()),YEAR(${a})=YEAR(TODAY()))`
    case 'lastMonth':
      return `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0-1)),YEAR(${a})=YEAR(EDATE(TODAY(),0-1)))`
    default:
      return `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0+1)),YEAR(${a})=YEAR(EDATE(TODAY(),0+1)))`
  }
}

/** A typed comparison value as a formula: numbers and formulas as they are, other text quoted. */
export function operand(v: string): string {
  const t = v.trim()
  if (t.startsWith('=')) return t.slice(1)
  if (t !== '' && Number.isFinite(Number(t))) return t
  return `"${t.replace(/"/g, '""')}"`
}

/** The rule for a preset dialog's inputs. `anchor` is the first cell of the selection (A1 style). */
export function presetSpec(p: CfPreset, values: string[], dxf: Style, anchor: string, opts: { unique?: boolean; period?: string } = {}): CfSpec {
  switch (p) {
    case 'greaterThan':
    case 'lessThan':
    case 'equal':
      return { type: 'cellIs', operator: p, formulas: [operand(values[0])], dxf }
    case 'between':
      return { type: 'cellIs', operator: 'between', formulas: [operand(values[0]), operand(values[1])], dxf }
    case 'text': {
      const q = values[0].replace(/"/g, '""')
      return { type: 'containsText', text: values[0], formulas: [`NOT(ISERROR(SEARCH("${q}",${anchor})))`], dxf }
    }
    case 'date':
      return { type: 'timePeriod', timePeriod: opts.period ?? 'today', formulas: [periodFormula(opts.period ?? 'today', anchor)], dxf }
    case 'duplicate':
      return { type: opts.unique ? 'uniqueValues' : 'duplicateValues', dxf }
    case 'top':
    case 'topPercent':
    case 'bottom':
    case 'bottomPercent':
      return { type: 'top10', rank: Math.max(1, Math.round(Number(values[0]) || 10)), percent: p.endsWith('Percent') || undefined, bottom: p.startsWith('bottom') || undefined, dxf }
    default:
      return { type: 'expression', formulas: [values[0].replace(/^=?/, '')], dxf }
  }
}

export function CfRuleDialog({ preset }: { preset: CfPreset }) {
  const c = A.ctx()
  const [a, setA] = useState(preset.startsWith('top') || preset.startsWith('bottom') ? '10' : preset === 'formula' ? '=' : '')
  const [b, setB] = useState('')
  const [fmt, setFmt] = useState(0)
  const [period, setPeriod] = useState('today')
  const [unique, setUnique] = useState(false)
  const [error, setError] = useState('')
  if (!c) return null
  const first = areasOf(c.view)[0]
  const anchor = cellName(first.r0, first.c0)
  const [title, prompt] = TITLES[preset]
  const ok = async () => {
    const spec = presetSpec(preset, [a, b], FORMAT_PRESETS[fmt].style, anchor, { unique, period })
    const r = await c.book.api.addCondFormat(c.view.sheet, areasOf(c.view), spec)
    if (!r.ok) return setError(r.error ?? 'That rule isn’t valid.')
    markDirty(c.tab)
    close()
    refocus()
  }
  const input = (v: string, on: (v: string) => void, label: string) => (
    <input aria-label={label} autoFocus={label !== 'and'} className={cx(inputCls, 'flex-1')} value={v} onChange={(e) => on(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && void ok()} />
  )
  return (
    <Modal
      open
      onOpenChange={(o) => !o && close()}
      title={title}
      width={480}
      footer={
        <>
          <Btn onClick={close}>Cancel</Btn>
          <Btn variant="primary" onClick={() => void ok()}>
            OK
          </Btn>
        </>
      }
    >
      <p className="mb-2 text-[12.5px] text-text-2">{prompt}</p>
      <div className="mb-3 flex items-center gap-2 text-[12.5px]">
        {preset === 'date' ? (
          <select aria-label="Date" className={cx(inputCls, 'flex-1')} value={period} onChange={(e) => setPeriod(e.target.value)}>
            {PERIODS.map(([v, l]) => (
              <option key={v} value={v}>
                {l}
              </option>
            ))}
          </select>
        ) : preset === 'duplicate' ? (
          <select aria-label="Values" className={cx(inputCls, 'flex-1')} value={unique ? 'u' : 'd'} onChange={(e) => setUnique(e.target.value === 'u')}>
            <option value="d">Duplicate values</option>
            <option value="u">Unique values</option>
          </select>
        ) : (
          input(a, setA, preset === 'formula' ? 'Formula' : 'Value')
        )}
        {preset === 'between' && (
          <>
            <span className="text-text-3">and</span>
            {input(b, setB, 'and')}
          </>
        )}
      </div>
      <Field label="Format with">
        <select aria-label="Format with" className={inputCls} value={fmt} onChange={(e) => setFmt(Number(e.target.value))}>
          {FORMAT_PRESETS.map((f, i) => (
            <option key={f.label} value={i}>
              {f.label}
            </option>
          ))}
        </select>
      </Field>
      {preset === 'formula' && <p className="mt-2 text-[11.5px] text-text-3">The formula is written for {anchor}, the first cell of the selection; it adjusts for the other cells like a copied formula.</p>}
      {error && (
        <p role="alert" className="mt-2 text-[12px] text-danger">
          {error}
        </p>
      )}
    </Modal>
  )
}

const OPS: Record<string, string> = { greaterThan: '>', lessThan: '<', greaterThanOrEqual: '≥', lessThanOrEqual: '≤', equal: '=', notEqual: '≠' }

export function describeRule(r: CfListItem): string {
  const f = r.formulas ?? []
  switch (r.type) {
    case 'cellIs':
      if (r.operator === 'between' || r.operator === 'notBetween') return `Cell value ${r.operator === 'between' ? 'between' : 'not between'} ${f[0]?.slice(1)} and ${f[1]?.slice(1)}`
      return `Cell value ${OPS[r.operator ?? ''] ?? r.operator} ${f[0]?.slice(1)}`
    case 'expression':
      return `Formula: ${f[0]}`
    case 'top10':
      return `${r.bottom ? 'Bottom' : 'Top'} ${r.rank ?? 10}${r.percent ? '%' : ''}`
    case 'aboveAverage':
      return r.aboveAverage === false ? 'Below average' : 'Above average'
    case 'duplicateValues':
      return 'Duplicate values'
    case 'uniqueValues':
      return 'Unique values'
    case 'containsText':
      return `Cell value contains “${r.text ?? ''}”`
    case 'notContainsText':
      return `Cell value does not contain “${r.text ?? ''}”`
    case 'beginsWith':
      return `Cell value begins with “${r.text ?? ''}”`
    case 'endsWith':
      return `Cell value ends with “${r.text ?? ''}”`
    case 'timePeriod':
      return `Date occurring: ${PERIODS.find(([v]) => v === r.timePeriod)?.[1].toLowerCase() ?? r.timePeriod}`
    case 'colorScale':
      return `${r.colors?.length === 3 ? '3' : '2'}-color scale`
    case 'dataBar':
      return 'Data bar'
    case 'iconSet':
      return 'Icon set'
    case 'containsBlanks':
      return 'Blank cells'
    case 'notContainsBlanks':
      return 'No blanks'
    case 'containsErrors':
      return 'Cells with errors'
    case 'notContainsErrors':
      return 'No errors'
  }
}

function Preview({ r }: { r: CfListItem }) {
  if (r.type === 'iconSet' && r.iconSet) return <IconStrip set={r.iconSet} />
  if (r.colors) return <Swatch colors={r.type === 'dataBar' ? [r.colors[0], '#ffffff'] : r.colors} />
  const s = r.dxf ?? {}
  return (
    <span
      className="inline-flex h-5 w-[72px] items-center justify-center rounded-sm border border-border text-[12px]"
      style={{ background: s.fill, color: s.color, fontWeight: s.bold ? 700 : undefined, fontStyle: s.italic ? 'italic' : undefined, textDecoration: s.strike ? 'line-through' : s.underline ? 'underline' : undefined, borderColor: s.borderTop?.color }}
    >
      AaBbCc
    </span>
  )
}

export function CfManageDialog() {
  const c = A.ctx()
  const [rules, setRules] = useState<CfListItem[]>([])
  const [sel, setSel] = useState<number | null>(null)
  const load = async () => {
    if (c) setRules(await c.book.api.listCondFormats(c.view.sheet))
  }
  useEffect(() => {
    void load()
  }, [])
  if (!c) return null
  const act = async (fn: () => Promise<unknown>) => {
    await A.edit(c, () => fn())
    await load()
  }
  const chosen = rules.find((r) => r.id === sel)
  return (
    <Modal open onOpenChange={(o) => !o && (close(), refocus())} title="Conditional formatting rules" description="Rules on this sheet, highest priority first." width={640} footer={<Btn variant="primary" onClick={() => (close(), refocus())}>Close</Btn>}>
      <div className="mb-2 flex gap-1.5">
        <Btn size="sm" disabled={!chosen} onClick={() => chosen && void act(() => c.book.api.moveCondRule(c.view.sheet, chosen.id, -1))}>
          <ArrowUp className="size-3.5" /> Move up
        </Btn>
        <Btn size="sm" disabled={!chosen} onClick={() => chosen && void act(() => c.book.api.moveCondRule(c.view.sheet, chosen.id, 1))}>
          <ArrowDown className="size-3.5" /> Move down
        </Btn>
        <Btn size="sm" disabled={!chosen} onClick={() => chosen && void act(() => c.book.api.deleteCondRule(c.view.sheet, chosen.id))}>
          <Trash2 className="size-3.5" /> Delete rule
        </Btn>
      </div>
      <div role="listbox" aria-label="Rules" className="max-h-[320px] overflow-y-auto rounded-lg border border-border">
        {!rules.length && <p className="p-4 text-center text-[12.5px] text-text-3">This sheet has no conditional formatting.</p>}
        {rules.map((r) => (
          <div
            key={r.id}
            role="option"
            aria-selected={r.id === sel}
            tabIndex={0}
            onClick={() => setSel(r.id)}
            onKeyDown={(e) => e.key === ' ' && setSel(r.id)}
            className={cx('grid cursor-default grid-cols-[1fr_90px_150px] items-center gap-3 border-b border-border px-3 py-2 text-[12.5px] last:border-b-0', r.id === sel ? 'bg-accent-soft' : 'hover:bg-hover')}
          >
            <span className="truncate" title={describeRule(r)}>
              {describeRule(r)}
              {r.stopIfTrue && <span className="ml-2 text-[11px] text-text-3">stop if true</span>}
            </span>
            <Preview r={r} />
            <span className="truncate font-mono text-[11.5px] text-text-2" title={r.ranges}>
              {r.ranges}
            </span>
          </div>
        ))}
      </div>
    </Modal>
  )
}
