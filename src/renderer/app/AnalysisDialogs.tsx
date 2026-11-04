// Data tools dialogs: Remove Duplicates, Text to Columns, Goal Seek, Descriptive Statistics.
import { useEffect, useState } from 'react'
import type { SolverConstraint, SolverModel, SolverRel, SolverResult } from '@/core/solver'
import type { Scenario, ScenarioSet } from '@/core/model/workbook'
import { setView } from '../state/view'
import { areaName, cellName, colName, parseArea, parseCell } from '@/core/refs'
import { set, toast, type Dialog } from '../state/store'
import { Btn, Field, Modal, cx, inputCls } from '../ui'
import * as A from './actions'
import { currentRegion, guessHeader } from './commands'
import { closeDialog, focusGrid } from './focus'

const close = closeDialog
const refocus = focusGrid
const done = () => (close(), refocus())

export type ToolKind = 'dedupe' | 'text-to-columns' | 'goal-seek' | 'descriptive' | 'solver' | 'data-table' | 'scenarios'

export function openTool(tool: ToolKind) {
  set({ dialog: { kind: 'tool', tool } })
}

export function ToolDialog({ d }: { d: Extract<Dialog, { kind: 'tool' }> }) {
  if (d.tool === 'dedupe') return <Dedupe />
  if (d.tool === 'text-to-columns') return <TextToColumns />
  if (d.tool === 'goal-seek') return <GoalSeek />
  if (d.tool === 'solver') return <Solver />
  if (d.tool === 'data-table') return <DataTableDialog />
  if (d.tool === 'scenarios') return <ScenarioManager />
  return <Descriptive />
}

function useRegion() {
  const c = A.ctx()
  const [area, setArea] = useState<{ r0: number; c0: number; r1: number; c1: number } | null>(null)
  useEffect(() => {
    if (!c) return
    const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
    if (r.r0 === r.r1 && r.c0 === r.c1) void currentRegion(c).then(setArea)
    else setArea(r)
  }, [])
  return { c, area }
}

function Dedupe() {
  const { c, area } = useRegion()
  const [headers, setHeaders] = useState(true)
  const [cols, setCols] = useState<number[] | null>(null)
  useEffect(() => {
    if (c && area) {
      setHeaders(guessHeader(c, area))
      setCols(Array.from({ length: area.c1 - area.c0 + 1 }, (_, i) => i))
    }
  }, [area])
  if (!c || !area || !cols) return null
  const name = (j: number) => (headers ? c.book.cell(c.view.sheet, area.r0, area.c0 + j)?.t || `Column ${colName(area.c0 + j)}` : `Column ${colName(area.c0 + j)}`)
  const ok = async () => {
    const r = await c.book.api.removeDuplicates({ sheet: c.view.sheet, ...area }, cols, headers)
    done()
    if (!r.ok) return toast(r.error, 'error')
    A.markDirtyTab(c.tab)
    toast(r.removed ? `${r.removed} duplicate ${r.removed === 1 ? 'value was' : 'values were'} removed; ${r.kept} unique remain.` : 'No duplicate values were found.', 'info')
  }
  return (
    <Modal open onOpenChange={(o) => !o && done()} title="Remove duplicates" width={400} footer={<><Btn onClick={done}>Cancel</Btn><Btn variant="primary" disabled={!cols.length} onClick={() => void ok()}>OK</Btn></>}>
      <p className="mb-2 text-[12.5px] text-text-2">In {areaName(area)}, rows that match in the ticked columns are removed (the first stays).</p>
      <label className="mb-2 flex items-center gap-2 text-[12.5px]">
        <input type="checkbox" checked={headers} onChange={(e) => setHeaders(e.target.checked)} /> My data has headers
      </label>
      <div className="max-h-[220px] overflow-y-auto rounded border border-border p-1 text-[12.5px]">
        {Array.from({ length: area.c1 - area.c0 + 1 }, (_, j) => (
          <label key={j} className="flex items-center gap-2 px-1 py-0.5">
            <input type="checkbox" checked={cols.includes(j)} onChange={(e) => setCols((x) => (e.target.checked ? [...x!, j].sort((a, b) => a - b) : x!.filter((k) => k !== j)))} /> {name(j)}
          </label>
        ))}
      </div>
    </Modal>
  )
}

function TextToColumns() {
  const c = A.ctx()
  const [delims, setDelims] = useState<string[]>([','])
  const [other, setOther] = useState('')
  const [consecutive, setConsecutive] = useState(false)
  if (!c) return null
  const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  const opts: [string, string][] = [
    ['\t', 'Tab'],
    [';', 'Semicolon'],
    [',', 'Comma'],
    [' ', 'Space']
  ]
  const ok = async () => {
    const all = [...delims, ...(other ? [other[0]] : [])]
    const res = await c.book.api.textToColumns({ sheet: c.view.sheet, ...r, c1: r.c0 }, all, consecutive)
    done()
    if (!res.ok) return toast(res.error, 'error')
    if (!res.columns) return toast('Nothing to split: no cell contains those delimiters.', 'info')
    A.markDirtyTab(c.tab)
  }
  return (
    <Modal open onOpenChange={(o) => !o && done()} title="Text to columns" width={380} footer={<><Btn onClick={done}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
      <p className="mb-2 text-[12.5px] text-text-2">Split the text in {areaName({ ...r, c1: r.c0 })} into the columns to its right, at:</p>
      <div className="space-y-1 text-[12.5px]">
        {opts.map(([v, l]) => (
          <label key={l} className="flex items-center gap-2">
            <input type="checkbox" checked={delims.includes(v)} onChange={(e) => setDelims((x) => (e.target.checked ? [...x, v] : x.filter((k) => k !== v)))} /> {l}
          </label>
        ))}
        <label className="flex items-center gap-2">
          Other: <input aria-label="Other delimiter" maxLength={1} className={`${inputCls} h-7 w-12`} value={other} onChange={(e) => setOther(e.target.value)} />
        </label>
        <label className="flex items-center gap-2 pt-1">
          <input type="checkbox" checked={consecutive} onChange={(e) => setConsecutive(e.target.checked)} /> Treat consecutive delimiters as one
        </label>
      </div>
    </Modal>
  )
}

function GoalSeek() {
  const c = A.ctx()
  const a = c?.view.sel.active
  const [setCell, setSetCell] = useState(a ? cellName(a.r, a.c) : '')
  const [to, setTo] = useState('')
  const [changing, setChanging] = useState('')
  const [error, setError] = useState('')
  if (!c) return null
  const ok = async () => {
    const t = parseCell(setCell.replace(/\$/g, ''))
    const ch = parseCell(changing.replace(/\$/g, ''))
    const goal = Number(to)
    if (!t || !ch || to.trim() === '' || !Number.isFinite(goal)) return setError('Enter a cell with a formula, a number to reach, and the cell to change.')
    const r = await c.book.api.goalSeek({ sheet: c.view.sheet, ...t }, goal, { sheet: c.view.sheet, ...ch })
    if (!r.ok) return setError(r.error)
    done()
    A.markDirtyTab(c.tab)
    set({ dialog: { kind: 'message', title: 'Goal Seek status', body: r.found ? `Goal Seeking with cell ${setCell.toUpperCase()} found a solution.\n\nTarget value: ${goal}\nCurrent value: ${Math.round(r.result * 1e9) / 1e9}` : `Goal Seeking with cell ${setCell.toUpperCase()} may not have found a solution.` } })
  }
  return (
    <Modal open onOpenChange={(o) => !o && done()} title="Goal seek" width={360} footer={<><Btn onClick={done}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
      <div className="space-y-3">
        <Field label="Set cell">
          <input aria-label="Set cell" autoFocus className={inputCls} value={setCell} onChange={(e) => setSetCell(e.target.value)} />
        </Field>
        <Field label="To value">
          <input aria-label="To value" className={inputCls} value={to} onChange={(e) => setTo(e.target.value)} />
        </Field>
        <Field label="By changing cell">
          <input aria-label="By changing cell" className={inputCls} value={changing} onChange={(e) => setChanging(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && void ok()} />
        </Field>
        {error && <p role="alert" className="text-[12px] text-danger">{error}</p>}
      </div>
    </Modal>
  )
}

function Descriptive() {
  const { c, area } = useRegion()
  const [input, setInput] = useState('')
  const [out, setOut] = useState('')
  const [labels, setLabels] = useState(true)
  useEffect(() => {
    if (area) {
      setInput(areaName(area))
      setOut(cellName(area.r0, area.c1 + 2))
    }
  }, [area])
  if (!c) return null
  const ok = async () => {
    const i = parseArea(input.replace(/\$/g, ''), c.view.sheet)
    const o = parseCell(out.replace(/\$/g, ''))
    if (!i || !o) return toast('Enter an input range and an output cell.', 'error')
    const r = await c.book.api.descriptiveStats(i, { sheet: c.view.sheet, ...o }, labels)
    done()
    if (!r.ok) return toast(r.error ?? 'That didn’t work.', 'error')
    A.markDirtyTab(c.tab)
  }
  return (
    <Modal open onOpenChange={(o) => !o && done()} title="Descriptive statistics" width={380} footer={<><Btn onClick={done}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
      <div className="space-y-3 text-[12.5px]">
        <Field label="Input range (columns)">
          <input aria-label="Input range" autoFocus className={inputCls} value={input} onChange={(e) => setInput(e.target.value)} />
        </Field>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={labels} onChange={(e) => setLabels(e.target.checked)} /> Labels in first row
        </label>
        <Field label="Output range">
          <input aria-label="Output range" className={inputCls} value={out} onChange={(e) => setOut(e.target.value)} />
        </Field>
      </div>
    </Modal>
  )
}

// ---------- Solver ----------

const RELS: [SolverRel, string][] = [
  ['<=', '<='],
  ['=', '='],
  ['>=', '>='],
  ['int', 'int'],
  ['bin', 'bin']
]
const describe = (k: SolverConstraint) => (k.rel === 'int' ? `${k.lhs} = integer` : k.rel === 'bin' ? `${k.lhs} = binary` : `${k.lhs} ${k.rel} ${k.rhs}`)

function Solver() {
  const c = A.ctx()
  const [model, setModel] = useState<SolverModel | null>(null)
  const [initial, setInitial] = useState('')
  const [sel, setSel] = useState(-1)
  // The constraint being added (index −1) or changed.
  const [edit, setEdit] = useState<{ i: number; k: SolverConstraint } | null>(null)
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [result, setResult] = useState<SolverResult | null>(null)
  const [keep, setKeep] = useState(true)
  const [report, setReport] = useState(false)
  useEffect(() => {
    if (!c) return
    void c.book.api.solverModel(c.view.sheet).then((m) => {
      const a = c.view.sel.active
      const m2 = m.objective ? m : { ...m, objective: cellName(a.r, a.c).replace(/([A-Z]+)(\d+)/, '$$$1$$$2') }
      setModel(m2)
      setInitial(JSON.stringify(m2))
    })
  }, [])
  if (!c || !model) return null
  const up = (p: Partial<SolverModel>) => setModel({ ...model, ...p })
  const finish = async (keepIt: SolverResult | null, rep: SolverResult | null) => {
    // A model is saved once it has something to solve (as Excel does when you Solve or Close).
    const complete = !!model.objective.trim() && !!model.changing.trim()
    if (keepIt || rep || (complete && JSON.stringify(model) !== initial)) {
      const r = await c.book.api.finishSolver(c.view.sheet, model, keepIt, rep)
      if (!r.ok) toast(r.error ?? 'Solver couldn’t finish.', 'error')
      else A.markDirtyTab(c.tab)
    }
    done()
  }
  const run = async () => {
    setError('')
    setBusy(true)
    const r = await c.book.api.solve(c.view.sheet, model)
    setBusy(false)
    if (!r.ok) return setError(r.error)
    setResult(r.result)
    setKeep(r.result.status === 'optimal' || r.result.status === 'converged')
  }
  if (result) {
    const good = result.status === 'optimal' || result.status === 'converged'
    return (
      <Modal open onOpenChange={(o) => !o && void finish(null, null)} title="Solver results" width={440} footer={<><Btn onClick={() => void finish(null, null)}>Cancel</Btn><Btn variant="primary" onClick={() => void finish(keep && good ? result : null, report ? result : null)}>OK</Btn></>}>
        <div className="space-y-3 text-[12.5px]">
          <p role="status" className={good ? 'text-text' : 'text-danger'}>
            {result.message}
          </p>
          {good && (
            <p className="text-text-2">
              Objective: {Math.round(result.objective * 1e9) / 1e9} (was {Math.round(result.originalObjective * 1e9) / 1e9})
            </p>
          )}
          <fieldset className="space-y-1">
            <legend className="sr-only">Solution</legend>
            <label className="flex items-center gap-2">
              <input type="radio" name="solver-keep" checked={keep && good} disabled={!good} onChange={() => setKeep(true)} /> Keep Solver solution
            </label>
            <label className="flex items-center gap-2">
              <input type="radio" name="solver-keep" checked={!keep || !good} onChange={() => setKeep(false)} /> Restore original values
            </label>
          </fieldset>
          <label className="flex items-center gap-2">
            <input type="checkbox" checked={report} disabled={!good} onChange={(e) => setReport(e.target.checked)} /> Answer report (a new sheet)
          </label>
        </div>
      </Modal>
    )
  }
  const saveEdit = () => {
    if (!edit) return
    const k = edit.k
    if (!k.lhs.trim() || (k.rel !== 'int' && k.rel !== 'bin' && !k.rhs.trim())) return setError('Enter a cell reference and a constraint value.')
    const list = [...model.constraints]
    if (edit.i < 0) list.push(k)
    else list[edit.i] = k
    up({ constraints: list })
    setEdit(null)
    setError('')
  }
  return (
    <Modal open onOpenChange={(o) => !o && void finish(null, null)} title="Solver" width={560} footer={<><Btn onClick={() => void finish(null, null)}>Close</Btn><Btn variant="primary" disabled={busy || !!edit} onClick={() => void run()}>{busy ? 'Solving…' : 'Solve'}</Btn></>}>
      <div className="space-y-3 text-[12.5px]">
        <Field label="Set objective">
          <input aria-label="Set objective" autoFocus className={inputCls} value={model.objective} onChange={(e) => up({ objective: e.target.value })} />
        </Field>
        <fieldset className="flex flex-wrap items-center gap-x-4 gap-y-1">
          <legend className="mb-1 text-[12px] font-medium text-text-2">To</legend>
          {(['max', 'min', 'value'] as const).map((g) => (
            <label key={g} className="flex items-center gap-2">
              <input type="radio" name="solver-goal" checked={model.goal === g} onChange={() => up({ goal: g })} /> {g === 'max' ? 'Max' : g === 'min' ? 'Min' : 'Value of'}
            </label>
          ))}
          <input aria-label="Target value" className={cx(inputCls, 'w-28')} disabled={model.goal !== 'value'} value={String(model.value)} onChange={(e) => up({ value: Number(e.target.value) || 0 })} />
        </fieldset>
        <Field label="By changing variable cells">
          <input aria-label="By changing variable cells" className={inputCls} value={model.changing} onChange={(e) => up({ changing: e.target.value })} />
        </Field>
        <div>
          <span id="solver-constraints" className="mb-1.5 block text-[12px] font-medium text-text-2">
            Subject to the constraints
          </span>
          <div className="flex gap-2">
            <select aria-labelledby="solver-constraints" size={5} className={cx(inputCls, 'h-auto flex-1 py-1')} value={sel < 0 ? '' : String(sel)} onChange={(e) => setSel(Number(e.target.value))}>
              {model.constraints.map((k, i) => (
                <option key={i} value={i}>
                  {describe(k)}
                </option>
              ))}
            </select>
            <div className="flex w-24 flex-col gap-1.5">
              <Btn onClick={() => setEdit({ i: -1, k: { lhs: '', rel: '<=', rhs: '' } })}>Add</Btn>
              <Btn disabled={sel < 0 || sel >= model.constraints.length} onClick={() => setEdit({ i: sel, k: { ...model.constraints[sel] } })}>
                Change
              </Btn>
              <Btn
                disabled={sel < 0 || sel >= model.constraints.length}
                onClick={() => {
                  up({ constraints: model.constraints.filter((_, i) => i !== sel) })
                  setSel(-1)
                }}
              >
                Delete
              </Btn>
              <Btn onClick={() => (setModel({ ...model, objective: '', changing: '', constraints: [], goal: 'max', value: 0 }), setSel(-1))}>Reset all</Btn>
            </div>
          </div>
        </div>
        {edit && (
          <fieldset className="flex items-end gap-2 rounded-lg border border-border p-2">
            <legend className="px-1 text-[12px] font-medium text-text-2">{edit.i < 0 ? 'Add constraint' : 'Change constraint'}</legend>
            <Field label="Cell reference">
              <input aria-label="Cell reference" autoFocus className={cx(inputCls, 'w-32')} value={edit.k.lhs} onChange={(e) => setEdit({ ...edit, k: { ...edit.k, lhs: e.target.value } })} />
            </Field>
            <Field label="Relation">
              <select aria-label="Relation" className={cx(inputCls, 'w-20')} value={edit.k.rel} onChange={(e) => setEdit({ ...edit, k: { ...edit.k, rel: e.target.value as SolverRel } })}>
                {RELS.map(([v, l]) => (
                  <option key={v} value={v}>
                    {l}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Constraint">
              <input aria-label="Constraint" className={cx(inputCls, 'w-32')} disabled={edit.k.rel === 'int' || edit.k.rel === 'bin'} value={edit.k.rel === 'int' ? 'integer' : edit.k.rel === 'bin' ? 'binary' : edit.k.rhs} onChange={(e) => setEdit({ ...edit, k: { ...edit.k, rhs: e.target.value } })} onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), saveEdit())} />
            </Field>
            <Btn variant="primary" onClick={saveEdit}>
              {edit.i < 0 ? 'Add' : 'OK'}
            </Btn>
            <Btn onClick={() => setEdit(null)}>Cancel</Btn>
          </fieldset>
        )}
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={model.nonNegative} onChange={(e) => up({ nonNegative: e.target.checked })} /> Make unconstrained variables non-negative
        </label>
        <Field label="Solving method" hint={model.method === 'simplex' ? 'For linear problems; integer and binary constraints by branch and bound.' : 'Nelder–Mead with penalties: finds a local optimum. Not Excel’s GRG method.'}>
          <select aria-label="Solving method" className={inputCls} value={model.method} onChange={(e) => up({ method: e.target.value as SolverModel['method'] })}>
            <option value="simplex">Simplex LP</option>
            <option value="nonlinear">Nonlinear (Nelder–Mead)</option>
          </select>
        </Field>
        {error && (
          <p role="alert" className="text-[12px] text-danger">
            {error}
          </p>
        )}
      </div>
    </Modal>
  )
}

// ---------- Data Table (what-if) ----------

function DataTableDialog() {
  const c = A.ctx()
  const [row, setRow] = useState('')
  const [col, setCol] = useState('')
  const [error, setError] = useState('')
  if (!c) return null
  const area = c.view.sel.ranges[c.view.sel.ranges.length - 1]
  const ok = async () => {
    const cellOf = (t: string) => (t.trim() ? parseCell(t.replace(/\$/g, '').replace(/^.*!/, '')) : null)
    const ri = cellOf(row)
    const ci = cellOf(col)
    if ((row.trim() && !ri) || (col.trim() && !ci)) return setError('Input cells must be single cells, like $B$1.')
    const r = await c.book.api.createDataTable(c.view.sheet, area, ri, ci)
    if (!r.ok) return setError(r.error ?? 'The data table couldn’t be made.')
    A.markDirtyTab(c.tab)
    done()
  }
  return (
    <Modal open onOpenChange={(o) => !o && done()} title="Data table" description={`For ${areaName(area)}: formulas and input values in its first row and column, results in the rest.`} width={380} footer={<><Btn onClick={done}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
      <div className="space-y-3 text-[12.5px]">
        <Field label="Row input cell" hint="The cell the values across the first row go into.">
          <input aria-label="Row input cell" autoFocus className={inputCls} value={row} onChange={(e) => setRow(e.target.value)} />
        </Field>
        <Field label="Column input cell" hint="The cell the values down the first column go into.">
          <input aria-label="Column input cell" className={inputCls} value={col} onChange={(e) => setCol(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && void ok()} />
        </Field>
        {error && (
          <p role="alert" className="text-[12px] text-danger">
            {error}
          </p>
        )}
      </div>
    </Modal>
  )
}

// ---------- Scenario Manager ----------

const refText = (cells: { r: number; c: number }[]) => cells.map((p) => cellName(p.r, p.c).replace(/([A-Z]+)(\d+)/, '$$$1$$$2')).join(',')

function ScenarioManager() {
  const c = A.ctx()
  const [set0, setSet] = useState<ScenarioSet | null | undefined>(undefined)
  const [sel, setSel] = useState(0)
  // Adding (index −1) or editing a scenario: its name, cells and comment, then its values.
  const [form, setForm] = useState<{ i: number; name: string; cells: string; comment: string; values: string[] | null } | null>(null)
  const [summary, setSummary] = useState<string | null>(null)
  const [error, setError] = useState('')
  const load = async () => {
    if (!c) return
    const s = await c.book.api.scenarios(c.view.sheet)
    setSet(s)
    if (s) setSel(Math.min(sel, s.list.length - 1))
  }
  useEffect(() => void load(), [])
  if (!c || set0 === undefined) return null
  const list = set0?.list ?? []
  const save = async (next: Scenario[], current: number) => {
    const r = await c.book.api.setScenarios(c.view.sheet, next, current)
    if (!r.ok) return setError(r.error ?? 'That didn’t work.'), false
    A.markDirtyTab(c.tab)
    setError('')
    await load()
    setSel(current)
    return true
  }
  const startAdd = () => {
    const r = c.view.sel.ranges[c.view.sel.ranges.length - 1]
    setForm({ i: -1, name: '', cells: areaName(r).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'), comment: '', values: null })
    setError('')
  }
  const next = async () => {
    if (!form) return
    if (!form.name.trim()) return setError('Enter a name for the scenario.')
    if (form.values === null) {
      // Second step: one value per changing cell, starting from what the cells hold now (or the scenario's).
      const cells = cellsFromText(form.cells)
      if (typeof cells === 'string') return setError(cells)
      const old = form.i >= 0 ? list[form.i] : null
      const values = await Promise.all(cells.map(async (p, k) => (old && refText(old.cells) === refText(cells) ? old.values[k] : await c.book.api.editText(c.view.sheet, p.r, p.c))))
      setError('')
      return setForm({ ...form, values })
    }
    const cells = cellsFromText(form.cells) as { r: number; c: number }[]
    const sc: Scenario = { ...(form.i >= 0 ? list[form.i] : {}), name: form.name.trim(), cells, values: form.values, comment: form.comment.trim() || undefined }
    const all = form.i >= 0 ? list.map((x, k) => (k === form.i ? sc : x)) : [...list, sc]
    if (await save(all, form.i >= 0 ? form.i : all.length - 1)) setForm(null)
  }
  if (form) {
    return (
      <Modal open onOpenChange={(o) => !o && setForm(null)} title={form.i < 0 ? 'Add scenario' : 'Edit scenario'} width={420} footer={<><Btn onClick={() => setForm(null)}>Cancel</Btn><Btn variant="primary" onClick={() => void next()}>{form.values === null ? 'Next' : 'OK'}</Btn></>}>
        <div className="space-y-3 text-[12.5px]">
          {form.values === null ? (
            <>
              <Field label="Scenario name">
                <input aria-label="Scenario name" autoFocus className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
              </Field>
              <Field label="Changing cells" hint="Up to 32 cells, such as $B$1:$B$3,$D$1.">
                <input aria-label="Changing cells" className={inputCls} value={form.cells} onChange={(e) => setForm({ ...form, cells: e.target.value })} />
              </Field>
              <Field label="Comment">
                <input aria-label="Comment" className={inputCls} value={form.comment} onChange={(e) => setForm({ ...form, comment: e.target.value })} />
              </Field>
            </>
          ) : (
            <div className="max-h-[50vh] space-y-2 overflow-y-auto">
              <p className="text-text-2">Values for each changing cell:</p>
              {(cellsFromText(form.cells) as { r: number; c: number }[]).map((p, k) => (
                <Field key={k} label={cellName(p.r, p.c)}>
                  <input aria-label={`Value for ${cellName(p.r, p.c)}`} autoFocus={k === 0} className={inputCls} value={form.values![k]} onChange={(e) => setForm({ ...form, values: form.values!.map((v, j) => (j === k ? e.target.value : v)) })} />
                </Field>
              ))}
            </div>
          )}
          {error && (
            <p role="alert" className="text-[12px] text-danger">
              {error}
            </p>
          )}
        </div>
      </Modal>
    )
  }
  if (summary !== null) {
    const ok = async () => {
      const r = await c.book.api.scenarioSummary(c.view.sheet, summary)
      if (!r.ok) return setError(r.error)
      A.markDirtyTab(c.tab)
      await c.book.refreshMeta()
      setView(c.tab, { sheet: r.sheet })
      done()
    }
    return (
      <Modal open onOpenChange={(o) => !o && setSummary(null)} title="Scenario summary" width={380} footer={<><Btn onClick={() => setSummary(null)}>Cancel</Btn><Btn variant="primary" onClick={() => void ok()}>OK</Btn></>}>
        <div className="space-y-3 text-[12.5px]">
          <Field label="Result cells" hint="The cells whose values the report shows for each scenario.">
            <input aria-label="Result cells" autoFocus className={inputCls} value={summary} onChange={(e) => setSummary(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && void ok()} />
          </Field>
          {error && (
            <p role="alert" className="text-[12px] text-danger">
              {error}
            </p>
          )}
        </div>
      </Modal>
    )
  }
  const cur = list[sel]
  return (
    <Modal open onOpenChange={(o) => !o && done()} title="Scenario manager" width={480} footer={<Btn variant="primary" onClick={done}>Close</Btn>}>
      <div className="space-y-3 text-[12.5px]">
        <div className="flex gap-2">
          <div className="flex-1">
            <span id="scenario-list" className="mb-1.5 block text-[12px] font-medium text-text-2">
              Scenarios
            </span>
            {list.length ? (
              <select aria-labelledby="scenario-list" size={6} className={cx(inputCls, 'h-auto w-full py-1')} value={String(sel)} onChange={(e) => setSel(Number(e.target.value))}>
                {list.map((x, k) => (
                  <option key={k} value={k}>
                    {x.name}
                  </option>
                ))}
              </select>
            ) : (
              <p className="rounded-lg border border-border p-3 text-text-3">No scenarios on this sheet. Choose Add to make one.</p>
            )}
          </div>
          <div className="flex w-28 flex-col gap-1.5 pt-6">
            <Btn onClick={startAdd}>Add…</Btn>
            <Btn disabled={!cur} onClick={() => void save(list.filter((_, k) => k !== sel), Math.max(0, sel - 1))}>
              Delete
            </Btn>
            <Btn disabled={!cur} onClick={() => cur && (setForm({ i: sel, name: cur.name, cells: refText(cur.cells), comment: cur.comment ?? '', values: null }), setError(''))}>
              Edit…
            </Btn>
            <Btn
              disabled={!cur}
              onClick={async () => {
                const r = await c.book.api.showScenario(c.view.sheet, sel)
                if (!r.ok) return setError(r.error ?? 'That didn’t work.')
                A.markDirtyTab(c.tab)
              }}
            >
              Show
            </Btn>
            <Btn disabled={!cur} onClick={() => (setSummary(set0?.resultCells ?? cellName(c.view.sel.active.r, c.view.sel.active.c)), setError(''))}>
              Summary…
            </Btn>
          </div>
        </div>
        {cur && (
          <dl className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-1 text-text-2">
            <dt>Changing cells</dt>
            <dd className="font-mono text-[12px] text-text">{refText(cur.cells)}</dd>
            {cur.comment && (
              <>
                <dt>Comment</dt>
                <dd className="text-text">{cur.comment}</dd>
              </>
            )}
          </dl>
        )}
        {error && (
          <p role="alert" className="text-[12px] text-danger">
            {error}
          </p>
        )}
      </div>
    </Modal>
  )
}

/** "B1:B2,D1" → cells (at most 32), or an error message. */
function cellsFromText(text: string): { r: number; c: number }[] | string {
  const out: { r: number; c: number }[] = []
  for (const raw of text.split(',')) {
    const t = raw.trim().replace(/^.*!/, '')
    if (!t) continue
    const a = parseArea(t.replace(/\$/g, ''))
    if (!a) return `“${raw.trim()}” isn’t a cell or range.`
    for (let r = a.r0; r <= a.r1; r++) for (let cc = a.c0; cc <= a.c1; cc++) out.push({ r, c: cc })
  }
  if (!out.length) return 'Enter the changing cells.'
  if (out.length > 32) return 'A scenario can change at most 32 cells.'
  return out
}

