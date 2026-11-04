// Data tools dialogs: Remove Duplicates, Text to Columns, Goal Seek, Descriptive Statistics.
import { useEffect, useState } from 'react'
import { areaName, cellName, colName, parseArea, parseCell } from '@/core/refs'
import { set, toast, type Dialog } from '../state/store'
import { Btn, Field, Modal, inputCls } from '../ui'
import * as A from './actions'
import { currentRegion, guessHeader } from './commands'
import { closeDialog, focusGrid } from './focus'

const close = closeDialog
const refocus = focusGrid
const done = () => (close(), refocus())

export type ToolKind = 'dedupe' | 'text-to-columns' | 'goal-seek' | 'descriptive'

export function openTool(tool: ToolKind) {
  set({ dialog: { kind: 'tool', tool } })
}

export function ToolDialog({ d }: { d: Extract<Dialog, { kind: 'tool' }> }) {
  if (d.tool === 'dedupe') return <Dedupe />
  if (d.tool === 'text-to-columns') return <TextToColumns />
  if (d.tool === 'goal-seek') return <GoalSeek />
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
