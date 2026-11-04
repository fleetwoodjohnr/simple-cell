// Data validation in the UI: the dialog, the alert after a rejected entry, and the grid overlay
// (in-cell dropdown, input message, Circle Invalid Data).
import { useEffect, useRef, useState } from 'react'
import { ChevronDown } from 'lucide-react'
import type { DvInfo, DvSpec } from '@/core/doc'
import type { DvAlert, DvOperator, DvType } from '@/core/validation'
import { colX, rowY, type Layout } from '@/grid/layout'
import type { Book } from '../state/client'
import { set, type Dialog } from '../state/store'
import { areasOf, getView, setView, useViews } from '../state/view'
import { Btn, Field, Modal, Segmented, cx, inputCls } from '../ui'
import * as A from './actions'
import { markDirty } from '../state/documents'
import { closeDialog, focusGrid } from './focus'

const close = closeDialog
const refocus = focusGrid

// ---------- alert ----------

export type AlertChoice = 'retry' | 'cancel' | 'yes' | 'no' | 'ok'

export function showValidationAlert(alert: DvAlert): Promise<AlertChoice> {
  return new Promise((resolve) => set({ dialog: { kind: 'dv-alert', alert, resolve } }))
}

export function ValidationAlert({ d }: { d: Extract<Dialog, { kind: 'dv-alert' }> }) {
  const done = (c: AlertChoice) => {
    close()
    d.resolve(c)
  }
  const { style, title, message } = d.alert
  const buttons: [AlertChoice, string][] = style === 'stop' ? [['cancel', 'Cancel'], ['retry', 'Retry']] : style === 'warning' ? [['cancel', 'Cancel'], ['no', 'No'], ['yes', 'Yes']] : [['cancel', 'Cancel'], ['ok', 'OK']]
  return (
    <Modal
      open
      onOpenChange={(o) => !o && done('cancel')}
      title={title}
      width={420}
      footer={buttons.map(([k, l], i) => (
        <Btn key={k} variant={i === buttons.length - 1 ? 'primary' : undefined} autoFocus={i === buttons.length - 1} onClick={() => done(k)}>
          {l}
        </Btn>
      ))}
    >
      <p role="alert" className="whitespace-pre-line text-[13px] leading-relaxed text-text-2">
        {message}
        {style === 'warning' && '\n\nContinue?'}
      </p>
    </Modal>
  )
}

// ---------- dialog ----------

const TYPES: [DvType, string][] = [
  ['none', 'Any value'],
  ['whole', 'Whole number'],
  ['decimal', 'Decimal'],
  ['list', 'List'],
  ['date', 'Date'],
  ['time', 'Time'],
  ['textLength', 'Text length'],
  ['custom', 'Custom']
]
const OPERATORS: [DvOperator, string][] = [
  ['between', 'between'],
  ['notBetween', 'not between'],
  ['equal', 'equal to'],
  ['notEqual', 'not equal to'],
  ['greaterThan', 'greater than'],
  ['lessThan', 'less than'],
  ['greaterThanOrEqual', 'greater than or equal to'],
  ['lessThanOrEqual', 'less than or equal to']
]

export function ValidationDialog() {
  const c = A.ctx()
  const [tab, setTab] = useState<'settings' | 'input' | 'error'>('settings')
  const [s, setS] = useState<DvSpec & { f1: string; f2: string }>({ type: 'none', operator: 'between', allowBlank: true, showInput: true, showError: true, errorStyle: 'stop', f1: '', f2: '' })
  const [error, setError] = useState('')
  useEffect(() => {
    if (!c) return
    void c.book.api.validationInfo(c.view.sheet, c.view.sel.active.r, c.view.sel.active.c).then((info: DvInfo | null) => {
      if (info) setS((x) => ({ ...x, ...info, operator: info.operator ?? 'between', allowBlank: info.allowBlank !== false, showInput: info.showInput !== false, showError: info.showError !== false, errorStyle: info.errorStyle ?? 'stop', f1: (info.formulas?.[0] ?? '').replace(/^=(?=[\d.-]+$)/, ''), f2: (info.formulas?.[1] ?? '').replace(/^=(?=[\d.-]+$)/, '') }))
    })
  }, [])
  if (!c) return null
  const up = (p: Partial<typeof s>) => setS((x) => ({ ...x, ...p }))
  const save = async (spec: DvSpec) => {
    const r = await c.book.api.setValidation(c.view.sheet, areasOf(c.view), spec)
    if (!r.ok) return setError(r.error ?? 'That rule isn’t valid.')
    markDirty(c.tab)
    close()
    refocus()
  }
  const ok = () => {
    const { f1, f2, ranges: _r, anchor: _a, items: _i, ...rest } = s as typeof s & Partial<DvInfo>
    void _r
    void _a
    void _i
    const two = s.operator === 'between' || s.operator === 'notBetween'
    const formulas = s.type === 'none' ? undefined : s.type === 'list' || s.type === 'custom' || !two ? [f1] : [f1, f2]
    if (formulas?.some((f) => !f.trim())) return setError(s.type === 'list' ? 'Enter the list items or a range for the source.' : 'Enter a value or formula.')
    void save({ ...rest, formulas, operator: s.type === 'list' || s.type === 'custom' || s.type === 'none' ? undefined : s.operator })
  }
  const hasOp = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(s.type)
  const two = s.operator === 'between' || s.operator === 'notBetween'
  const label1 = s.type === 'list' ? 'Source' : s.type === 'custom' ? 'Formula' : two ? 'Minimum' : 'Value'
  return (
    <Modal
      open
      onOpenChange={(o) => !o && (close(), refocus())}
      title="Data validation"
      width={520}
      footer={
        <>
          <Btn className="mr-auto" onClick={() => void save({ type: 'none' })}>
            Clear all
          </Btn>
          <Btn onClick={() => (close(), refocus())}>Cancel</Btn>
          <Btn variant="primary" onClick={ok}>
            OK
          </Btn>
        </>
      }
    >
      <Segmented
        label="Section"
        value={tab}
        onChange={(v) => setTab(v as typeof tab)}
        options={[
          { value: 'settings', label: 'Settings' },
          { value: 'input', label: 'Input message' },
          { value: 'error', label: 'Error alert' }
        ]}
      />
      <div className="mt-4 space-y-3 text-[12.5px]">
        {tab === 'settings' && (
          <>
            <Field label="Allow">
              <select aria-label="Allow" className={inputCls} value={s.type} onChange={(e) => up({ type: e.target.value as DvType })}>
                {TYPES.map(([v, l]) => (
                  <option key={v} value={v}>
                    {l}
                  </option>
                ))}
              </select>
            </Field>
            {hasOp && (
              <Field label="Data">
                <select aria-label="Data" className={inputCls} value={s.operator} onChange={(e) => up({ operator: e.target.value as DvOperator })}>
                  {OPERATORS.map(([v, l]) => (
                    <option key={v} value={v}>
                      {l}
                    </option>
                  ))}
                </select>
              </Field>
            )}
            {s.type !== 'none' && (
              <Field label={label1} hint={s.type === 'list' ? 'Items separated by commas, or a range such as =$D$1:$D$10' : s.type === 'custom' ? 'A formula for the first selected cell that is TRUE for valid entries' : undefined}>
                <input aria-label={label1} className={inputCls} value={s.f1} onChange={(e) => up({ f1: e.target.value })} />
              </Field>
            )}
            {hasOp && two && (
              <Field label="Maximum">
                <input aria-label="Maximum" className={inputCls} value={s.f2} onChange={(e) => up({ f2: e.target.value })} />
              </Field>
            )}
            {s.type !== 'none' && (
              <label className="flex items-center gap-2">
                <input type="checkbox" checked={s.allowBlank !== false} onChange={(e) => up({ allowBlank: e.target.checked })} /> Ignore blank
              </label>
            )}
            {s.type === 'list' && (
              <label className="flex items-center gap-2">
                <input type="checkbox" checked={!s.noDropdown} onChange={(e) => up({ noDropdown: !e.target.checked })} /> In-cell dropdown
              </label>
            )}
          </>
        )}
        {tab === 'input' && (
          <>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={s.showInput !== false} onChange={(e) => up({ showInput: e.target.checked })} /> Show input message when the cell is selected
            </label>
            <Field label="Title">
              <input aria-label="Input title" className={inputCls} value={s.promptTitle ?? ''} onChange={(e) => up({ promptTitle: e.target.value || undefined })} />
            </Field>
            <Field label="Input message">
              <textarea aria-label="Input message" className={cx(inputCls, 'h-20 py-1.5')} value={s.prompt ?? ''} onChange={(e) => up({ prompt: e.target.value || undefined })} />
            </Field>
          </>
        )}
        {tab === 'error' && (
          <>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={s.showError !== false} onChange={(e) => up({ showError: e.target.checked })} /> Show error alert after invalid data is entered
            </label>
            <Field label="Style">
              <select aria-label="Style" className={inputCls} value={s.errorStyle} onChange={(e) => up({ errorStyle: e.target.value as DvSpec['errorStyle'] })}>
                <option value="stop">Stop</option>
                <option value="warning">Warning</option>
                <option value="information">Information</option>
              </select>
            </Field>
            <Field label="Title">
              <input aria-label="Error title" className={inputCls} value={s.errorTitle ?? ''} onChange={(e) => up({ errorTitle: e.target.value || undefined })} />
            </Field>
            <Field label="Error message">
              <textarea aria-label="Error message" className={cx(inputCls, 'h-20 py-1.5')} value={s.error ?? ''} onChange={(e) => up({ error: e.target.value || undefined })} />
            </Field>
          </>
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

// ---------- grid overlay ----------

/** Dropdown arrow and list, input message and invalid-data circles for the active sheet. */
export function ValidationOverlay({ tab, book, L }: { tab: string; book: Book; L: Layout }) {
  const view = useViews((s) => s.views[tab])
  const [info, setInfo] = useState<DvInfo | null>(null)
  const [pick, setPick] = useState(0)
  const listRef = useRef<HTMLDivElement>(null)
  const { r, c } = view.sel.active
  const editing = !!view.editing
  useEffect(() => {
    let live = true
    void book.api.validationInfo(view.sheet, r, c).then((i: DvInfo | null) => live && setInfo(i))
    return () => {
      live = false
    }
  }, [book, view.sheet, r, c, book.version])
  // Moving to another cell closes the dropdown.
  useEffect(() => {
    if (getView(tab).listOpen) setView(tab, { listOpen: false })
  }, [view.sheet, r, c])
  const open = !!view.listOpen && !!info?.items?.length
  useEffect(() => {
    if (open) {
      const cur = book.cell(view.sheet, r, c)?.t ?? ''
      setPick(Math.max(0, info!.items!.indexOf(cur)))
      // Now (the list is rendered by this commit): a frame later, arrow keys pressed meanwhile were lost.
      listRef.current?.focus()
    }
  }, [open])
  const choose = async (text: string) => {
    setView(tab, { listOpen: false })
    const ctx = A.ctx()
    if (!ctx) return
    setView(tab, { editing: { text, where: 'cell', enterMode: true, caret: text.length } })
    await A.commitEdit({ ...ctx, view: getView(tab) }, 'none')
    refocus()
  }
  const x = colX(L, c)
  const y = rowY(L, r)
  const w = L.cols.size(c) * L.zoom
  const h = L.rows.size(r) * L.zoom
  const visible = w > 0 && h > 0 && x >= L.headerW - 1 && y >= L.headerH - 1 && x < L.width && y < L.height
  const circles = view.circles ?? []
  return (
    <>
      {circles.map((p) => {
        const cx0 = colX(L, p.c)
        const cy0 = rowY(L, p.r)
        const cw = L.cols.size(p.c) * L.zoom
        const ch = L.rows.size(p.r) * L.zoom
        if (!cw || !ch || cx0 + cw < L.headerW || cy0 + ch < L.headerH || cx0 > L.width || cy0 > L.height) return null
        return <div key={`${p.r},${p.c}`} data-testid="invalid-circle" className="pointer-events-none absolute z-[2] rounded-[50%] border-2 border-[#e0352b]" style={{ left: cx0 - 4, top: cy0 - 3, width: cw + 8, height: ch + 6 }} />
      })}
      {visible && info?.items?.length && !editing ? (
        <button
          type="button"
          aria-label="Show list"
          aria-haspopup="listbox"
          aria-expanded={open}
          data-testid="list-arrow"
          tabIndex={-1}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setView(tab, (v) => ({ listOpen: !v.listOpen }))}
          className="absolute z-[6] flex items-center justify-center rounded-sm border border-border bg-panel text-text-2 shadow-sm hover:bg-hover"
          style={{ left: x + w + 1, top: y + h - 18, width: 17, height: 17 }}
        >
          <ChevronDown className="size-3.5" />
        </button>
      ) : null}
      {visible && open && (
        <div
          ref={listRef}
          role="listbox"
          aria-label="Allowed values"
          tabIndex={0}
          aria-activedescendant={`dv-opt-${pick}`}
          className="absolute z-[7] max-h-[176px] overflow-y-auto rounded-md border border-border bg-panel py-1 text-[12.5px] shadow-lg outline-none"
          style={{ left: x, top: y + h + 1, minWidth: w + 18 }}
          onBlur={() => setView(tab, { listOpen: false })}
          onKeyDown={(e) => {
            const n = info!.items!.length
            if (e.key === 'ArrowDown') setPick((p) => Math.min(n - 1, p + 1))
            else if (e.key === 'ArrowUp') setPick((p) => Math.max(0, p - 1))
            else if (e.key === 'Enter' || e.key === 'Tab') void choose(info!.items![pick])
            else if (e.key === 'Escape') {
              setView(tab, { listOpen: false })
              refocus()
            } else return
            e.preventDefault()
            e.stopPropagation()
          }}
        >
          {info!.items!.map((t, i) => (
            <div key={i} id={`dv-opt-${i}`} role="option" aria-selected={i === pick} onMouseDown={(e) => e.preventDefault()} onClick={() => void choose(t)} className={cx('cursor-default truncate px-2.5 py-1', i === pick ? 'bg-accent-soft text-accent-ink' : 'hover:bg-hover')}>
              {t || ' '}
            </div>
          ))}
        </div>
      )}
      {visible && !open && !editing && info && info.showInput !== false && (info.prompt || info.promptTitle) && (
        <div role="tooltip" data-testid="input-message" className="pointer-events-none absolute z-[6] max-w-[260px] rounded-md border border-border bg-panel px-2.5 py-1.5 text-[12px] text-text-2 shadow-md" style={{ left: x + Math.min(w, 40), top: y + h + 6 }}>
          {info.promptTitle && <div className="mb-0.5 font-semibold text-text">{info.promptTitle}</div>}
          <div className="whitespace-pre-line">{info.prompt}</div>
        </div>
      )}
    </>
  )
}
