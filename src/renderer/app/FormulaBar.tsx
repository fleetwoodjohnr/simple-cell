// Name Box and formula bar.
import { useEffect, useRef, useState } from 'react'
import { Check, ChevronDown, ChevronUp, FunctionSquare, X } from 'lucide-react'
import { useShallow } from 'zustand/react/shallow'
import { books, set, toast, useApp } from '../state/store'
import { shortcutFor } from './commands'
import { getView, setView, useViews } from '../state/view'
import * as A from './actions'
import { IconBtn, cx } from '../ui'
import { useBookVersion } from './hooks'
import { useCaret } from '@/grid/Grid'
import { catalogEntry } from '@/core/engine/catalog'
import { CATALOG } from '@/core/engine/catalog'

export function FormulaBar({ tab }: { tab: string }) {
  const book = books.get(tab)!
  useBookVersion(book)
  const { sel, editing, sheet } = useViews(useShallow((s) => ({ sel: s.views[tab]?.sel, editing: s.views[tab]?.editing, sheet: s.views[tab]?.sheet })))
  const [cellText, setCellText] = useState('')
  const [nameText, setNameText] = useState<string | null>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const tall = useApp((s) => s.formulaBarTall)
  const c = A.ctx()
  const last = sel?.ranges[sel.ranges.length - 1]
  const dragging = !!last && (last.r0 !== last.r1 || last.c0 !== last.c1) && sel!.focus.r !== sel!.active.r
  const nameBox = c && c.tab === tab ? (nameText ?? (dragging && last ? A.selectionLabel(last) : A.nameBoxText(c))) : ''

  useEffect(() => {
    if (!sel || sheet === undefined) return
    let cancelled = false
    void book.api.editText(sheet, sel.active.r, sel.active.c).then((t) => !cancelled && setCellText(t))
    return () => {
      cancelled = true
    }
  }, [sel?.active.r, sel?.active.c, sheet, book.version])

  useCaret(inputRef, editing)
  const text = editing ? editing.text : cellText
  // Function tooltip: the innermost function whose parentheses contain the caret.
  const fnHint = editing?.text.startsWith('=') ? currentFunction(editing.text, editing.caret) : null

  return (
    <div className={cx('flex shrink-0 items-stretch gap-0 border-b border-border bg-panel text-[13px]', tall ? 'h-[92px]' : 'h-9')} data-testid="formula-bar">
      <input
        aria-label="Name Box"
        data-testid="name-box"
        className="w-[124px] shrink-0 border-r border-border bg-transparent px-2.5 font-medium text-text outline-none focus:bg-hover"
        value={nameBox}
        onFocus={(e) => {
          setNameText(nameBox)
          // After the click's mouseup (which would clear the selection). Only while still focused: select() also focuses,
          // and a late frame would otherwise pull focus back from the grid after Enter, losing what is typed next.
          const input = e.target
          requestAnimationFrame(() => document.activeElement === input && input.select())
        }}
        onChange={(e) => setNameText(e.target.value)}
        onBlur={() => setNameText(null)}
        onKeyDown={async (e) => {
          if (e.key === 'Escape') {
            setNameText(null)
            ;(e.target as HTMLInputElement).blur()
            document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus()
          }
          if (e.key !== 'Enter' || !c) return
          e.preventDefault()
          const t = (nameText ?? '').trim()
          if (!(await A.goTo(c, t))) {
            // A new name for the selection (Excel: typing a name in the Name Box defines it).
            if (/^[A-Za-z_\\][\w.]*$/.test(t)) {
              const r = last!
              const sh = book.sheetMeta(c.view.sheet)!.name
              const ref = `'${sh.replace(/'/g, "''")}'!${A.areaName({ ...r }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`
              const res = await book.api.defineName(t, '=' + ref, null)
              if (!res.ok) toast(res.error ?? 'That name could not be defined', 'error')
              else await book.refreshMeta()
            } else toast('Reference isn’t valid.', 'error')
          }
          setNameText(null)
          document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus()
        }}
      />
      <div className="flex items-center gap-0.5 border-r border-border px-1">
        <IconBtn label="Cancel" size="sm" disabled={!editing} onClick={() => c && A.cancelEdit(c)}>
          <X className="size-4" />
        </IconBtn>
        <IconBtn label="Enter" size="sm" disabled={!editing} onClick={() => c && void A.commitEdit(c, 'none')}>
          <Check className="size-4" />
        </IconBtn>
        <IconBtn label="Insert function" size="sm" onClick={() => set({ dialog: { kind: 'insert-function' } })}>
          <FunctionSquare className="size-4" />
        </IconBtn>
      </div>
      <div className="relative min-w-0 flex-1">
        <textarea
          ref={inputRef}
          aria-label="Formula bar"
          data-testid="formula-input"
          rows={1}
          spellCheck={false}
          className={cx('formula-input absolute inset-0 resize-none bg-transparent px-2.5 py-[7px] text-[13px] leading-[20px] text-text outline-none', tall ? 'overflow-y-auto whitespace-pre-wrap' : 'overflow-hidden whitespace-pre')}
          value={text}
          onFocus={() => {
            const cc = A.ctx()
            if (cc && !cc.view.editing) void A.startEdit(cc, undefined, 'bar').then(() => {
              const v = getView(tab).editing
              if (v) setView(tab, { editing: { ...v, where: 'bar', enterMode: false } })
            })
            else if (cc?.view.editing) setView(tab, { editing: { ...cc.view.editing, where: 'bar', enterMode: false } })
          }}
          onChange={(e) => {
            const cc = A.ctx()
            if (!cc) return
            const t = e.target.value
            if (!cc.view.editing) setView(tab, { editing: { text: t, where: 'bar', enterMode: false, caret: t.length } })
            else A.updateEdit(cc, t, e.target.selectionStart ?? t.length)
          }}
          onSelect={(e) => {
            const cc = A.ctx()
            const ed = cc?.view.editing
            const pos = (e.target as HTMLTextAreaElement).selectionStart ?? 0
            if (cc && ed && pos !== ed.caret) setView(tab, { editing: { ...ed, caret: pos } })
          }}
          onKeyDown={(e) => {
            const cc = A.ctx()
            if (!cc) return
            if (e.key === 'Escape') {
              e.preventDefault()
              A.cancelEdit(cc)
              document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus()
            } else if (e.key === 'Enter' && !e.altKey) {
              e.preventDefault()
              const mod = e.ctrlKey || e.metaKey
              void A.commitEdit(cc, mod ? 'none' : e.shiftKey ? 'up' : 'down', mod && e.shiftKey ? 'array' : mod ? 'all' : 'one').then((ok) => ok && document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus())
            } else if (e.key === 'Enter' && e.altKey) {
              e.preventDefault()
              const t = e.currentTarget
              const pos = t.selectionStart
              const next = text.slice(0, pos) + '\n' + text.slice(t.selectionEnd)
              A.updateEdit(cc, next, pos + 1)
            } else if (e.key === 'Tab') {
              e.preventDefault()
              void A.commitEdit(cc, e.shiftKey ? 'left' : 'right').then((ok) => ok && document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus())
            }
          }}
        />
        {fnHint && <FunctionHint name={fnHint.name} arg={fnHint.arg} />}
        {editing && <Autocomplete tab={tab} text={editing.text} caret={editing.caret} />}
      </div>
      <IconBtn label={tall ? 'Collapse formula bar' : 'Expand formula bar'} size="sm" shortcut={shortcutFor('view.formulaBarTall')} className="mx-1 self-start mt-1" onClick={() => set((x) => ({ formulaBarTall: !x.formulaBarTall }))}>
        {tall ? <ChevronUp className="size-4" /> : <ChevronDown className="size-4" />}
      </IconBtn>
    </div>
  )
}

/** The function name and argument index at the caret: "=ROUND(A1, |" → ROUND, 1. */
export function currentFunction(text: string, caret: number): { name: string; arg: number } | null {
  const stack: { name: string; arg: number }[] = []
  let inStr = false
  const src = text.slice(0, caret)
  for (let i = 0; i < src.length; i++) {
    const ch = src[i]
    if (ch === '"') inStr = !inStr
    if (inStr) continue
    if (ch === '(') {
      const m = /([A-Za-z_][\w.]*)$/.exec(src.slice(0, i))
      stack.push({ name: m ? m[1].toUpperCase() : '', arg: 0 })
    } else if (ch === ')') stack.pop()
    else if (ch === ',' && stack.length) stack[stack.length - 1].arg++
  }
  const top = [...stack].reverse().find((s) => s.name)
  return top ?? null
}

function FunctionHint({ name, arg }: { name: string; arg: number }) {
  const e = catalogEntry(name)
  if (!e) return null
  return (
    <div className="pointer-events-none absolute left-2 top-[38px] z-30 max-w-[520px] rounded-lg border border-border bg-elev px-2.5 py-1.5 text-[12px] shadow-[var(--shadow-pop)]" data-testid="function-hint">
      <span className="font-semibold">{e.n}</span>
      <span className="text-text-3"> — argument {arg + 1}</span>
      <div className="mt-0.5 text-text-2">{e.d}</div>
    </div>
  )
}

/** Function-name completion while typing a formula (Tab accepts, like Excel). */
function Autocomplete({ tab, text, caret }: { tab: string; text: string; caret: number }) {
  const m = /(?:^=|[=(,+\-*/^&<>:\s])([A-Za-z][\w.]*)$/.exec(text.slice(0, caret))
  if (!text.startsWith('=') || !m || m[1].length < 1) return null
  const prefix = m[1].toUpperCase()
  const list = CATALOG.filter((f) => f.n.startsWith(prefix)).slice(0, 8)
  if (!list.length || (list.length === 1 && list[0].n === prefix)) return null
  return (
    <ul className="absolute left-2 top-[38px] z-40 w-[360px] overflow-hidden rounded-lg border border-border bg-elev py-1 text-[12.5px] shadow-[var(--shadow-pop)]" role="listbox" aria-label="Functions" data-testid="autocomplete">
      {list.map((f) => (
        <li
          key={f.n}
          role="option"
          aria-selected={false}
          className={cx('cursor-default px-2.5 py-1 hover:bg-hover')}
          onMouseDown={(e) => {
            e.preventDefault()
            const c = A.ctx()
            if (!c || c.tab !== tab) return
            const start = caret - m[1].length
            const next = text.slice(0, start) + f.n + '(' + text.slice(caret)
            A.updateEdit(c, next, start + f.n.length + 1)
          }}
        >
          <span className="font-medium">{f.n}</span>
          <span className="ml-2 text-text-3">{f.d.slice(0, 60)}</span>
        </li>
      ))}
    </ul>
  )
}
