// Notes and hyperlinks in the UI: the note popup on hover, the link chip under the active cell, and their dialogs.
import { useEffect, useState } from 'react'
import { ExternalLink, Link2, Pencil, Trash2 } from 'lucide-react'
import type { Link, Note } from '@/core/model/workbook'
import { colX, rowY, type Layout } from '@/grid/layout'
import type { Book } from '../state/client'
import { get, set, toast, type Dialog } from '../state/store'
import { areasOf, useViews } from '../state/view'
import { Btn, Field, Modal, cx, inputCls } from '../ui'
import { isSafeExternalUrl } from '@/shared/urls'
import * as A from './actions'
import { markDirty } from '../state/documents'

const close = () => set({ dialog: null })
const refocus = () => requestAnimationFrame(() => document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')?.focus())

/** The grid reports the cell under the pointer (for note popups) without re-rendering itself. */
export function reportHover(tab: string, r: number, c: number) {
  window.dispatchEvent(new CustomEvent('sc:hover', { detail: { tab, r, c } }))
}

/** Follow a link: web and mail addresses open outside; places in the workbook are selected. */
export async function followLink(link: Link) {
  if (link.url) {
    if (!isSafeExternalUrl(link.url)) return toast('Only web (http, https) and email links can be opened.', 'error')
    try {
      await window.simpleCell.openExternal(link.url)
    } catch {
      toast('That link could not be opened.', 'error')
    }
    return
  }
  const c = A.ctx()
  if (c && link.location) await A.goTo(c, link.location)
}

export function NotesOverlay({ tab, book, L }: { tab: string; book: Book; L: Layout }) {
  const view = useViews((s) => s.views[tab])
  const [hover, setHover] = useState<{ r: number; c: number } | null>(null)
  const [note, setNote] = useState<(Note & { r: number; c: number }) | null>(null)
  const [link, setLink] = useState<Link | null>(null)
  useEffect(() => {
    const on = (e: Event) => {
      const d = (e as CustomEvent).detail
      if (d.tab === tab) setHover((h) => (h && h.r === d.r && h.c === d.c ? h : { r: d.r, c: d.c }))
    }
    window.addEventListener('sc:hover', on)
    return () => window.removeEventListener('sc:hover', on)
  }, [tab])
  // The note under the pointer.
  useEffect(() => {
    let live = true
    if (hover && book.cell(view.sheet, hover.r, hover.c)?.note) void book.api.noteAt(view.sheet, hover.r, hover.c).then((n) => live && setNote(n ? { ...n, ...hover } : null))
    else setNote(null)
    return () => {
      live = false
    }
  }, [hover?.r, hover?.c, view.sheet, book.version])
  // The active cell's link.
  const { r, c } = view.sel.active
  const hasLink = !!book.cell(view.sheet, r, c)?.link
  useEffect(() => {
    let live = true
    if (hasLink) void book.api.linkAt(view.sheet, r, c).then((l) => live && setLink(l))
    else setLink(null)
    return () => {
      live = false
    }
  }, [r, c, view.sheet, hasLink, book.version])
  const place = (rr: number, cc: number) => ({ x: colX(L, cc), y: rowY(L, rr), w: L.cols.size(cc) * L.zoom, h: L.rows.size(rr) * L.zoom })
  const onScreen = (p: { x: number; y: number; w: number; h: number }) => p.w > 0 && p.h > 0 && p.x + p.w > L.headerW && p.y + p.h > L.headerH && p.x < L.width && p.y < L.height
  const np = note ? place(note.r, note.c) : null
  const lp = link ? place(r, c) : null
  return (
    <>
      {note && np && onScreen(np) && (
        <div role="note" data-testid="note-popup" className="pointer-events-none absolute z-[7] w-[220px] whitespace-pre-wrap rounded-sm border border-[#b9a96b] bg-[#ffffe1] px-2 py-1.5 text-[12px] leading-snug text-[#1f1f1f] shadow-md" style={{ left: np.x + np.w + 8, top: np.y }}>
          {note.author && <div className="mb-0.5 font-bold">{note.author}:</div>}
          {note.text}
        </div>
      )}
      {link && lp && onScreen(lp) && !view.editing && (
        <div data-testid="link-chip" className="absolute z-[6] flex max-w-[360px] items-center gap-1 rounded-md border border-border bg-panel py-0.5 pl-2 pr-1 text-[12px] shadow-md" style={{ left: lp.x, top: lp.y + lp.h + 4 }}>
          <Link2 className="size-3.5 shrink-0 text-text-3" />
          <button type="button" className="truncate text-accent-ink underline-offset-2 hover:underline" title={link.tooltip ?? link.url ?? link.location} onClick={() => void followLink(link)}>
            {link.url?.replace(/^mailto:/, '') ?? link.location}
          </button>
          <button type="button" aria-label="Open link" className="rounded p-1 text-text-2 hover:bg-hover" onClick={() => void followLink(link)}>
            <ExternalLink className="size-3.5" />
          </button>
          <button type="button" aria-label="Edit link" className="rounded p-1 text-text-2 hover:bg-hover" onClick={() => set({ dialog: { kind: 'link' } })}>
            <Pencil className="size-3.5" />
          </button>
          <button
            type="button"
            aria-label="Remove link"
            className="rounded p-1 text-text-2 hover:bg-hover"
            onClick={() => {
              const ctx = A.ctx()
              if (ctx) void A.edit(ctx, (b) => b.api.setLink(areasOf(ctx.view), null))
            }}
          >
            <Trash2 className="size-3.5" />
          </button>
        </div>
      )}
    </>
  )
}

export function NoteDialog() {
  const c = A.ctx()
  const [text, setText] = useState('')
  const [had, setHad] = useState(false)
  useEffect(() => {
    if (!c) return
    void c.book.api.noteAt(c.view.sheet, c.view.sel.active.r, c.view.sel.active.c).then((n) => {
      setText(n?.text ?? '')
      setHad(!!n)
    })
  }, [])
  if (!c) return null
  const { r, c: col } = c.view.sel.active
  const save = async (value: string | null) => {
    await c.book.api.setNote(c.view.sheet, r, col, value, get().settings.authorName)
    markDirty(c.tab)
    close()
    refocus()
  }
  return (
    <Modal
      open
      onOpenChange={(o) => !o && (close(), refocus())}
      title={had ? 'Edit note' : 'New note'}
      width={420}
      footer={
        <>
          {had && (
            <Btn className="mr-auto" onClick={() => void save(null)}>
              Delete note
            </Btn>
          )}
          <Btn onClick={() => (close(), refocus())}>Cancel</Btn>
          <Btn variant="primary" onClick={() => void save(text.trim() ? text : null)}>
            Save
          </Btn>
        </>
      }
    >
      <textarea aria-label="Note" autoFocus className={cx(inputCls, 'h-32 py-2')} value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && (e.ctrlKey || e.metaKey) && void save(text.trim() ? text : null)} />
      {!get().settings.authorName && <p className="mt-2 text-[11.5px] text-text-3">Notes are signed with the author name from Settings ▸ General.</p>}
    </Modal>
  )
}

export function LinkDialog({ d }: { d: Extract<Dialog, { kind: 'link' }> }) {
  void d
  const c = A.ctx()
  const [address, setAddress] = useState('')
  const [display, setDisplay] = useState('')
  const [tip, setTip] = useState('')
  const [had, setHad] = useState(false)
  useEffect(() => {
    if (!c) return
    const { r, c: col } = c.view.sel.active
    void c.book.api.linkAt(c.view.sheet, r, col).then((l) => {
      setAddress(l?.url?.replace(/^mailto:/, 'mailto:') ?? l?.location ?? '')
      setTip(l?.tooltip ?? '')
      setHad(!!l)
    })
    setDisplay(c.book.cell(c.view.sheet, r, col)?.t ?? '')
  }, [])
  if (!c) return null
  const ok = async () => {
    const a = address.trim()
    if (!a) return
    // An address with a scheme or a dot-domain goes outside; anything else is a place in the workbook.
    const outside = /^[a-z][a-z0-9+.-]*:/i.test(a) || /^www\./i.test(a) || /^[^\s!]+\.[a-z]{2,}(\/|$)/i.test(a)
    const link: Link = outside ? { url: /^[a-z][a-z0-9+.-]*:/i.test(a) ? a : /@/.test(a) && !a.includes('/') ? `mailto:${a}` : `https://${a}` } : { location: a.replace(/^#/, '') }
    if (tip.trim()) link.tooltip = tip.trim()
    const cur = c.book.cell(c.view.sheet, c.view.sel.active.r, c.view.sel.active.c)?.t ?? ''
    await A.edit(c, (b) => b.api.setLink(areasOf(c.view), link, display && display !== cur ? display : cur ? undefined : display || undefined))
    close()
    refocus()
  }
  return (
    <Modal
      open
      onOpenChange={(o) => !o && (close(), refocus())}
      title={had ? 'Edit link' : 'Insert link'}
      width={460}
      footer={
        <>
          {had && (
            <Btn className="mr-auto" onClick={() => void A.edit(c, (b) => b.api.setLink(areasOf(c.view), null)).then(() => (close(), refocus()))}>
              Remove link
            </Btn>
          )}
          <Btn onClick={() => (close(), refocus())}>Cancel</Btn>
          <Btn variant="primary" onClick={() => void ok()}>
            OK
          </Btn>
        </>
      }
    >
      <div className="space-y-3">
        <Field label="Text to display">
          <input aria-label="Text to display" className={inputCls} value={display} onChange={(e) => setDisplay(e.target.value)} />
        </Field>
        <Field label="Address" hint="A web address (example.com/page), an email address, or a place in this workbook such as Sheet2!A1 or a defined name">
          <input aria-label="Address" autoFocus className={inputCls} value={address} onChange={(e) => setAddress(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && void ok()} />
        </Field>
        <Field label="ScreenTip">
          <input aria-label="ScreenTip" className={inputCls} value={tip} onChange={(e) => setTip(e.target.value)} />
        </Field>
      </div>
    </Modal>
  )
}
