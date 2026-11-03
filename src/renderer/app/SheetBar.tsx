// Sheet tabs (bottom left) and the status bar with quick calculations and zoom (bottom right).
import { useEffect, useRef, useState } from 'react'
import { ChevronLeft, ChevronRight, Plus, Minus } from 'lucide-react'
import { useShallow } from 'zustand/react/shallow'
import { books, toast } from '../state/store'
import { setView, useViews, areasOf } from '../state/view'
import { markDirty, switchSheet } from '../state/documents'
import { runCommand } from './commands'
import { useBookVersion } from './hooks'
import { COLORS, CtxContent, CtxItem, CtxMenu, CtxSep, CtxSub, CtxTrigger, cx } from '../ui'
import type { Stats } from '@/core/host'
import { generalNumber } from '@/core/numfmt/format'
import { colName } from '@/core/refs'

export function SheetTabs({ tab }: { tab: string }) {
  const book = books.get(tab)!
  useBookVersion(book)
  const sheet = useViews((s) => s.views[tab]?.sheet)
  const [renaming, setRenaming] = useState<number | null>(null)
  const [dragId, setDragId] = useState<number | null>(null)
  const scroller = useRef<HTMLDivElement>(null)
  const sheets = book.meta.sheets
  const hidden = sheets.filter((s) => s.hidden)

  const rename = async (id: number, name: string) => {
    setRenaming(null)
    const cur = sheets.find((s) => s.id === id)
    if (!cur || cur.name === name) return
    const r = await book.api.renameSheet(id, name)
    if (!r.ok) toast(r.error ?? 'Couldn’t rename the sheet', 'error')
    else markDirty(tab)
  }
  const act = async (p: Promise<{ ok: boolean; error?: string }>) => {
    const r = await p
    if (!r.ok) toast(r.error ?? 'That didn’t work', 'error')
    else markDirty(tab)
  }

  return (
    <div className="flex min-w-0 flex-1 items-center gap-0.5" role="tablist" aria-label="Sheets" data-testid="sheet-tabs">
      <button aria-label="Scroll sheet tabs left" className="flex size-7 items-center justify-center rounded text-text-3 hover:bg-hover" onClick={() => scroller.current?.scrollBy({ left: -120 })}>
        <ChevronLeft className="size-4" />
      </button>
      <button aria-label="Scroll sheet tabs right" className="flex size-7 items-center justify-center rounded text-text-3 hover:bg-hover" onClick={() => scroller.current?.scrollBy({ left: 120 })}>
        <ChevronRight className="size-4" />
      </button>
      <div ref={scroller} className="flex min-w-0 items-center gap-0.5 overflow-x-auto [scrollbar-width:none]">
        {sheets
          .filter((s) => !s.hidden)
          .map((s) => (
            <CtxMenu key={s.id}>
              <CtxTrigger asChild>
                <div
                  role="tab"
                  tabIndex={0}
                  aria-selected={s.id === sheet}
                  data-testid={`sheet-tab-${s.name}`}
                  draggable={renaming !== s.id}
                  onDragStart={() => setDragId(s.id)}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={() => {
                    if (dragId === null || dragId === s.id) return
                    void act(book.api.moveSheet(dragId, sheets.findIndex((x) => x.id === s.id)))
                    setDragId(null)
                  }}
                  onMouseDown={() => switchSheet(tab, s.id)}
                  onDoubleClick={() => setRenaming(s.id)}
                  onKeyDown={(e) => e.key === 'Enter' && switchSheet(tab, s.id)}
                  className={cx(
                    'relative flex h-7 shrink-0 cursor-default items-center rounded-md px-3 text-[12.5px] transition-colors',
                    s.id === sheet ? 'bg-panel font-semibold text-accent-ink shadow-[0_0_0_1px_var(--border)]' : 'text-text-2 hover:bg-hover'
                  )}
                >
                  {s.tabColor && <span className="absolute inset-x-2 bottom-0.5 h-[3px] rounded-full" style={{ background: s.tabColor }} />}
                  {renaming === s.id ? (
                    <input
                      autoFocus
                      aria-label="Sheet name"
                      defaultValue={s.name}
                      className="w-28 bg-transparent outline-none"
                      onFocus={(e) => e.target.select()}
                      onBlur={(e) => void rename(s.id, e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') void rename(s.id, (e.target as HTMLInputElement).value)
                        if (e.key === 'Escape') setRenaming(null)
                        e.stopPropagation()
                      }}
                    />
                  ) : (
                    s.name
                  )}
                </div>
              </CtxTrigger>
              <CtxContent>
                <CtxItem onSelect={() => void runCommand('sheet.insert')}>Insert sheet</CtxItem>
                <CtxItem onSelect={() => void act(book.api.deleteSheet(s.id))}>Delete</CtxItem>
                <CtxItem onSelect={() => setRenaming(s.id)}>Rename</CtxItem>
                <CtxSub label="Move">
                  <CtxItem onSelect={() => void act(book.api.moveSheet(s.id, Math.max(0, sheets.findIndex((x) => x.id === s.id) - 1)))}>Move left</CtxItem>
                  <CtxItem onSelect={() => void act(book.api.moveSheet(s.id, Math.min(sheets.length - 1, sheets.findIndex((x) => x.id === s.id) + 1)))}>Move right</CtxItem>
                </CtxSub>
                <CtxSub label="Tab color">
                  <div className="grid grid-cols-6 gap-1 p-2">
                    {COLORS.flat().map((col) => (
                      <button key={col} aria-label={col} className="size-5 rounded border border-black/10" style={{ background: col }} onClick={() => void act(book.api.sheetProps(s.id, { tabColor: col }))} />
                    ))}
                  </div>
                  <CtxItem onSelect={() => void act(book.api.sheetProps(s.id, { tabColor: null }))}>No color</CtxItem>
                </CtxSub>
                <CtxSep />
                <CtxItem onSelect={() => void act(book.api.sheetProps(s.id, { hidden: true }))}>Hide</CtxItem>
                {hidden.length > 0 && (
                  <CtxSub label="Unhide">
                    {hidden.map((h) => (
                      <CtxItem key={h.id} onSelect={() => void act(book.api.sheetProps(h.id, { hidden: false }))}>
                        {h.name}
                      </CtxItem>
                    ))}
                  </CtxSub>
                )}
              </CtxContent>
            </CtxMenu>
          ))}
      </div>
      <button aria-label="New sheet" title="New sheet (Shift+F11)" data-testid="add-sheet" className="ml-1 flex size-7 shrink-0 items-center justify-center rounded-md text-text-2 hover:bg-hover hover:text-text" onClick={() => void runCommand('sheet.insert')}>
        <Plus className="size-4" />
      </button>
    </div>
  )
}

const fmt = (n: number | null) => (n === null ? '' : generalNumber(Number(n.toPrecision(15)), 14))

export function StatusBar({ tab }: { tab: string }) {
  const book = books.get(tab)!
  const version = useBookVersion(book)
  const { view } = useViews(useShallow((s) => ({ view: s.views[tab] })))
  const [stats, setStats] = useState<Stats | null>(null)
  useEffect(() => {
    if (!view) return
    const t = setTimeout(async () => setStats(await book.api.stats(areasOf(view))), 60)
    return () => clearTimeout(t)
  }, [view?.sel, view?.sheet, version])
  const multi = view && (view.sel.ranges.length > 1 || view.sel.ranges[0].r0 !== view.sel.ranges[0].r1 || view.sel.ranges[0].c0 !== view.sel.ranges[0].c1)
  const zoom = Math.round((view?.zoom ?? 1) * 100)
  const editing = view?.editing
  const circ = book.meta.circular[0]
  return (
    <div className="flex h-8 shrink-0 items-center gap-3 border-t border-border bg-bg pl-2 pr-3 text-[12px] text-text-2" data-testid="status-bar">
      <SheetTabs tab={tab} />
      <span className="shrink-0 text-text-3" data-testid="mode">
        {editing ? (editing.enterMode ? 'Enter' : 'Edit') : 'Ready'}
        {book.meta.calcMode === 'manual' && ' · Manual calculation'}
      </span>
      {circ && (
        <span className="shrink-0 font-medium text-warn" data-testid="circular">
          Circular references: {colName(circ.c)}
          {circ.r + 1}
        </span>
      )}
      {multi && stats && stats.count > 0 && (
        <span className="flex shrink-0 gap-3 tabular-nums" data-testid="quick-stats">
          {stats.numbers > 0 && <span>Average: {fmt(stats.average)}</span>}
          <span>Count: {stats.count}</span>
          {stats.numbers > 0 && <span>Sum: {fmt(stats.sum)}</span>}
          {stats.numbers > 1 && <span>Min: {fmt(stats.min)}</span>}
          {stats.numbers > 1 && <span>Max: {fmt(stats.max)}</span>}
        </span>
      )}
      <div className="flex shrink-0 items-center gap-1.5">
        <button aria-label="Zoom out" className="flex size-6 items-center justify-center rounded hover:bg-hover" onClick={() => setView(tab, (v) => ({ zoom: Math.max(0.1, Math.round((v.zoom - 0.1) * 10) / 10) }))}>
          <Minus className="size-3.5" />
        </button>
        <input aria-label="Zoom" type="range" min={10} max={400} step={5} value={zoom} onChange={(e) => setView(tab, { zoom: Number(e.target.value) / 100 })} className="w-24 accent-[var(--accent)]" />
        <button aria-label="Zoom in" className="flex size-6 items-center justify-center rounded hover:bg-hover" onClick={() => setView(tab, (v) => ({ zoom: Math.min(4, Math.round((v.zoom + 0.1) * 10) / 10) }))}>
          <Plus className="size-3.5" />
        </button>
        <button className="w-10 rounded text-right tabular-nums hover:bg-hover" onClick={() => setView(tab, { zoom: 1 })} aria-label="Reset zoom to 100%">
          {zoom}%
        </button>
      </div>
    </div>
  )
}
