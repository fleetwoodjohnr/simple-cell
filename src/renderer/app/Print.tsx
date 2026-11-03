// Printing and PDF: pages built from the core's pagination, drawn as DOM, and a Print dialog with a preview.
// The pages render into a host outside the app root that only shows when printing (see styles.css).
import { useEffect, useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import { create } from 'zustand'
import type { ViewCell } from '@/core/doc'
import type { SheetMeta } from '@/core/host'
import type { PageSetup, PrintJob, PrintPage } from '@/core/print'
import { PAPER, headerFooterText } from '@/core/print'
import type { Area } from '@/core/engine/values'
import type { BorderSide, Style } from '@/core/model/styles'
import { colName } from '@/core/refs'
import type { Book } from '../state/client'
import { get, set, toast } from '../state/store'
import { getView } from '../state/view'
import { Btn, Field, Modal, Segmented, cx, inputCls } from '../ui'
import * as A from './actions'

interface Prepared {
  job: PrintJob
  meta: SheetMeta
  name: string
  cells: Map<number, ViewCell>
}

const usePrint = create<{ docs: Prepared[] | null }>(() => ({ docs: null }))

export type PrintScope = 'sheet' | 'workbook' | 'selection'

/** Pagination and cell contents for the scope. */
export async function preparePrint(book: Book, tab: string, scope: PrintScope): Promise<Prepared[]> {
  const view = getView(tab)
  const sheets = scope === 'workbook' ? book.meta.sheets.filter((s) => !s.hidden).map((s) => s.id) : [view.sheet]
  const out: Prepared[] = []
  for (const id of sheets) {
    const sel = view.sel.ranges[view.sel.ranges.length - 1]
    const area: Area | undefined = scope === 'selection' ? { sheet: id, ...sel } : undefined
    const job = await book.api.printJob(id, area)
    if (!job || !job.pages.length) continue
    const meta = book.sheetMeta(id)!
    const rows = job.pages.flatMap((p) => [...p.rows, ...p.titleRows])
    const cols = job.pages.flatMap((p) => [...p.cols, ...p.titleCols])
    const cells = new Map<number, ViewCell>()
    if (rows.length && cols.length) {
      // ponytail: one viewport request per sheet; print areas beyond ~200k cells are slow.
      const list = await book.api.viewport(id, Math.min(...rows), Math.max(...rows), Math.min(...cols), Math.max(...cols))
      for (const v of list) cells.set(v.r * 16384 + v.c, v)
    }
    out.push({ job, meta, name: meta.name, cells })
  }
  return out
}

const sizeOf = (list: [number, number][], dflt: number) => {
  const m = new Map(list)
  return (i: number) => m.get(i) ?? dflt
}

const borderCss = (b?: BorderSide) => (b ? `${b.style === 'thick' ? 3 : b.style === 'medium' ? 2 : 1}px ${b.style === 'double' ? 'double' : /dash/i.test(b.style) ? 'dashed' : /dot|hair/i.test(b.style) ? 'dotted' : 'solid'} ${b.color ?? '#000'}` : undefined)

/** One printed page at full size (96 px per inch). */
export function PageView({ p, page, total, styles, defaultFont, defaultSize, file }: { p: Prepared; page: PrintPage; total: number; styles: Style[]; defaultFont: string; defaultSize: number; file: string }) {
  const { job, meta } = p
  const s: PageSetup = job.setup
  const m = s.margins
  const colW = sizeOf(meta.cols, meta.defaultColWidth)
  const rowH = sizeOf(meta.rows, meta.defaultRowHeight)
  const cols = [...page.titleCols, ...page.cols]
  const rows = [...page.titleRows, ...page.rows]
  const head = s.headings ? { w: 32, h: 20 } : { w: 0, h: 0 }
  const xs = new Map<number, number>()
  let x = head.w
  for (const c of cols) {
    xs.set(c, x)
    x += colW(c)
  }
  const ys = new Map<number, number>()
  let y = head.h
  for (const r of rows) {
    ys.set(r, y)
    y += rowH(r)
  }
  const contentW = x
  const contentH = y
  const availW = job.paper.w - (m.left + m.right) * 96
  const availH = job.paper.h - (m.top + m.bottom) * 96
  const offX = s.centerH ? Math.max(0, (availW - contentW * job.scale) / 2) : 0
  const offY = s.centerV ? Math.max(0, (availH - contentH * job.scale) / 2) : 0
  const hf = (code: string) => headerFooterText(code, { page: page.number, pages: total, sheet: p.name, file, date: new Date() })
  const header = hf(s.header)
  const footer = hf(s.footer)
  const covered = new Set<number>()
  const merges = meta.merges.filter((mg) => cols.includes(mg.c0) && rows.includes(mg.r0))
  for (const mg of meta.merges) for (let r = mg.r0; r <= mg.r1; r++) for (let c = mg.c0; c <= mg.c1; c++) if (r !== mg.r0 || c !== mg.c0) covered.add(r * 16384 + c)
  const band = (text: { l: string; c: string; r: string }, top: boolean) =>
    text.l || text.c || text.r ? (
      <div className="absolute flex text-[9pt] text-black" style={{ left: m.left * 96, right: m.right * 96, [top ? 'top' : 'bottom']: (top ? m.header : m.footer) * 96 }}>
        <span className="flex-1 text-left">{text.l}</span>
        <span className="flex-1 text-center">{text.c}</span>
        <span className="flex-1 text-right">{text.r}</span>
      </div>
    ) : null
  return (
    <div className="print-page relative overflow-hidden bg-white text-black" style={{ width: job.paper.w, height: job.paper.h }}>
      {band(header, true)}
      <div className="absolute overflow-hidden" style={{ left: m.left * 96 + offX, top: m.top * 96 + offY, width: availW - offX, height: availH - offY }}>
        <div className="relative origin-top-left" style={{ width: contentW, height: contentH, transform: `scale(${job.scale})` }}>
          {s.headings && (
            <>
              {cols.map((c) => (
                <div key={`h${c}`} className="absolute border border-[#bbb] bg-[#f3f3f3] text-center text-[8pt] leading-[19px]" style={{ left: xs.get(c), top: 0, width: colW(c), height: head.h }}>
                  {colName(c)}
                </div>
              ))}
              {rows.map((r) => (
                <div key={`v${r}`} className="absolute border border-[#bbb] bg-[#f3f3f3] text-center text-[8pt]" style={{ left: 0, top: ys.get(r), width: head.w, height: rowH(r), lineHeight: `${rowH(r)}px` }}>
                  {r + 1}
                </div>
              ))}
            </>
          )}
          {rows.flatMap((r) =>
            cols.map((c) => {
              const k = r * 16384 + c
              if (covered.has(k)) return null
              const v = p.cells.get(k)
              const mg = merges.find((g) => g.r0 === r && g.c0 === c)
              const w = mg ? cols.filter((cc) => cc >= mg.c0 && cc <= mg.c1).reduce((a, cc) => a + colW(cc), 0) : colW(c)
              const h = mg ? rows.filter((rr) => rr >= mg.r0 && rr <= mg.r1).reduce((a, rr) => a + rowH(rr), 0) : rowH(r)
              const st: Style = v?.x ?? styles[v?.s ?? 0] ?? {}
              if (!v && !s.gridlines) return null
              const align = st.hAlign && st.hAlign !== 'general' ? (st.hAlign === 'centerContinuous' ? 'center' : st.hAlign) : v?.k === 'n' || v?.k === 'd' ? 'right' : v?.k === 'b' || v?.k === 'e' ? 'center' : 'left'
              const grid = s.gridlines ? '1px solid #d4d4d4' : undefined
              const text = v?.hide ? '' : (v?.t ?? '')
              return (
                <div
                  key={k}
                  className="absolute"
                  style={{
                    left: xs.get(c),
                    top: ys.get(r),
                    width: w,
                    height: h,
                    background: st.fill,
                    borderTop: borderCss(st.borderTop) ?? grid,
                    borderLeft: borderCss(st.borderLeft) ?? grid,
                    borderBottom: borderCss(st.borderBottom),
                    borderRight: borderCss(st.borderRight),
                    display: 'flex',
                    alignItems: st.vAlign === 'top' ? 'flex-start' : st.vAlign === 'center' ? 'center' : 'flex-end',
                    justifyContent: align === 'right' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start',
                    padding: '0 3px',
                    fontFamily: `"${st.font ?? defaultFont}", Carlito, Calibri, sans-serif`,
                    fontSize: `${st.size ?? defaultSize}pt`,
                    fontWeight: st.bold ? 700 : 400,
                    fontStyle: st.italic ? 'italic' : undefined,
                    textDecoration: [st.underline && 'underline', st.strike && 'line-through'].filter(Boolean).join(' ') || undefined,
                    color: v?.fc ?? st.color ?? '#000',
                    whiteSpace: st.wrap ? 'pre-wrap' : 'pre',
                    overflow: st.wrap || align !== 'left' || v?.k !== 's' ? 'hidden' : 'visible',
                    lineHeight: 1.2
                  }}
                >
                  {v?.bar && <span className="absolute left-[2px] top-[2px] bottom-[3px]" style={{ width: `calc(${v.bar.w * 100}% - 4px)`, background: `linear-gradient(to right, ${v.bar.color}, #fff)` }} />}
                  <span className="relative">{text}</span>
                </div>
              )
            })
          )}
        </div>
      </div>
      {band(footer, false)}
    </div>
  )
}

/** The pages for printing/PDF, rendered outside the app root. */
export function PrintHost() {
  const docs = usePrint((s) => s.docs)
  const settings = get().settings
  if (!docs) return null
  const total = docs.reduce((n, d) => n + d.job.pages.length, 0)
  const file = get().tabs.find((t) => t.id === get().active)?.title ?? ''
  const styles = (A.ctx()?.book.meta.styles ?? []) as Style[]
  const first = docs[0]?.job
  return createPortal(
    <div className="print-host">
      {first && <style>{`@page { size: ${first.paper.w / 96}in ${first.paper.h / 96}in; margin: 0 } .print-page { break-after: page; page-break-after: always }`}</style>}
      {docs.flatMap((d) => d.job.pages.map((pg) => <PageView key={`${d.job.sheet}:${pg.number}`} p={d} page={pg} total={total} styles={styles} defaultFont={settings.defaultFont} defaultSize={settings.defaultFontSize} file={file} />))}
    </div>,
    document.body
  )
}

/** Render the pages, then save a PDF (or open the system print dialog). */
export async function runPrint(mode: 'pdf' | 'print', scope: PrintScope) {
  const c = A.ctx()
  if (!c) return
  const docs = await preparePrint(c.book, c.tab, scope)
  if (!docs.length) return toast('There’s nothing to print: the sheet is empty.', 'info')
  usePrint.setState({ docs })
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))
  await document.fonts.ready
  try {
    if (mode === 'print') {
      await window.simpleCell.print()
      return
    }
    const title = get().tabs.find((t) => t.id === c.tab)?.title ?? 'Workbook'
    const data = await window.simpleCell.exportPdf({ title })
    const path = await window.simpleCell.saveDialog(`${title}.pdf`, ['pdf'])
    if (!path) return
    await window.simpleCell.writeFile(path, data)
    toast('Saved as PDF.', 'success')
  } catch (e) {
    toast(`Printing failed: ${(e as Error).message}`, 'error')
  } finally {
    usePrint.setState({ docs: null })
  }
}

const SCALES: [string, string][] = [
  ['actual', 'Actual size'],
  ['sheet', 'Fit sheet on one page'],
  ['cols', 'Fit all columns on one page'],
  ['rows', 'Fit all rows on one page']
]

export function PrintDialog() {
  const c = A.ctx()
  const [scope, setScope] = useState<PrintScope>('sheet')
  const [setup, setSetup] = useState<(PageSetup & { area: Area | null }) | null>(null)
  const [docs, setDocs] = useState<Prepared[] | null>(null)
  const [version, setVersion] = useState(0)
  useEffect(() => {
    if (c) void c.book.api.pageSetup(c.view.sheet).then((s) => setSetup(s as PageSetup & { area: Area | null }))
  }, [version])
  useEffect(() => {
    if (c) void preparePrint(c.book, c.tab, scope).then(setDocs)
  }, [scope, version])
  const total = docs?.reduce((n, d) => n + d.job.pages.length, 0) ?? 0
  const styles = (c?.book.meta.styles ?? []) as Style[]
  const settings = get().settings
  const file = get().tabs.find((t) => t.id === c?.tab)?.title ?? ''
  const scaleMode = useMemo(() => (!setup?.fitToPage ? 'actual' : setup.fitWidth === 1 && setup.fitHeight === 1 ? 'sheet' : setup.fitWidth === 1 ? 'cols' : 'rows'), [setup])
  if (!c || !setup) return null
  const change = async (patch: Partial<PageSetup>) => {
    await A.edit(c, (b) => b.api.setPageSetup(c.view.sheet, patch))
    setVersion((v) => v + 1)
  }
  const go = (mode: 'pdf' | 'print') => {
    set({ dialog: null })
    void runPrint(mode, scope)
  }
  const thumb = 0.32
  return (
    <Modal
      open
      onOpenChange={(o) => !o && set({ dialog: null })}
      title="Print"
      width={920}
      footer={
        <>
          <span className="mr-auto text-[12px] text-text-3" data-testid="print-pages">
            {total} {total === 1 ? 'page' : 'pages'}
          </span>
          <Btn onClick={() => set({ dialog: null })}>Cancel</Btn>
          <Btn onClick={() => go('print')}>Print…</Btn>
          <Btn variant="primary" onClick={() => go('pdf')}>
            Save as PDF
          </Btn>
        </>
      }
    >
      <div className="flex gap-4">
        <div className="w-[260px] shrink-0 space-y-3 text-[12.5px]">
          <Field label="Print">
            <select aria-label="What to print" className={inputCls} value={scope} onChange={(e) => setScope(e.target.value as PrintScope)}>
              <option value="sheet">Active sheet</option>
              <option value="workbook">Entire workbook</option>
              <option value="selection">Selection</option>
            </select>
          </Field>
          <Field label="Orientation">
            <Segmented label="Orientation" value={setup.orientation} onChange={(v) => void change({ orientation: v })} options={[{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }]} />
          </Field>
          <Field label="Paper size">
            <select aria-label="Paper size" className={inputCls} value={setup.paper} onChange={(e) => void change({ paper: Number(e.target.value) })}>
              {Object.entries(PAPER).map(([k, p]) => (
                <option key={k} value={k}>
                  {p.name} ({p.w}″ × {p.h}″)
                </option>
              ))}
            </select>
          </Field>
          <Field label="Margins">
            <select
              aria-label="Margins"
              className={inputCls}
              value={setup.margins.left === 0.7 && setup.margins.top === 0.75 ? 'normal' : setup.margins.left === 0.25 ? 'narrow' : setup.margins.left === 1 ? 'wide' : 'custom'}
              onChange={(e) => {
                const v = e.target.value
                const mm = v === 'narrow' ? { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } : v === 'wide' ? { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 } : { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
                void change({ margins: mm })
              }}
            >
              <option value="normal">Normal</option>
              <option value="wide">Wide</option>
              <option value="narrow">Narrow</option>
              <option value="custom" disabled>
                Custom
              </option>
            </select>
          </Field>
          <Field label="Scaling">
            <select
              aria-label="Scaling"
              className={inputCls}
              value={scaleMode}
              onChange={(e) => {
                const v = e.target.value
                void change(v === 'actual' ? { fitToPage: false, scale: 100 } : v === 'sheet' ? { fitToPage: true, fitWidth: 1, fitHeight: 1 } : v === 'cols' ? { fitToPage: true, fitWidth: 1, fitHeight: 0 } : { fitToPage: true, fitWidth: 0, fitHeight: 1 })
              }}
            >
              {SCALES.map(([v, l]) => (
                <option key={v} value={v}>
                  {l}
                </option>
              ))}
            </select>
          </Field>
          <label className="flex items-center gap-2">
            <input type="checkbox" checked={setup.gridlines} onChange={(e) => void change({ gridlines: e.target.checked })} /> Print gridlines
          </label>
          <label className="flex items-center gap-2">
            <input type="checkbox" checked={setup.headings} onChange={(e) => void change({ headings: e.target.checked })} /> Print row and column headings
          </label>
          <Field label="Header" hint="&P page, &N pages, &A sheet, &F file, &D date; &L &C &R for left, centre, right">
            <input aria-label="Header" className={inputCls} defaultValue={setup.header} onBlur={(e) => e.target.value !== setup.header && void change({ header: e.target.value })} />
          </Field>
          <Field label="Footer">
            <input aria-label="Footer" className={inputCls} defaultValue={setup.footer} onBlur={(e) => e.target.value !== setup.footer && void change({ footer: e.target.value })} />
          </Field>
        </div>
        <div className="h-[520px] flex-1 overflow-y-auto rounded-lg bg-hover p-3" data-testid="print-preview">
          {!docs && <p className="text-[12.5px] text-text-3">Preparing pages…</p>}
          {docs && !total && <p className="text-[12.5px] text-text-3">There’s nothing to print.</p>}
          <div className="flex flex-wrap gap-3">
            {docs?.flatMap((d) =>
              d.job.pages.map((pg) => (
                <div key={`${d.job.sheet}:${pg.number}`} className={cx('overflow-hidden shadow-md')} style={{ width: d.job.paper.w * thumb, height: d.job.paper.h * thumb }}>
                  <div className="origin-top-left" style={{ transform: `scale(${thumb})` }}>
                    <PageView p={d} page={pg} total={total} styles={styles} defaultFont={settings.defaultFont} defaultSize={settings.defaultFontSize} file={file} />
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </Modal>
  )
}
