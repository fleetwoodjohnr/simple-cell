// Home ▸ Cell Styles: Excel's built-in styles, applied as formatting (fonts, fills, borders, number formats).
import { useState } from 'react'
import { ChevronDown, Palette } from 'lucide-react'
import type { BorderSide, Style } from '@/core/model/styles'
import { Pop, PopContent, PopTrigger, cx } from '../ui'
import * as A from './actions'

type Cat = 'font' | 'fill' | 'border' | 'number'
const KEYS: Record<Cat, (keyof Style)[]> = {
  font: ['bold', 'italic', 'underline', 'strike', 'color', 'size', 'font'],
  fill: ['fill', 'pattern'],
  border: ['borderTop', 'borderBottom', 'borderLeft', 'borderRight'],
  number: ['numFmt']
}

const thin = (color: string): BorderSide => ({ style: 'thin', color })
const box = (b: BorderSide) => ({ borderTop: b, borderBottom: b, borderLeft: b, borderRight: b })
const ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
const tint = (hex: string, t: number) => '#' + [1, 3, 5].map((i) => Math.round(parseInt(hex.slice(i, i + 2), 16) + (255 - parseInt(hex.slice(i, i + 2), 16)) * t).toString(16).padStart(2, '0')).join('').toUpperCase()

export interface CellStyle {
  name: string
  group: string
  style: Style
  cats: Cat[]
}

export const CELL_STYLES: CellStyle[] = [
  { name: 'Normal', group: 'Good, bad and neutral', style: {}, cats: ['font', 'fill', 'border', 'number'] },
  { name: 'Bad', group: 'Good, bad and neutral', style: { fill: '#FFC7CE', color: '#9C0006' }, cats: ['font', 'fill'] },
  { name: 'Good', group: 'Good, bad and neutral', style: { fill: '#C6EFCE', color: '#006100' }, cats: ['font', 'fill'] },
  { name: 'Neutral', group: 'Good, bad and neutral', style: { fill: '#FFEB9C', color: '#9C5700' }, cats: ['font', 'fill'] },
  { name: 'Calculation', group: 'Data and model', style: { fill: '#F2F2F2', color: '#FA7D00', bold: true, ...box(thin('#7F7F7F')) }, cats: ['font', 'fill', 'border'] },
  { name: 'Check cell', group: 'Data and model', style: { fill: '#A5A5A5', color: '#FFFFFF', bold: true, ...box({ style: 'double', color: '#3F3F3F' }) }, cats: ['font', 'fill', 'border'] },
  { name: 'Explanatory text', group: 'Data and model', style: { italic: true, color: '#7F7F7F' }, cats: ['font'] },
  { name: 'Input', group: 'Data and model', style: { fill: '#FFCC99', color: '#3F3F76', ...box(thin('#7F7F7F')) }, cats: ['font', 'fill', 'border'] },
  { name: 'Linked cell', group: 'Data and model', style: { color: '#FA7D00', borderBottom: { style: 'double', color: '#FF8001' } }, cats: ['font', 'border'] },
  { name: 'Note', group: 'Data and model', style: { fill: '#FFFFCC', ...box(thin('#B2B2B2')) }, cats: ['fill', 'border'] },
  { name: 'Output', group: 'Data and model', style: { fill: '#F2F2F2', color: '#3F3F3F', bold: true, ...box(thin('#3F3F3F')) }, cats: ['font', 'fill', 'border'] },
  { name: 'Warning text', group: 'Data and model', style: { color: '#FF0000' }, cats: ['font'] },
  { name: 'Title', group: 'Titles and headings', style: { size: 18, color: '#44546A', font: 'Calibri Light' }, cats: ['font', 'border'] },
  { name: 'Heading 1', group: 'Titles and headings', style: { size: 15, bold: true, color: '#44546A', borderBottom: { style: 'thick', color: '#4472C4' } }, cats: ['font', 'border'] },
  { name: 'Heading 2', group: 'Titles and headings', style: { size: 13, bold: true, color: '#44546A', borderBottom: { style: 'thick', color: '#A2B8E1' } }, cats: ['font', 'border'] },
  { name: 'Heading 3', group: 'Titles and headings', style: { bold: true, color: '#44546A', borderBottom: { style: 'medium', color: '#8EA9DB' } }, cats: ['font', 'border'] },
  { name: 'Heading 4', group: 'Titles and headings', style: { bold: true, color: '#44546A' }, cats: ['font'] },
  { name: 'Total', group: 'Titles and headings', style: { bold: true, borderTop: thin('#4472C4'), borderBottom: { style: 'double', color: '#4472C4' } }, cats: ['font', 'border'] },
  ...[0.8, 0.6, 0.4].flatMap((t) => ACCENTS.map((c, i) => ({ name: `${Math.round((1 - t) * 100)}% - Accent${i + 1}`, group: 'Themed cell styles', style: { fill: tint(c, t), color: t === 0.4 ? '#FFFFFF' : undefined }, cats: ['font', 'fill'] as Cat[] }))),
  ...ACCENTS.map((c, i) => ({ name: `Accent${i + 1}`, group: 'Themed cell styles', style: { fill: c, color: '#FFFFFF' }, cats: ['font', 'fill'] as Cat[] })),
  { name: 'Comma', group: 'Number format', style: { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, cats: ['number'] },
  { name: 'Comma [0]', group: 'Number format', style: { numFmt: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)' }, cats: ['number'] },
  { name: 'Currency', group: 'Number format', style: { numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' }, cats: ['number'] },
  { name: 'Currency [0]', group: 'Number format', style: { numFmt: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)' }, cats: ['number'] },
  { name: 'Percent', group: 'Number format', style: { numFmt: '0%' }, cats: ['number'] }
]

/** The formatting patch for a style: what it sets, and the rest of its categories cleared. */
export function stylePatch(s: CellStyle): Partial<Style> {
  const patch: Record<string, unknown> = {}
  for (const cat of s.cats) for (const k of KEYS[cat]) patch[k] = (s.style as Record<string, unknown>)[k]
  return patch as Partial<Style>
}

export function CellStylesButton({ expanded }: { expanded: boolean }) {
  const [open, setOpen] = useState(false)
  const groups = [...new Set(CELL_STYLES.map((s) => s.group))]
  return (
    <Pop open={open} onOpenChange={setOpen}>
      <PopTrigger asChild>
        <button
          type="button"
          aria-label="Cell styles"
          onMouseDown={(e) => e.preventDefault()}
          className={cx('flex shrink-0 items-center gap-1 rounded-md px-1.5 text-text-2 hover:bg-hover hover:text-text data-[state=open]:bg-hover', expanded ? 'h-[58px] min-w-[52px] flex-col justify-center text-[11.5px] [&>svg:first-child]:size-5' : 'h-7 text-[12.5px] [&>svg:first-child]:size-4')}
        >
          <Palette />
          {expanded && (
            <span className="flex items-center gap-0.5 leading-tight">
              Styles
              <ChevronDown className="size-3 opacity-60" />
            </span>
          )}
        </button>
      </PopTrigger>
      <PopContent className="max-h-[460px] w-[520px] overflow-y-auto p-2">
        {groups.map((g) => (
          <div key={g} className="mb-2">
            <div className="mb-1 px-1 text-[11px] font-semibold text-text-3">{g}</div>
            <div className="grid grid-cols-6 gap-1">
              {CELL_STYLES.filter((s) => s.group === g).map((s) => {
                const st = s.style
                const b = st.borderBottom
                return (
                  <button
                    key={s.name}
                    type="button"
                    onClick={() => {
                      setOpen(false)
                      const c = A.ctx()
                      if (c) void A.formatSel(c, stylePatch(s))
                    }}
                    className="h-8 truncate rounded-sm border border-border px-1 text-[11px] hover:ring-2 hover:ring-accent"
                    style={{ background: st.fill ?? '#fff', color: st.color ?? '#000', fontWeight: st.bold ? 700 : 400, fontStyle: st.italic ? 'italic' : undefined, borderBottom: b ? `${b.style === 'thick' ? 3 : b.style === 'double' ? 3 : 1}px ${b.style === 'double' ? 'double' : 'solid'} ${b.color}` : undefined }}
                  >
                    {s.name}
                  </button>
                )
              })}
            </div>
          </div>
        ))}
      </PopContent>
    </Pop>
  )
}
