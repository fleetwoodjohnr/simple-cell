// The ribbon: Excel's tabs and groups. "Expanded" shows group labels like Excel's ribbon; "compact" is one toolbar row.
import { useEffect, useState, type ReactNode } from 'react'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  ArrowUpAZ,
  ArrowDownWideNarrow,
  Bold,
  ChevronDown,
  ChevronUp,
  ClipboardPaste,
  Copy,
  Eraser,
  Grid3x3,
  Italic,
  Keyboard,
  CalendarDays,
  Clock,
  Merge,
  PaintBucket,
  Percent,
  Scissors,
  Search,
  Sigma,
  Snowflake,
  Strikethrough,
  Table2,
  Underline,
  WrapText,
  ZoomIn,
  ZoomOut,
  DollarSign,
  ArrowDownToLine,
  ArrowUpToLine,
  FoldVertical,
  IndentIncrease,
  IndentDecrease,
  Rows3,
  Columns3,
  FunctionSquare,
  Calculator,
  Tag,
  Eye,
  Info,
  RefreshCw,
  PanelTop,
  Sheet,
  Baseline,
  BetweenHorizontalStart,
  Trash2,
  Square,
  Highlighter,
  ListChecks,
  Filter,
  FilterX,
  PaintbrushVertical,
  Wand2,
  ListTree,
  ListX,
  Plus,
  Minus,
  Printer,
  FileDown,
  RectangleVertical,
  File,
  Scan,
  SeparatorHorizontal,
  TrendingUp,
  ChartNoAxesColumn,
  ArrowUpDown,
  Grid2x2,
  CopyMinus,
  Target,
  SlidersHorizontal,
  FileCode2,
  Code2,
  DecimalsArrowLeft,
  DecimalsArrowRight
} from 'lucide-react'
import { useShallow } from 'zustand/react/shallow'
import { useApp, set, updateSettings, books, type RibbonTab } from '../state/store'
import { setView, useViews } from '../state/view'
import { runCommand, shortcutFor } from './commands'
import * as A from './actions'
import { ColorGrid, IconBtn, Menu, MenuContent, MenuItem, MenuSep, MenuLabel, Pop, PopContent, PopTrigger, SubMenu, Tip, cx, MenuTrigger } from '../ui'
import { useBookVersion } from './hooks'
import { CondFormatMenuItems } from './CondFormat'
import { CellStylesButton } from './CellStyles'
import { PAPER } from '@/core/print'
import { CHART_TYPES, ChartTypeIcon, insertChart } from './Charts'
import { StyleGallery, StyleSwatch, openCreateTable, useActiveTable } from './Tables'
import type { Style, BorderStyle } from '@/core/model/styles'

const TABS: { id: RibbonTab; label: string }[] = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'page', label: 'Page Layout' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'data', label: 'Data' },
  { id: 'view', label: 'View' },
  { id: 'automate', label: 'Automate' },
  { id: 'help', label: 'Help' }
]

let expandedLayout = true

function Group({ children, label }: { children: ReactNode; label: string }) {
  return (
    <div role="group" aria-label={label} className={cx('flex shrink-0 flex-col items-stretch', expandedLayout ? 'border-r border-border px-2 last:border-r-0' : '')}>
      <div className={cx('flex items-center gap-0.5', expandedLayout && 'min-h-[60px] flex-wrap content-center gap-y-1')}>{children}</div>
      {expandedLayout && <div className="pb-0.5 text-center text-[10.5px] font-medium text-text-3">{label}</div>}
    </div>
  )
}

const Sep = () => (expandedLayout ? null : <div className="mx-1.5 h-6 w-px shrink-0 bg-border" />)

/** Icon button wired to a registry command. */
function Cmd({ id, icon, active, label, text }: { id: string; icon: ReactNode; active?: boolean; label: string; text?: string }) {
  if (text && expandedLayout)
    return (
      <Tip label={label} shortcut={shortcutFor(id)}>
        <button
          type="button"
          aria-label={label}
          aria-pressed={active}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => runCommand(id)}
          className={cx('flex h-[58px] min-w-[52px] shrink-0 flex-col items-center justify-center gap-1 rounded-md px-1.5 text-[11.5px] text-text-2 hover:bg-hover hover:text-text [&>svg]:size-5', active && 'bg-accent-soft text-accent-ink')}
        >
          {icon}
          <span className="leading-tight">{text}</span>
        </button>
      </Tip>
    )
  return (
    <IconBtn label={label} shortcut={shortcutFor(id)} active={active} onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand(id)}>
      {icon}
    </IconBtn>
  )
}

function DropBtn({ icon, children, label, text }: { icon?: ReactNode; children?: ReactNode; label: string; text?: string }) {
  return (
    <MenuTrigger asChild>
      <button
        type="button"
        aria-label={label}
        onMouseDown={(e) => e.preventDefault()}
        className={cx(
          'flex shrink-0 items-center gap-1 rounded-md px-1.5 text-[12.5px] text-text-2 transition-colors hover:bg-hover hover:text-text data-[state=open]:bg-hover [&>svg:first-child]:size-4',
          text && expandedLayout ? 'h-[58px] min-w-[52px] flex-col justify-center gap-1 text-[11.5px] [&>svg:first-child]:size-5' : 'h-7'
        )}
      >
        {icon}
        {text && expandedLayout ? (
          <span className="flex items-center gap-0.5 leading-tight">
            {text}
            <ChevronDown className="size-3 opacity-60" />
          </span>
        ) : (
          <>
            {children}
            <ChevronDown className="size-3 opacity-60" />
          </>
        )}
      </button>
    </MenuTrigger>
  )
}

const FONTS = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New', 'Aptos', 'Georgia', 'Verdana', 'Helvetica', 'Carlito', 'Liberation Sans', 'Liberation Serif', 'Liberation Mono']
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
const NUMBER_FORMATS: { label: string; id: string; sample: string }[] = [
  { label: 'General', id: 'fmt.general', sample: 'No specific format' },
  { label: 'Number', id: 'fmt.number', sample: '1,234.00' },
  { label: 'Currency', id: 'fmt.currency', sample: '$1,234.00' },
  { label: 'Accounting', id: 'fmt.accounting', sample: '$ 1,234.00' },
  { label: 'Short date', id: 'fmt.shortDate', sample: '1/15/2024' },
  { label: 'Long date', id: 'fmt.longDate', sample: 'Monday, January 15, 2024' },
  { label: 'Time', id: 'fmt.time', sample: '1:30 PM' },
  { label: 'Percentage', id: 'fmt.percent', sample: '12%' },
  { label: 'Scientific', id: 'fmt.scientific', sample: '1.23E+03' },
  { label: 'Text', id: 'fmt.text', sample: 'abc' }
]

function formatName(code: string | undefined): string {
  if (!code || code === 'General') return 'General'
  if (code === '@') return 'Text'
  if (/E\+/i.test(code)) return 'Scientific'
  if (/%/.test(code)) return 'Percentage'
  if (/_\(.*\*/.test(code)) return 'Accounting'
  if (/\$|€|£/.test(code)) return 'Currency'
  if (/[dmy]/i.test(code.replace(/"[^"]*"/g, '')) && !/^[#0,.]+$/.test(code)) return /h/i.test(code) && !/[dy]/i.test(code) ? 'Time' : 'Date'
  if (/[0#]/.test(code)) return 'Number'
  return 'Custom'
}

/** Style of the active cell, refreshed when the selection or workbook changes. */
function useActiveStyle(tab: string): Style {
  const book = books.get(tab)
  const version = useBookVersion(book)
  const { active, sheet } = useViews(useShallow((s) => ({ active: s.views[tab]?.sel.active, sheet: s.views[tab]?.sheet })))
  const [style, setStyle] = useState<Style>({})
  useEffect(() => {
    if (!book || !active || sheet === undefined) return
    let cancelled = false
    void book.api.styleAt(sheet, active.r, active.c).then((s) => !cancelled && setStyle(s))
    return () => {
      cancelled = true
    }
  }, [book, active?.r, active?.c, sheet, version])
  return style
}

function withCtx(fn: (c: A.Ctx) => unknown) {
  const c = A.ctx()
  if (c) void fn(c)
}

function BorderMenu() {
  const set1 = (patch: Partial<Style>) => withCtx((c) => A.formatSel(c, patch))
  const thin = (style: BorderStyle = 'thin') => ({ style })
  const sides = (o: { t?: boolean; b?: boolean; l?: boolean; r?: boolean }, style: BorderStyle = 'thin') =>
    set1({ borderTop: o.t ? thin(style) : undefined, borderBottom: o.b ? thin(style) : undefined, borderLeft: o.l ? thin(style) : undefined, borderRight: o.r ? thin(style) : undefined })
  return (
    <Menu>
      <DropBtn icon={<Grid3x3 />} label="Borders" text="Borders" />
      <MenuContent>
        <MenuItem onSelect={() => withCtx((c) => A.formatSel(c, { borderBottom: thin() }))}>Bottom border</MenuItem>
        <MenuItem onSelect={() => withCtx((c) => A.formatSel(c, { borderTop: thin() }))}>Top border</MenuItem>
        <MenuItem onSelect={() => withCtx((c) => A.formatSel(c, { borderLeft: thin() }))}>Left border</MenuItem>
        <MenuItem onSelect={() => withCtx((c) => A.formatSel(c, { borderRight: thin() }))}>Right border</MenuItem>
        <MenuSep />
        <MenuItem shortcut={shortcutFor('fmt.noBorder')} onSelect={() => runCommand('fmt.noBorder')}>
          No border
        </MenuItem>
        <MenuItem onSelect={() => sides({ t: true, b: true, l: true, r: true })}>All borders</MenuItem>
        <MenuItem shortcut={shortcutFor('fmt.outlineBorder')} onSelect={() => runCommand('fmt.outlineBorder')}>
          Outside borders
        </MenuItem>
        <MenuItem onSelect={() => withCtx((c) => A.formatSel(c, { borderBottom: thin('thick') }))}>Thick bottom border</MenuItem>
        <MenuItem onSelect={() => withCtx((c) => A.formatSel(c, { borderBottom: thin('double') }))}>Bottom double border</MenuItem>
        <MenuItem onSelect={() => withCtx((c) => A.formatSel(c, { borderTop: thin(), borderBottom: thin('double') }))}>Top and double bottom border</MenuItem>
      </MenuContent>
    </Menu>
  )
}

function ColorBtn({ label, icon, current, onPick, onClear, clearLabel }: { label: string; icon: ReactNode; current?: string; onPick: (c: string) => void; onClear: () => void; clearLabel: string }) {
  return (
    <Pop>
      <Tip label={label}>
        <PopTrigger asChild>
          <button type="button" aria-label={label} onMouseDown={(e) => e.preventDefault()} className="flex h-7 shrink-0 flex-col items-center justify-center rounded-md px-1.5 text-text-2 hover:bg-hover hover:text-text">
            <span className="[&>svg]:size-4">{icon}</span>
            <span className="mt-[1px] h-[3px] w-4 rounded-full" style={{ background: current ?? 'transparent', boxShadow: current ? undefined : 'inset 0 0 0 1px var(--border-strong)' }} />
          </button>
        </PopTrigger>
      </Tip>
      <PopContent>
        <ColorGrid value={current} onPick={onPick} onClear={onClear} clearLabel={clearLabel} />
      </PopContent>
    </Pop>
  )
}

/** Format Painter: click once to paint the next selection; double-click to keep painting until Esc. */
function PainterBtn({ tab }: { tab: string }) {
  const on = useViews((s) => !!s.views[tab]?.painter)
  return (
    <IconBtn
      label="Format painter (double-click to keep it on)"
      active={on}
      onMouseDown={(e) => e.preventDefault()}
      onClick={(e) => {
        if (e.detail > 1) return
        withCtx((c) => (on ? setView(tab, { painter: undefined }) : A.startPainter(c, false)))
      }}
      onDoubleClick={() => withCtx((c) => A.startPainter(c, true))}
      data-testid="format-painter"
    >
      <PaintbrushVertical className="size-4" />
    </IconBtn>
  )
}

/** Home ▸ Format as Table: pick a style, then confirm the range. */
function FormatAsTable() {
  const [open, setOpen] = useState(false)
  return (
    <Pop open={open} onOpenChange={setOpen}>
      <PopTrigger asChild>
        <button
          type="button"
          aria-label="Format as table"
          onMouseDown={(e) => e.preventDefault()}
          className={cx('flex shrink-0 items-center gap-1 rounded-md px-1.5 text-text-2 hover:bg-hover hover:text-text data-[state=open]:bg-hover', expandedLayout ? 'h-[58px] min-w-[52px] flex-col justify-center text-[11.5px] [&>svg:first-child]:size-5' : 'h-7 text-[12.5px] [&>svg:first-child]:size-4')}
        >
          <Table2 />
          {expandedLayout && (
            <span className="flex items-center gap-0.5 leading-tight">
              Table
              <ChevronDown className="size-3 opacity-60" />
            </span>
          )}
        </button>
      </PopTrigger>
      <PopContent className="p-0">
        <StyleGallery
          onPick={(style) => {
            setOpen(false)
            void openCreateTable(style)
          }}
        />
      </PopContent>
    </Pop>
  )
}

/** Table Design: shown while the active cell is in a table. */
function TableTab({ tab }: { tab: string }) {
  const book = books.get(tab)
  const t = useActiveTable(tab, book)
  const [name, setName] = useState(t?.name ?? '')
  useEffect(() => setName(t?.name ?? ''), [t?.name])
  if (!t || !book) return <span className="px-2 py-4 text-[12px] text-text-3">Select a cell in a table.</span>
  const opt = (patch: Parameters<typeof book.api.setTableOptions>[1]) => withCtx((c) => A.edit(c, (b) => b.api.setTableOptions(t.name, patch)))
  const check = (key: 'totalsRow' | 'showRowStripes' | 'showColumnStripes' | 'showFirstColumn' | 'showLastColumn' | 'autoFilter', label: string) => (
    <label className="flex items-center gap-1.5 whitespace-nowrap px-1 text-[12px] text-text-2">
      <input type="checkbox" checked={!!t[key]} onChange={(e) => opt({ [key]: e.target.checked })} /> {label}
    </label>
  )
  return (
    <>
      <Group label="Properties">
        <div className="flex flex-col gap-1 px-1">
          <span className="text-[11px] text-text-3">Table name</span>
          <input
            aria-label="Table name"
            className="h-7 w-32 rounded-md border border-border bg-panel px-2 text-[12.5px] outline-none focus:border-accent"
            value={name}
            onChange={(e) => setName(e.target.value)}
            onBlur={() => name !== t.name && opt({ name })}
            onKeyDown={(e) => e.key === 'Enter' && (e.currentTarget as HTMLInputElement).blur()}
          />
        </div>
      </Group>
      <Sep />
      <Group label="Tools">
        <Cmd id="table.convert" label="Convert to range" text="To range" icon={<Grid3x3 />} />
      </Group>
      <Sep />
      <Group label="Style options">
        <div className="grid grid-cols-2 gap-x-2 gap-y-1">
          {check('totalsRow', 'Total row')}
          {check('showFirstColumn', 'First column')}
          {check('showRowStripes', 'Banded rows')}
          {check('showLastColumn', 'Last column')}
          {check('showColumnStripes', 'Banded columns')}
          {check('autoFilter', 'Filter button')}
        </div>
      </Group>
      <Sep />
      <Group label="Table styles">
        <Pop>
          <PopTrigger asChild>
            <button type="button" aria-label="Table style" className="flex items-center gap-1 rounded-md px-1.5 py-1 hover:bg-hover">
              <StyleSwatch name={t.style || 'TableStyleMedium2'} />
              <ChevronDown className="size-3 opacity-60" />
            </button>
          </PopTrigger>
          <PopContent className="p-0">
            <StyleGallery current={t.style} onPick={(style) => opt({ style })} />
          </PopContent>
        </Pop>
      </Group>
    </>
  )
}

function HomeTab({ tab }: { tab: string }) {
  const s = useActiveStyle(tab)
  const defaultFont = useApp((x) => x.settings.defaultFont)
  const defaultSize = useApp((x) => x.settings.defaultFontSize)
  const font = s.font ?? defaultFont
  const size = s.size ?? defaultSize
  const fmt = (patch: Partial<Style>) => withCtx((c) => A.formatSel(c, patch))
  return (
    <>
      <Group label="Clipboard">
        <Cmd id="edit.paste" label="Paste" text="Paste" icon={<ClipboardPaste />} />
        <div className={cx('flex gap-0.5', expandedLayout && 'flex-col')}>
          <Cmd id="edit.cut" label="Cut" icon={<Scissors className="size-4" />} />
          <Cmd id="edit.copy" label="Copy" icon={<Copy className="size-4" />} />
          <PainterBtn tab={tab} />
        </div>
      </Group>
      <Sep />
      <Group label="Font">
        <div className={cx('flex flex-col gap-1', !expandedLayout && 'flex-row')}>
          <div className="flex gap-1">
            <Menu>
              <DropBtn label={`Font: ${font}`}>
                <span className="w-[104px] truncate text-left" style={{ fontFamily: font }}>
                  {font}
                </span>
              </DropBtn>
              <MenuContent className="max-h-[360px] overflow-y-auto">
                {FONTS.map((f) => (
                  <MenuItem key={f} onSelect={() => fmt({ font: f })}>
                    <span style={{ fontFamily: f }}>{f}</span>
                  </MenuItem>
                ))}
              </MenuContent>
            </Menu>
            <Menu>
              <DropBtn label={`Font size: ${size}`}>
                <span className="w-6 text-left tabular-nums">{size}</span>
              </DropBtn>
              <MenuContent className="max-h-[360px] min-w-[90px] overflow-y-auto">
                {SIZES.map((z) => (
                  <MenuItem key={z} onSelect={() => fmt({ size: z })}>
                    {z}
                  </MenuItem>
                ))}
              </MenuContent>
            </Menu>
          </div>
          <div className="flex items-center gap-0.5">
            <Cmd id="fmt.bold" label="Bold" active={!!s.bold} icon={<Bold className="size-4" />} />
            <Cmd id="fmt.italic" label="Italic" active={!!s.italic} icon={<Italic className="size-4" />} />
            <Cmd id="fmt.underline" label="Underline" active={!!s.underline} icon={<Underline className="size-4" />} />
            <Cmd id="fmt.strike" label="Strikethrough" active={!!s.strike} icon={<Strikethrough className="size-4" />} />
            <BorderMenu />
            <ColorBtn label="Fill color" icon={<PaintBucket />} current={s.fill} onPick={(c) => fmt({ fill: c })} onClear={() => fmt({ fill: undefined })} clearLabel="No fill" />
            <ColorBtn label="Font color" icon={<Baseline />} current={s.color} onPick={(c) => fmt({ color: c })} onClear={() => fmt({ color: undefined })} clearLabel="Automatic" />
          </div>
        </div>
      </Group>
      <Sep />
      <Group label="Alignment">
        <div className={cx('flex gap-0.5', expandedLayout && 'flex-col gap-1')}>
          <div className="flex gap-0.5">
            <Cmd id="fmt.alignTop" label="Top align" active={s.vAlign === 'top'} icon={<ArrowUpToLine className="size-4" />} />
            <Cmd id="fmt.alignMiddle" label="Middle align" active={s.vAlign === 'center'} icon={<FoldVertical className="size-4" />} />
            <Cmd id="fmt.alignBottom" label="Bottom align" active={!s.vAlign || s.vAlign === 'bottom'} icon={<ArrowDownToLine className="size-4" />} />
            <Cmd id="fmt.wrap" label="Wrap text" active={!!s.wrap} icon={<WrapText className="size-4" />} />
            <Cmd id="fmt.outdent" label="Decrease indent" icon={<IndentDecrease className="size-4" />} />
          </div>
          <div className="flex gap-0.5">
            <Cmd id="fmt.alignLeft" label="Align left" active={s.hAlign === 'left'} icon={<AlignLeft className="size-4" />} />
            <Cmd id="fmt.alignCenter" label="Center" active={s.hAlign === 'center'} icon={<AlignCenter className="size-4" />} />
            <Cmd id="fmt.alignRight" label="Align right" active={s.hAlign === 'right'} icon={<AlignRight className="size-4" />} />
            <Menu>
              <DropBtn icon={<Merge />} label="Merge" />
              <MenuContent>
                <MenuItem onSelect={() => runCommand('fmt.mergeCenter')}>Merge & center</MenuItem>
                <MenuItem onSelect={() => runCommand('fmt.mergeAcross')}>Merge across</MenuItem>
                <MenuItem onSelect={() => runCommand('fmt.merge')}>Merge cells</MenuItem>
                <MenuItem onSelect={() => runCommand('fmt.unmerge')}>Unmerge cells</MenuItem>
              </MenuContent>
            </Menu>
            <Cmd id="fmt.indent" label="Increase indent" icon={<IndentIncrease className="size-4" />} />
          </div>
        </div>
      </Group>
      <Sep />
      <Group label="Number">
        <div className={cx('flex gap-1', expandedLayout && 'flex-col')}>
          <Menu>
            <DropBtn label={`Number format: ${formatName(s.numFmt)}`}>
              <span className="w-[92px] truncate text-left">{formatName(s.numFmt)}</span>
            </DropBtn>
            <MenuContent>
              {NUMBER_FORMATS.map((f) => (
                <MenuItem key={f.id} onSelect={() => runCommand(f.id)}>
                  <span className="flex w-full items-center justify-between gap-6">
                    <span>{f.label}</span>
                    <span className="text-[11.5px] text-text-3">{f.sample}</span>
                  </span>
                </MenuItem>
              ))}
              <MenuSep />
              <MenuItem shortcut={shortcutFor('fmt.cells')} onSelect={() => runCommand('fmt.cells')}>
                More number formats…
              </MenuItem>
            </MenuContent>
          </Menu>
          <div className="flex gap-0.5">
            <Cmd id="fmt.accounting" label="Accounting number format" icon={<DollarSign className="size-4" />} />
            <Cmd id="fmt.percent" label="Percent style" icon={<Percent className="size-4" />} />
            <Cmd id="fmt.thousands" label="Comma style" icon={<span aria-hidden className="text-[15px] font-bold leading-none">,</span>} />
            <Cmd id="fmt.moreDecimals" label="Increase decimal" icon={<DecimalsArrowRight className="size-4" />} />
            <Cmd id="fmt.lessDecimals" label="Decrease decimal" icon={<DecimalsArrowLeft className="size-4" />} />
          </div>
        </div>
      </Group>
      <Sep />
      <Group label="Styles">
        <Menu>
          <DropBtn icon={<Highlighter />} label="Conditional formatting" text="Conditional" />
          <MenuContent>
            <CondFormatMenuItems />
          </MenuContent>
        </Menu>
        <FormatAsTable />
        <CellStylesButton expanded={expandedLayout} />
      </Group>
      <Sep />
      <Group label="Cells">
        <Menu>
          <DropBtn icon={<BetweenHorizontalStart />} label="Insert" text="Insert" />
          <MenuContent>
            <MenuItem shortcut={shortcutFor('edit.insertCells')} onSelect={() => runCommand('edit.insertCells')}>Insert cells…</MenuItem>
            <MenuItem onSelect={() => runCommand('edit.insertRows')}>Insert sheet rows</MenuItem>
            <MenuItem onSelect={() => runCommand('edit.insertCols')}>Insert sheet columns</MenuItem>
            <MenuItem shortcut={shortcutFor('sheet.insert')} onSelect={() => runCommand('sheet.insert')}>
              Insert sheet
            </MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<Trash2 />} label="Delete" text="Delete" />
          <MenuContent>
            <MenuItem shortcut={shortcutFor('edit.deleteCells')} onSelect={() => runCommand('edit.deleteCells')}>Delete cells…</MenuItem>
            <MenuItem onSelect={() => runCommand('edit.deleteRows')}>Delete sheet rows</MenuItem>
            <MenuItem onSelect={() => runCommand('edit.deleteCols')}>Delete sheet columns</MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<Table2 />} label="Format" text="Format" />
          <MenuContent>
            <MenuLabel>Cell size</MenuLabel>
            <MenuItem onSelect={() => runCommand('fmt.rowHeight')}>Row height…</MenuItem>
            <MenuItem onSelect={() => runCommand('fmt.colWidth')}>Column width…</MenuItem>
            <MenuLabel>Visibility</MenuLabel>
            <SubMenu label="Hide & unhide">
              <MenuItem onSelect={() => runCommand('fmt.hideRows')}>Hide rows</MenuItem>
              <MenuItem onSelect={() => runCommand('fmt.hideCols')}>Hide columns</MenuItem>
              <MenuItem onSelect={() => runCommand('fmt.unhideRows')}>Unhide rows</MenuItem>
              <MenuItem onSelect={() => runCommand('fmt.unhideCols')}>Unhide columns</MenuItem>
            </SubMenu>
            <MenuSep />
            <MenuItem shortcut={shortcutFor('fmt.cells')} onSelect={() => runCommand('fmt.cells')}>
              Format cells…
            </MenuItem>
          </MenuContent>
        </Menu>
      </Group>
      <Sep />
      <Group label="Editing">
        <Menu>
          <DropBtn icon={<Sigma />} label="AutoSum" text="AutoSum" />
          <MenuContent>
            <MenuItem shortcut={shortcutFor('formulas.autosum')} onSelect={() => runCommand('formulas.autosum')}>
              Sum
            </MenuItem>
            <MenuItem onSelect={() => runCommand('formulas.average')}>Average</MenuItem>
            <MenuItem onSelect={() => runCommand('formulas.count')}>Count numbers</MenuItem>
            <MenuItem onSelect={() => runCommand('formulas.max')}>Max</MenuItem>
            <MenuItem onSelect={() => runCommand('formulas.min')}>Min</MenuItem>
            <MenuSep />
            <MenuItem onSelect={() => runCommand('formulas.insertFunction')}>More functions…</MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<ArrowDownWideNarrow />} label="Fill" />
          <MenuContent>
            <MenuItem shortcut={shortcutFor('edit.fillDown')} onSelect={() => runCommand('edit.fillDown')}>
              Down
            </MenuItem>
            <MenuItem shortcut={shortcutFor('edit.fillRight')} onSelect={() => runCommand('edit.fillRight')}>
              Right
            </MenuItem>
            <MenuItem onSelect={() => runCommand('edit.fillUp')}>Up</MenuItem>
            <MenuItem onSelect={() => runCommand('edit.fillLeft')}>Left</MenuItem>
            <MenuSep />
            <MenuItem shortcut={shortcutFor('data.flashFill')} onSelect={() => runCommand('data.flashFill')}>
              Flash Fill
            </MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<Eraser />} label="Clear" />
          <MenuContent>
            <MenuItem onSelect={() => runCommand('edit.clearAll')}>Clear all</MenuItem>
            <MenuItem onSelect={() => runCommand('edit.clearFormats')}>Clear formats</MenuItem>
            <MenuItem onSelect={() => runCommand('edit.clear')}>Clear contents</MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<ArrowDownAZ />} label="Sort & filter" text="Sort" />
          <MenuContent>
            <MenuItem onSelect={() => runCommand('data.sortAsc')}>Sort A to Z</MenuItem>
            <MenuItem onSelect={() => runCommand('data.sortDesc')}>Sort Z to A</MenuItem>
            <MenuItem onSelect={() => runCommand('data.sort')}>Custom sort…</MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<Search />} label="Find & select" text="Find" />
          <MenuContent>
            <MenuItem shortcut={shortcutFor('edit.find')} onSelect={() => runCommand('edit.find')}>
              Find…
            </MenuItem>
            <MenuItem shortcut={shortcutFor('edit.replace')} onSelect={() => runCommand('edit.replace')}>
              Replace…
            </MenuItem>
            <MenuItem shortcut={shortcutFor('edit.goto')} onSelect={() => runCommand('edit.goto')}>
              Go to…
            </MenuItem>
            <MenuItem onSelect={() => runCommand('edit.gotoSpecial')}>Go to special…</MenuItem>
            <MenuSep />
            <MenuItem onSelect={() => runCommand('select.formulas')}>Formulas</MenuItem>
            <MenuItem onSelect={() => runCommand('select.notes')}>Notes</MenuItem>
            <MenuItem onSelect={() => runCommand('select.condFormats')}>Conditional formatting</MenuItem>
            <MenuItem onSelect={() => runCommand('select.constants')}>Constants</MenuItem>
            <MenuItem onSelect={() => runCommand('select.validation')}>Data validation</MenuItem>
          </MenuContent>
        </Menu>
      </Group>
    </>
  )
}

function InsertTab() {
  return (
    <>
      <Group label="Tables">
        <Cmd id="insert.pivot" label="PivotTable" text="PivotTable" icon={<Grid2x2 />} />
        <Cmd id="data.table" label="Table" text="Table" icon={<Table2 />} />
      </Group>
      <Sep />
      <Group label="Sparklines">
        <Cmd id="insert.sparkLine" label="Line sparklines" text="Line" icon={<TrendingUp />} />
        <Cmd id="insert.sparkColumn" label="Column sparklines" text="Column" icon={<ChartNoAxesColumn />} />
        <Cmd id="insert.sparkWinLoss" label="Win/loss sparklines" text="Win/Loss" icon={<ArrowUpDown />} />
      </Group>
      <Sep />
      <Group label="Charts">
        {CHART_TYPES.map(([t, label]) => (
          <button
            key={t}
            type="button"
            aria-label={`Insert ${label.toLowerCase()} chart`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => void insertChart(t)}
            className={cx('flex shrink-0 flex-col items-center justify-center gap-1 rounded-md px-1.5 text-[11.5px] text-text-2 hover:bg-hover hover:text-text', expandedLayout ? 'h-[58px] min-w-[46px] [&>svg]:size-5' : 'h-7 [&>svg]:size-4')}
          >
            <ChartTypeIcon type={t} />
            {expandedLayout && <span className="leading-tight">{label}</span>}
          </button>
        ))}
      </Group>
      <Sep />
      <Group label="Cells">
        <Cmd id="edit.insertRows" label="Insert sheet rows" text="Rows" icon={<Rows3 />} />
        <Cmd id="edit.insertCols" label="Insert sheet columns" text="Columns" icon={<Columns3 />} />
        <Cmd id="sheet.insert" label="Insert sheet" text="Sheet" icon={<Sheet />} />
      </Group>
      <Sep />
      <Group label="Text">
        <Cmd id="edit.insertDate" label="Insert today’s date" text="Date" icon={<CalendarDays />} />
        <Cmd id="edit.insertTime" label="Insert the current time" text="Time" icon={<Clock />} />
      </Group>
    </>
  )
}

function FormulasTab({ tab }: { tab: string }) {
  const book = books.get(tab)
  useBookVersion(book)
  const showFormulas = useApp((s) => s.showFormulas)
  const manual = book?.meta.calcMode === 'manual'
  return (
    <>
      <Group label="Function library">
        <Cmd id="formulas.insertFunction" label="Insert function" text="Insert function" icon={<FunctionSquare />} />
        <Cmd id="formulas.autosum" label="AutoSum" text="AutoSum" icon={<Sigma />} />
      </Group>
      <Sep />
      <Group label="Defined names">
        <Cmd id="formulas.names" label="Name manager" text="Name manager" icon={<Tag />} />
      </Group>
      <Sep />
      <Group label="Formula auditing">
        <Cmd id="formulas.showFormulas" label="Show formulas" text="Show formulas" active={showFormulas} icon={<Eye />} />
      </Group>
      <Sep />
      <Group label="Calculation">
        <Menu>
          <DropBtn icon={<Calculator />} label={`Calculation options: ${manual ? 'Manual' : 'Automatic'}`} text={manual ? 'Manual' : 'Automatic'} />
          <MenuContent>
            <MenuItem onSelect={() => runCommand('formulas.calcAuto')}>{!manual ? '✓ ' : ''}Automatic</MenuItem>
            <MenuItem onSelect={() => runCommand('formulas.calcManual')}>{manual ? '✓ ' : ''}Manual</MenuItem>
          </MenuContent>
        </Menu>
        <Cmd id="formulas.calculate" label="Calculate now" text="Calculate now" icon={<RefreshCw />} />
      </Group>
    </>
  )
}

function PageTab({ tab }: { tab: string }) {
  const book = books.get(tab)
  const sheet = useViews((s) => s.views[tab]?.sheet)
  const [ps, setPs] = useState<{ orientation: string; paper: number; gridlines: boolean; headings: boolean } | null>(null)
  const v = book?.version
  useEffect(() => {
    if (book && sheet !== undefined) void book.api.pageSetup(sheet).then(setPs)
  }, [book, sheet, v])
  const change = (patch: Record<string, unknown>) => withCtx((c) => A.edit(c, (b) => b.api.setPageSetup(c.view.sheet, patch)))
  return (
    <>
      <Group label="Print">
        <Cmd id="file.print" label="Print" text="Print" icon={<Printer />} />
        <Cmd id="file.exportPdf" label="Save as PDF" text="PDF" icon={<FileDown />} />
      </Group>
      <Sep />
      <Group label="Page setup">
        <Menu>
          <DropBtn icon={<RectangleVertical />} label="Orientation" text="Orientation" />
          <MenuContent>
            <MenuItem onSelect={() => change({ orientation: 'portrait' })}>Portrait{ps?.orientation === 'portrait' ? ' ✓' : ''}</MenuItem>
            <MenuItem onSelect={() => change({ orientation: 'landscape' })}>Landscape{ps?.orientation === 'landscape' ? ' ✓' : ''}</MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<File />} label="Size" text="Size" />
          <MenuContent>
            {Object.entries(PAPER).map(([k, p]) => (
              <MenuItem key={k} onSelect={() => change({ paper: Number(k) })}>
                {p.name}
                {ps?.paper === Number(k) ? ' ✓' : ''}
              </MenuItem>
            ))}
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<Scan />} label="Print area" text="Print area" />
          <MenuContent>
            <MenuItem onSelect={() => runCommand('page.setPrintArea')}>Set print area</MenuItem>
            <MenuItem onSelect={() => runCommand('page.clearPrintArea')}>Clear print area</MenuItem>
          </MenuContent>
        </Menu>
        <Menu>
          <DropBtn icon={<SeparatorHorizontal />} label="Breaks" text="Breaks" />
          <MenuContent>
            <MenuItem onSelect={() => runCommand('page.breakRow')}>Insert page break above the active row</MenuItem>
            <MenuItem onSelect={() => runCommand('page.resetBreaks')}>Reset all page breaks</MenuItem>
          </MenuContent>
        </Menu>
        <Cmd id="page.printTitles" label="Print titles: repeat the selected rows at the top of every page" text="Titles" icon={<Rows3 />} />
      </Group>
      <Sep />
      <Group label="Sheet options">
        <div className="flex flex-col gap-2 px-1 text-[12px] text-text-2">
          <label className="flex items-center gap-1.5">
            <input type="checkbox" checked={!!ps?.gridlines} onChange={(e) => change({ gridlines: e.target.checked })} /> Print gridlines
          </label>
          <label className="flex items-center gap-1.5">
            <input type="checkbox" checked={!!ps?.headings} onChange={(e) => change({ headings: e.target.checked })} /> Print headings
          </label>
        </div>
      </Group>
    </>
  )
}

function AutomateTab() {
  return (
    <>
      <Group label="Scripts">
        <Cmd id="tools.macros" label="Macros" text="Macros" icon={<FileCode2 />} />
        <Cmd id="tools.vba" label="View VBA code" text="VBA" icon={<Code2 />} />
      </Group>
    </>
  )
}

function DataTab() {
  return (
    <>
      <Group label="Sort">
        <Cmd id="data.sortAsc" label="Sort A to Z" text="A to Z" icon={<ArrowDownAZ />} />
        <Cmd id="data.sortDesc" label="Sort Z to A" text="Z to A" icon={<ArrowUpAZ />} />
        <Cmd id="data.sort" label="Custom sort" text="Sort" icon={<ArrowDownWideNarrow />} />
      </Group>
      <Sep />
      <Group label="Queries">
        <Cmd id="data.refreshAll" label="Refresh all" text="Refresh" icon={<RefreshCw />} />
      </Group>
      <Sep />
      <Group label="Filter">
        <Cmd id="data.filter" label="Filter" text="Filter" icon={<Filter />} />
        <Cmd id="data.clearFilter" label="Clear filters" text="Clear" icon={<FilterX />} />
        <Cmd id="data.reapply" label="Reapply filters" text="Reapply" icon={<RefreshCw />} />
      </Group>
      <Sep />
      <Group label="Analysis">
        <Cmd id="data.goalSeek" label="Goal seek" text="Goal Seek" icon={<Target />} />
        <Cmd id="data.solver" label="Solver" text="Solver" icon={<SlidersHorizontal />} />
        <Cmd id="data.descriptiveStats" label="Descriptive statistics" text="Statistics" icon={<Sigma />} />
      </Group>
      <Sep />
      <Group label="Outline">
        <Cmd id="data.group" label="Group" text="Group" icon={<ListTree />} />
        <Cmd id="data.ungroup" label="Ungroup" text="Ungroup" icon={<ListX />} />
        <div className={cx('flex gap-0.5', expandedLayout && 'flex-col')}>
          <Cmd id="data.showDetail" label="Show detail" icon={<Plus className="size-4" />} />
          <Cmd id="data.hideDetail" label="Hide detail" icon={<Minus className="size-4" />} />
        </div>
      </Group>
      <Sep />
      <Group label="Data tools">
        <Cmd id="data.flashFill" label="Flash Fill" text="Flash Fill" icon={<Wand2 />} />
        <Cmd id="data.textToColumns" label="Split text to columns" text="Split" icon={<Columns3 />} />
        <Cmd id="data.removeDuplicates" label="Remove duplicates" text="Duplicates" icon={<CopyMinus />} />
        <Menu>
          <DropBtn icon={<ListChecks />} label="Data validation" text="Validation" />
          <MenuContent>
            <MenuItem onSelect={() => runCommand('data.validation')}>Data validation…</MenuItem>
            <MenuItem onSelect={() => runCommand('data.circleInvalid')}>Circle invalid data</MenuItem>
            <MenuItem onSelect={() => runCommand('data.clearCircles')}>Clear validation circles</MenuItem>
          </MenuContent>
        </Menu>
      </Group>
    </>
  )
}

function ViewTab({ tab }: { tab: string }) {
  const book = books.get(tab)
  useBookVersion(book)
  const sheet = useViews((s) => s.views[tab]?.sheet)
  const m = sheet !== undefined ? book?.sheetMeta(sheet) : undefined
  const { headings, formulaBar } = useApp(useShallow((s) => ({ headings: s.showHeadings, formulaBar: s.showFormulaBar })))
  const frozen = !!m && (m.freeze.rows > 0 || m.freeze.cols > 0)
  return (
    <>
      <Group label="Show">
        <Cmd id="view.gridlines" label="Gridlines" text="Gridlines" active={m?.showGridlines} icon={<Grid3x3 />} />
        <Cmd id="view.headings" label="Headings" text="Headings" active={headings} icon={<PanelTop />} />
        <Cmd id="view.formulaBar" label="Formula bar" text="Formula bar" active={formulaBar} icon={<FunctionSquare />} />
      </Group>
      <Sep />
      <Group label="Zoom">
        <Cmd id="view.zoomOut" label="Zoom out" text="Zoom out" icon={<ZoomOut />} />
        <Cmd id="view.zoom100" label="100%" text="100%" icon={<Square />} />
        <Cmd id="view.zoomIn" label="Zoom in" text="Zoom in" icon={<ZoomIn />} />
      </Group>
      <Sep />
      <Group label="Window">
        <Menu>
          <DropBtn icon={<Snowflake />} label="Freeze panes" text="Freeze panes" />
          <MenuContent>
            <MenuItem onSelect={() => runCommand('view.freezePanes')}>{frozen ? 'Unfreeze panes' : 'Freeze panes'}</MenuItem>
            <MenuItem onSelect={() => runCommand('view.freezeTopRow')}>Freeze top row</MenuItem>
            <MenuItem onSelect={() => runCommand('view.freezeFirstColumn')}>Freeze first column</MenuItem>
          </MenuContent>
        </Menu>
      </Group>
      <Sep />
      <Group label="Layout">
        <Cmd id="view.ribbonMode" label="Ribbon layout: compact toolbar or full ribbon" text="Ribbon layout" icon={<PanelTop />} />
      </Group>
    </>
  )
}

function HelpTab() {
  return (
    <>
      <Group label="Help">
        <Cmd id="tools.shortcuts" label="Keyboard shortcuts" text="Shortcuts" icon={<Keyboard />} />
        <Cmd id="tools.palette" label="Search commands" text="Search" icon={<Search />} />
        <Cmd id="tools.checkUpdates" label="Check for updates" text="Updates" icon={<RefreshCw />} />
        <Cmd id="tools.about" label="About Simple Cell" text="About" icon={<Info />} />
      </Group>
    </>
  )
}

export function Ribbon({ tab: docTab }: { tab: string }) {
  const { tab: chosen, collapsed, mode } = useApp(useShallow((x) => ({ tab: x.ribbonTab, collapsed: x.settings.ribbonCollapsed, mode: x.settings.ribbonMode })))
  const book = books.get(docTab)
  useBookVersion(book)
  const activeTable = useActiveTable(docTab, book)
  const tab = chosen === 'table' && !activeTable ? 'home' : chosen
  expandedLayout = mode !== 'compact'
  return (
    <div className="shrink-0 border-b border-border bg-panel" data-testid="ribbon">
      <div className="flex h-9 items-center gap-0.5 px-2">
        <button
          onClick={() => set((x) => ({ backstage: !x.backstage }))}
          data-testid="file-tab"
          className="mr-1 h-[26px] rounded-md bg-accent px-3 text-[12.5px] font-semibold text-accent-fg shadow-[inset_0_1px_0_rgb(255_255_255/0.2)] hover:brightness-110"
        >
          File
        </button>
        <div role="tablist" aria-label="Ribbon" className="flex items-center gap-0.5">
          {[...TABS, ...(activeTable ? [{ id: 'table' as RibbonTab, label: 'Table Design' }] : [])].map((t) => (
            <button
              key={t.id}
              role="tab"
              aria-selected={tab === t.id && !collapsed}
              data-testid={`ribbon-tab-${t.id}`}
              onClick={() => {
                set({ ribbonTab: t.id })
                if (collapsed) void updateSettings({ ribbonCollapsed: false })
              }}
              className={cx('relative h-[26px] rounded-md px-2.5 text-[12.5px] font-medium transition-colors', tab === t.id && !collapsed ? 'text-text' : 'text-text-2 hover:bg-hover hover:text-text')}
            >
              {t.label}
              {tab === t.id && !collapsed && <span className="absolute inset-x-2.5 -bottom-[5px] h-[2px] rounded-full bg-accent" />}
            </button>
          ))}
        </div>
        <div className="flex-1" />
        <IconBtn label={collapsed ? 'Expand ribbon' : 'Collapse ribbon'} size="sm" shortcut={shortcutFor('view.ribbon')} onClick={() => runCommand('view.ribbon')}>
          {collapsed ? <ChevronDown className="size-3.5" /> : <ChevronUp className="size-3.5" />}
        </IconBtn>
      </div>
      {!collapsed && (
        <div className={cx('flex items-stretch overflow-x-auto overflow-y-hidden px-2 [scrollbar-width:none]', expandedLayout ? 'py-1' : 'h-11 items-center gap-0.5')} role="toolbar" aria-label={`${tab} tools`}>
          {tab === 'home' && <HomeTab tab={docTab} />}
          {tab === 'insert' && <InsertTab />}
          {tab === 'page' && <PageTab tab={docTab} />}
          {tab === 'formulas' && <FormulasTab tab={docTab} />}
          {tab === 'data' && <DataTab />}
          {tab === 'view' && <ViewTab tab={docTab} />}
          {tab === 'help' && <HelpTab />}
          {tab === 'automate' && <AutomateTab />}
          {tab === 'table' && <TableTab tab={docTab} />}
        </div>
      )}
    </div>
  )
}
