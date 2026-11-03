// Right-click menu for cells, rows and columns. It adapts to what is selected.
import type { ReactNode } from 'react'
import { ClipboardPaste, Copy, Eraser, EyeOff, Eye, Scissors, Rows3, Columns3, Trash2, ArrowDownAZ, ArrowUpAZ, Tag, Paintbrush, Merge, StickyNote, Link2 } from 'lucide-react'
import { MAX_COLS, MAX_ROWS } from '@/core/engine/values'
import { CtxContent, CtxItem, CtxMenu, CtxSep, CtxSub, CtxTrigger } from '@/renderer/ui'
import { runCommand, shortcutFor } from '@/renderer/app/commands'
import { useViews } from '@/renderer/state/view'
import { set } from '@/renderer/state/store'

export function GridMenu({ tab, children }: { tab: string; children: ReactNode }) {
  const sel = useViews((s) => s.views[tab]?.sel)
  const last = sel?.ranges[sel.ranges.length - 1]
  const rows = !!last && last.c0 === 0 && last.c1 === MAX_COLS - 1
  const cols = !!last && last.r0 === 0 && last.r1 === MAX_ROWS - 1
  const item = (id: string, label: string, icon?: ReactNode) => (
    <CtxItem icon={icon} shortcut={shortcutFor(id)} onSelect={() => void runCommand(id)}>
      {label}
    </CtxItem>
  )
  return (
    <CtxMenu>
      <CtxTrigger asChild>{children}</CtxTrigger>
      <CtxContent>
        {item('edit.cut', 'Cut', <Scissors />)}
        {item('edit.copy', 'Copy', <Copy />)}
        {item('edit.paste', 'Paste', <ClipboardPaste />)}
        <CtxSub label="Paste special" icon={<ClipboardPaste />}>
          {item('edit.pasteValues', 'Values')}
          {item('edit.pasteFormulas', 'Formulas')}
          {item('edit.pasteFormats', 'Formatting')}
          {item('edit.pasteSpecial', 'Paste special…')}
        </CtxSub>
        <CtxSep />
        {rows ? (
          <>
            {item('edit.insertRows', 'Insert rows', <Rows3 />)}
            {item('edit.deleteRows', 'Delete rows', <Trash2 />)}
          </>
        ) : cols ? (
          <>
            {item('edit.insertCols', 'Insert columns', <Columns3 />)}
            {item('edit.deleteCols', 'Delete columns', <Trash2 />)}
          </>
        ) : (
          <>
            <CtxSub label="Insert" icon={<Rows3 />}>
              {item('edit.insertCells', 'Cells…')}
              {item('edit.insertRows', 'Entire row')}
              {item('edit.insertCols', 'Entire column')}
            </CtxSub>
            <CtxSub label="Delete" icon={<Trash2 />}>
              {item('edit.deleteCells', 'Cells…')}
              {item('edit.deleteRows', 'Entire row')}
              {item('edit.deleteCols', 'Entire column')}
            </CtxSub>
          </>
        )}
        {item('edit.clear', 'Clear contents', <Eraser />)}
        {item('edit.clearFormats', 'Clear formats', <Paintbrush />)}
        <CtxSep />
        {!rows && !cols && (
          <>
            {item('data.sortAsc', 'Sort A to Z', <ArrowDownAZ />)}
            {item('data.sortDesc', 'Sort Z to A', <ArrowUpAZ />)}
            <CtxSep />
            {item('fmt.mergeCenter', 'Merge & center', <Merge />)}
          </>
        )}
        {item('fmt.cells', 'Format cells…')}
        {rows && (
          <>
            {item('fmt.rowHeight', 'Row height…')}
            {item('fmt.hideRows', 'Hide', <EyeOff />)}
            {item('fmt.unhideRows', 'Unhide', <Eye />)}
          </>
        )}
        {cols && (
          <>
            {item('fmt.colWidth', 'Column width…')}
            {item('fmt.hideCols', 'Hide', <EyeOff />)}
            {item('fmt.unhideCols', 'Unhide', <Eye />)}
          </>
        )}
        <CtxSep />
        {!rows && !cols && (
          <>
            {item('insert.note', 'New or edit note', <StickyNote />)}
            {item('note.delete', 'Delete note')}
            {item('insert.link', 'Link…', <Link2 />)}
            {item('link.remove', 'Remove link')}
            <CtxSep />
          </>
        )}
        <CtxItem icon={<Tag />} onSelect={() => set({ dialog: { kind: 'names' } })}>
          Define name…
        </CtxItem>
      </CtxContent>
    </CtxMenu>
  )
}
