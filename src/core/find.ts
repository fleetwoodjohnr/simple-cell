// Find and Replace (Ctrl+F, Ctrl+H).
import type { Doc } from './doc'
import { wildcardRegex } from './engine/functions/helpers'
import { MAX_COLS, MAX_ROWS } from './engine/values'

export interface FindOptions {
  text: string
  /** Sheet id to search, or null for the whole workbook. */
  sheet: number | null
  lookIn: 'formulas' | 'values'
  matchCase: boolean
  wholeCell: boolean
  byColumns?: boolean
  limit?: number
}

export interface Match {
  sheet: number
  r: number
  c: number
  text: string
}

function matcher(o: FindOptions): (s: string) => boolean {
  const pattern = o.wholeCell ? o.text : `*${o.text}*`
  const re = new RegExp(wildcardRegex(pattern).source, o.matchCase ? '' : 'i')
  return (s) => re.test(s)
}

export function findCells(doc: Doc, o: FindOptions): Match[] {
  if (!o.text) return []
  const test = matcher(o)
  const out: Match[] = []
  const limit = o.limit ?? 10_000
  for (const sh of doc.wb.sheets) {
    if (o.sheet !== null && sh.id !== o.sheet) continue
    const hits: Match[] = []
    sh.forEachIn({ sheet: sh.id, r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }, (r, c, cell) => {
      if (cell.spillOf !== undefined && !cell.f) return
      const text = o.lookIn === 'formulas' ? doc.editText(sh.id, r, c) : (doc.viewOf(cell, r, c, 255).t ?? '')
      if (text && test(text)) hits.push({ sheet: sh.id, r, c, text })
    })
    if (o.byColumns) hits.sort((a, b) => a.c - b.c || a.r - b.r)
    for (const h of hits) {
      out.push(h)
      if (out.length >= limit) return out
    }
  }
  return out
}

/** Replace in formulas and constants; returns how many cells changed. */
export function replaceCells(doc: Doc, o: FindOptions & { replacement: string; all: boolean; at?: { sheet: number; r: number; c: number } }): number {
  const hits = findCells(doc, { ...o, lookIn: 'formulas' }).filter((m) => o.all || (o.at && m.sheet === o.at.sheet && m.r === o.at.r && m.c === o.at.c))
  if (!hits.length) return 0
  const flags = o.matchCase ? 'g' : 'gi'
  const re = o.wholeCell ? new RegExp(wildcardRegex(o.text).source, flags.replace('g', '')) : new RegExp(wildcardRegex(o.text).source.slice(1, -1), flags)
  doc.transact('Replace', (tx) => {
    for (const m of hits) {
      const next = o.wholeCell ? o.replacement : m.text.replace(re, () => o.replacement)
      doc.inputTx(tx, m.sheet, m.r, m.c, next)
    }
  })
  return hits.length
}
