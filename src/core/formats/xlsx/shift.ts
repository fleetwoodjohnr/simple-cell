// Keeps preserved XLSX content in step with inserted/deleted rows and columns: ranges in worksheet elements
// (conditional formatting, validation, hyperlinks, autofilter…), their formulas, drawing anchors, chart series
// formulas and comment anchors. Every change is recorded in the transaction, so undo restores it.
import type { Tx } from '../../doc'
import type { Workbook } from '../../model/workbook'
import { adjustFormula, shiftArea, type Shift } from '../../model/adjust'
import { parseFormula, printFormula } from '../../engine/parser'
import { areaName, parseArea, colIndex, colName } from '../../refs'
import { bytes, text } from '../zip'
import { resolveTarget } from '../opc'
import type { XlsxCarry } from './read'

function shiftRefList(list: string, sheet: number, s: Shift): string {
  return list
    .split(/\s+/)
    .filter(Boolean)
    .map((ref) => {
      const a = parseArea(ref, sheet)
      if (!a) return ref
      const b = shiftArea(a, s)
      return b ? areaName(b) : null
    })
    .filter((x): x is string => x !== null)
    .join(' ')
}

/** Rewrite a formula string (as stored in the file, relative to `origin`) for the shift. */
function shiftFormulaText(f: string, wb: Workbook, sheet: number, from: { row: number; col: number }, to: { row: number; col: number }, s: Shift): string {
  try {
    const sheetId = (n: string) => wb.sheetByName(n)?.id ?? null
    const ast = parseFormula('=' + f.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'), { ...from, sheetId })
    const next = adjustFormula(ast, sheet, from, to, s)
    if (next === ast && from.row === to.row && from.col === to.col) return f
    return printFormula(next, { ...to, sheetName: (id) => wb.sheetById(id)?.name ?? null, fileNames: (n) => n }).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  } catch {
    return f
  }
}

function shiftElement(xml: string, wb: Workbook, sheet: number, s: Shift): string | null {
  // Ranges in sqref/ref attributes and <xm:sqref> elements.
  let first: { row: number; col: number } | null = null
  let firstNew: { row: number; col: number } | null = null
  let out = xml.replace(/\b(sqref|ref)="([^"]*)"/g, (_m, k: string, v: string) => {
    const a = parseArea(v.split(/\s+/)[0] ?? '', sheet)
    if (a && !first) {
      first = { row: a.r0, col: a.c0 }
      const b = shiftArea(a, s)
      firstNew = b ? { row: b.r0, col: b.c0 } : first
    }
    const next = shiftRefList(v, sheet, s)
    return `${k}="${next}"`
  })
  out = out.replace(/(<xm:sqref>)([^<]*)(<\/xm:sqref>)/g, (_m, a: string, v: string, b: string) => `${a}${shiftRefList(v, sheet, s)}${b}`)
  // Empty range lists: the element covered only deleted cells.
  if (/\bsqref=""/.test(out) && /^<(\w+:)?(conditionalFormatting|dataValidation|selection)\b/.test(out)) return null
  // Formulas inside conditional formats and validations are relative to the top-left of their range.
  if (first)
    out = out.replace(/(<(?:\w+:)?(?:formula|formula1|formula2|f)>)([^<]*)(<\/(?:\w+:)?(?:formula|formula1|formula2|f)>)/g, (_m, a: string, f: string, b: string) => a + shiftFormulaText(f, wb, sheet, first!, firstNew!, s) + b)
  return out
}

function shiftAnchors(xml: string, s: Shift): string {
  const tag = s.axis === 'row' ? 'row' : 'col'
  return xml.replace(new RegExp(`(<(?:\\w+:)?${tag}>)(\\d+)(</(?:\\w+:)?${tag}>)`, 'g'), (_m, a: string, v: string, b: string) => {
    const n = Number(v)
    if (s.count > 0) return a + (n >= s.at ? n + s.count : n) + b
    const end = s.at - s.count - 1
    return a + (n > end ? n + s.count : n >= s.at ? s.at : n) + b
  })
}

function shiftVml(xml: string, s: Shift): string {
  const tag = s.axis === 'row' ? 'x:Row' : 'x:Column'
  return xml.replace(new RegExp(`(<${tag}>)(\\d+)(</${tag}>)`, 'g'), (_m, a: string, v: string, b: string) => {
    const n = Number(v)
    if (s.count > 0) return a + (n >= s.at ? n + s.count : n) + b
    const end = s.at - s.count - 1
    return a + (n > end ? n + s.count : n >= s.at ? s.at : n) + b
  })
}

function shiftComments(xml: string, s: Shift): string {
  return xml.replace(/(<comment\b[^>]*\bref=")([A-Z]+)(\d+)(")/g, (m, a: string, col: string, row: string, b: string) => {
    let r = Number(row) - 1
    let c = colIndex(col)
    const idx = s.axis === 'row' ? r : c
    let next = idx
    if (s.count > 0) next = idx >= s.at ? idx + s.count : idx
    else {
      const end = s.at - s.count - 1
      if (idx >= s.at && idx <= end) return m // a note on a deleted cell stays where Excel would show #REF!; kept
      next = idx > end ? idx + s.count : idx
    }
    if (s.axis === 'row') r = next
    else c = next
    return `${a}${colName(c)}${r + 1}${b}`
  })
}

export function shiftXlsxCarry(wb: Workbook, tx: Tx, s: Shift) {
  const carry = wb.carry as XlsxCarry | null
  if (!carry || !('workbookPath' in carry)) return
  const sc = carry.sheets[s.sheet]
  if (sc) {
    const elements = sc.elements.map((e) => ({ name: e.name, xml: shiftElement(e.xml, wb, s.sheet, s) })).filter((e): e is { name: string; xml: string } => e.xml !== null)
    if (elements.some((e, i) => e.xml !== sc.elements[i]?.xml) || elements.length !== sc.elements.length) tx.prop(sc, 'elements', elements)
    // Parts of this sheet: drawings (anchors), comments, VML notes.
    for (const r of sc.rels) {
      if (r.external) continue
      const p = resolveTarget(sc.path, r.target)
      const before = carry.files[p]
      if (!before) continue
      let after: string | null = null
      // ponytail: drawings and note shapes move only for whole rows/columns, not for Insert/Delete Cells.
      if (s.span && !/\/comments$/.test(r.type)) continue
      if (/\/drawing$/.test(r.type)) after = shiftAnchors(text(before), s)
      else if (/\/comments$/.test(r.type)) after = shiftComments(text(before), s)
      else if (/\/vmlDrawing$/.test(r.type)) after = shiftVml(text(before), s)
      if (after !== null && after !== text(before)) {
        const next = bytes(after)
        tx.custom(
          () => (carry.files[p] = next),
          () => (carry.files[p] = before)
        )
      }
    }
  }
  // Charts anywhere in the workbook may plot this sheet's cells: their series formulas are sheet-qualified.
  for (const cp of Object.keys(carry.files)) {
    if (!/\/charts\/[^/]+\.xml$/.test(cp)) continue
    const cb = carry.files[cp]
    const ct = text(cb)
    const nt = ct.replace(/(<(?:\w+:)?f>)([^<]*)(<\/(?:\w+:)?f>)/g, (_m, a: string, f: string, b: string) => a + shiftFormulaText(f, wb, s.sheet, { row: 0, col: 0 }, { row: 0, col: 0 }, s) + b)
    if (nt !== ct) {
      const next = bytes(nt)
      tx.custom(
        () => (carry.files[cp] = next),
        () => (carry.files[cp] = cb)
      )
    }
  }
}

export { areaName }
