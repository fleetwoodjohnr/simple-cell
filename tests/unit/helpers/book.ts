import { Doc } from '@/core/doc'
import { parseCell } from '@/core/refs'
import { CellError, type Scalar } from '@/core/engine/values'

/** A small test workbook: `b.set('A1', 5)`, `b.set('B1', '=A1*2')`, `b.get('B1')`. */
export function book(cells: Record<string, string | number | boolean> = {}) {
  const doc = new Doc()
  const sheet = doc.wb.sheets[0].id
  const at = (ref: string) => {
    const [s, a] = ref.includes('!') ? ref.split('!') : [null, ref]
    const p = parseCell(a)!
    return { sheet: s ? doc.wb.sheetByName(s)!.id : sheet, ...p }
  }
  const api = {
    doc,
    sheet,
    set(ref: string, v: string | number | boolean) {
      const p = at(ref)
      doc.input(p.sheet, p.r, p.c, typeof v === 'string' ? v : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v))
      return api
    },
    get(ref: string): Scalar | string {
      const p = at(ref)
      const v = doc.value(p.sheet, p.r, p.c)
      return v instanceof CellError ? v.code : v
    },
    text(ref: string) {
      const p = at(ref)
      return doc.display(p.sheet, p.r, p.c)?.t ?? ''
    },
    formula(ref: string) {
      const p = at(ref)
      return doc.formulaText(p.sheet, p.r, p.c)
    },
    /** Evaluate a formula in an empty cell far away and return its value. */
    calc(f: string): Scalar | string {
      doc.input(sheet, 999, 50, f.startsWith('=') ? f : '=' + f)
      const v = doc.value(sheet, 999, 50)
      return v instanceof CellError ? v.code : v
    }
  }
  for (const [k, v] of Object.entries(cells)) api.set(k, v)
  return api
}
