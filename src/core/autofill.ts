// AutoFill (dragging the fill handle): series of numbers, dates, weekdays, months and "Item 1"-style text.
import type { Doc } from './doc'
import type { Area, Scalar } from './engine/values'
import { relocate } from './model/adjust'
import type { Cell } from './model/workbook'
import { dateSerial, daysInMonth, serialToDate } from './numfmt/dates'
import { isDateFormat } from './numfmt/format'

const LISTS = [
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['Q1', 'Q2', 'Q3', 'Q4']
]

export const customLists: string[][] = []

function matchCase(word: string, like: string) {
  if (like === like.toUpperCase()) return word.toUpperCase()
  if (like === like.toLowerCase()) return word.toLowerCase()
  return word
}

/** Values for positions `n` (0-based, can be negative when filling up/left) of a series seeded by `seed`. */
function seriesFor(doc: Doc, seed: (Cell | undefined)[]): ((n: number) => Cell | undefined) | null {
  const vals = seed.map((c) => c?.v ?? null)
  if (seed.some((c) => c?.f)) return null
  // Numbers (and dates): linear trend; one date steps by a day, one number copies.
  if (vals.every((v) => typeof v === 'number')) {
    const nums = vals as number[]
    const isDate = isDateFormat(doc.wb.styles.get(seed[0]?.s).numFmt)
    if (nums.length === 1) {
      if (!isDate) return null
      return (n) => ({ ...seed[0]!, v: nums[0] + n })
    }
    if (isDate && nums.length >= 2) {
      // Month/year steps when the seed dates share their day of month.
      const ds = nums.map((x) => serialToDate(x, doc.wb.date1904))
      const months = (ds[1].y - ds[0].y) * 12 + (ds[1].m - ds[0].m)
      const sameDay = ds.every((d) => d.d === ds[0].d) || ds.every((d) => d.d === daysInMonth(d.y, d.m))
      if (months !== 0 && sameDay && ds.every((d, i) => i === 0 || (d.y - ds[i - 1].y) * 12 + (d.m - ds[i - 1].m) === months)) {
        const endOfMonth = ds.every((d) => d.d === daysInMonth(d.y, d.m))
        return (n) => {
          const t = ds[0].m - 1 + months * n
          const y = ds[0].y + Math.floor(t / 12)
          const m = (((t % 12) + 12) % 12) + 1
          const d = endOfMonth ? daysInMonth(y, m) : Math.min(ds[0].d, daysInMonth(y, m))
          return { ...seed[n % seed.length >= 0 ? n % seed.length : 0]!, v: dateSerial(y, m, d, doc.wb.date1904) + (nums[0] % 1) }
        }
      }
    }
    // Least-squares line through the seed values.
    const k = nums.length
    const mx = (k - 1) / 2
    const my = nums.reduce((a, b) => a + b, 0) / k
    let sxy = 0
    let sxx = 0
    nums.forEach((y, x) => {
      sxy += (x - mx) * (y - my)
      sxx += (x - mx) * (x - mx)
    })
    const step = sxx ? sxy / sxx : 0
    const base = my - step * mx
    return (n) => ({ ...seed[((n % k) + k) % k]!, v: Number((base + step * n).toPrecision(15)) })
  }
  // Text in a known list (weekdays, months…).
  if (vals.every((v) => typeof v === 'string')) {
    const strs = vals as string[]
    for (const list of [...customLists, ...LISTS]) {
      const idx = strs.map((s) => list.findIndex((x) => x.toLowerCase() === s.toLowerCase()))
      if (idx.every((i) => i >= 0)) {
        const step = idx.length > 1 ? idx[1] - idx[0] : 1
        return (n) => {
          const i = (((idx[0] + step * n) % list.length) + list.length) % list.length
          return { ...seed[0]!, v: matchCase(list[i], strs[0]) }
        }
      }
    }
    // "Item 1", "Week 3": the trailing number counts.
    const m = strs.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s))
    if (m.every((x) => x && x[1] === m[0]![1] && x[3] === m[0]![3])) {
      const nums = m.map((x) => Number(x![2]))
      const step = nums.length > 1 ? nums[1] - nums[0] : 1
      const width = m[0]![2].length
      return (n) => {
        const v = Math.abs(nums[0] + step * n)
        return { ...seed[0]!, v: m[0]![1] + String(v).padStart(m[0]![2].startsWith('0') ? width : 1, '0') + m[0]![3] }
      }
    }
  }
  return null
}

/** Fill `dst` (which contains `src` and extends it in one direction) from the seed cells in `src`. */
export function autofill(doc: Doc, src: Area, dst: Area) {
  const sh = doc.sheet(src.sheet)
  const vertical = dst.c0 === src.c0 && dst.c1 === src.c1
  const forward = vertical ? dst.r1 > src.r1 : dst.c1 > src.c1
  const lines = vertical ? src.c1 - src.c0 + 1 : src.r1 - src.r0 + 1
  const seedLen = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1
  doc.transact('AutoFill', (tx) => {
    for (let l = 0; l < lines; l++) {
      const seed: (Cell | undefined)[] = []
      for (let i = 0; i < seedLen; i++) seed.push(vertical ? sh.get(src.r0 + i, src.c0 + l) : sh.get(src.r0 + l, src.c0 + i))
      const series = seriesFor(doc, seed)
      const total = vertical ? dst.r1 - dst.r0 + 1 : dst.c1 - dst.c0 + 1
      for (let k = seedLen; k < total; k++) {
        // n: position in the series; filling backwards counts down from the seed's start.
        const n = forward ? k : -(k - seedLen + 1)
        const r = vertical ? (forward ? src.r0 + k : src.r0 + n) : src.r0 + l
        const c = vertical ? src.c0 + l : forward ? src.c0 + k : src.c0 + n
        let cell: Cell | undefined
        if (series) cell = series(n)
        else {
          const s = seed[((n % seedLen) + seedLen) % seedLen]
          cell = s ? { ...s, spillOf: undefined } : undefined
          if (cell?.f) cell = { ...cell, f: { ...cell.f, ast: relocate(cell.f.ast, { row: r, col: c }), array: undefined } }
        }
        tx.setCell(src.sheet, r, c, cell ? { ...cell, v: cell.v as Scalar } : undefined)
      }
    }
  })
}
