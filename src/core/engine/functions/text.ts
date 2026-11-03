// Text functions.
import { defineFunctions, toArray, toScalar, type EvalContext } from '../evaluate'
import { CellError, ERR, isArray, isError, isRef, makeArray, numberToText, parseNumber, toText, type Scalar, type Value } from '../values'
import { formatValue } from '../../numfmt/format'
import { parseDateTime } from '../../numfmt/dates'
import { allValues, fn, MANY, roundHalfAway, wildcardRegex } from './helpers'

const MAX_TEXT = 32767
const tooLong = (s: string) => (s.length > MAX_TEXT ? ERR.VALUE : s)

// CHAR/CODE use the Windows-1252 code page, as Excel for Windows does.
const CP1252 = new TextDecoder('windows-1252').decode(new Uint8Array(Array.from({ length: 256 }, (_, i) => i)))
const CP1252_CODE = new Map([...CP1252].map((ch, i) => [ch, i]))

const toHalf = (s: string) => s.replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/　/g, ' ')
const toFull = (s: string) => s.replace(/[!-~]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xfee0)).replace(/ /g, '　')

function valueOf(s: string, ctx: EvalContext): number | CellError {
  const t = s.trim()
  if (t === '') return 0
  const n = parseNumber(t, (x) => parseDateTime(x, ctx.wb.date1904))
  return n === null ? ERR.VALUE : n
}

function fixed(x: number, d: number, noCommas: boolean): string | CellError {
  d = Math.trunc(d)
  if (d > 127) return ERR.VALUE
  const r = roundHalfAway(x, d)
  const s = Math.abs(r).toFixed(Math.max(0, d))
  const [i, f] = s.split('.')
  const int = noCommas ? i : i.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return (r < 0 ? '-' : '') + int + (f ? '.' + f : '')
}

function find(needle: string, hay: string, start: number, insensitive: boolean): number | CellError {
  start = Math.trunc(start)
  if (start < 1 || start > hay.length + 1) return ERR.VALUE
  if (needle === '') return start
  if (!insensitive) {
    const i = hay.indexOf(needle, start - 1)
    return i < 0 ? ERR.VALUE : i + 1
  }
  // SEARCH: wildcards, case-insensitive, earliest match
  const re = new RegExp(wildcardRegex(needle).source.slice(1, -1), 'i')
  const m = re.exec(hay.slice(start - 1))
  return m ? m.index + start : ERR.VALUE
}

/** Text of each delimiter (TEXTBEFORE, TEXTAFTER, TEXTSPLIT accept an array of delimiters). */
function delimiters(v: Value, ctx: EvalContext): string[] | CellError {
  if (v === undefined || v === null) return []
  const vals = isRef(v) || isArray(v) ? allValues([v], ctx) : [v as Scalar]
  if (isError(vals)) return vals
  const out: string[] = []
  for (const x of vals) {
    const t = toText(x)
    if (isError(t)) return t
    out.push(t)
  }
  return out
}

function splitAround(text: string, delims: string[], instance: number, insensitive: boolean, matchEnd: boolean): { before: string; after: string } | null {
  const hay = insensitive ? text.toLowerCase() : text
  const ds = delims.map((d) => (insensitive ? d.toLowerCase() : d))
  // positions of every delimiter occurrence (earliest first; longest delimiter wins at a position)
  const hits: { i: number; len: number }[] = []
  for (let i = 0; i < hay.length; ) {
    let best = -1
    for (const d of ds) if (d && hay.startsWith(d, i) && d.length > best) best = d.length
    if (best > 0) {
      hits.push({ i, len: best })
      i += best
    } else i++
  }
  if (ds.some((d) => d === '')) {
    // An empty delimiter matches at the start (positive instance) or end (negative).
    return instance > 0 ? { before: '', after: text } : { before: text, after: '' }
  }
  if (matchEnd) {
    if (instance > 0) hits.push({ i: text.length, len: 0 })
    else hits.unshift({ i: 0, len: 0 })
  }
  const k = instance > 0 ? instance - 1 : hits.length + instance
  const h = hits[k]
  if (!h || k < 0) return null
  return { before: text.slice(0, h.i), after: text.slice(h.i + h.len) }
}

function regex(pattern: string, caseInsensitive: boolean, global = false): RegExp | CellError {
  try {
    return new RegExp(pattern, (caseInsensitive ? 'i' : '') + (global ? 'g' : '') + 'u')
  } catch {
    return ERR.VALUE
  }
}

function valueToText(v: Scalar, strict: boolean): string {
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v
  if (v === null) return ''
  if (v instanceof CellError) return v.code
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return numberToText(v)
}

defineFunctions([
  fn('LEN', 1, 1, 's', ([s]) => s.length),
  fn('LENB', 1, 1, 's', ([s]) => s.length),
  fn('LOWER', 1, 1, 's', ([s]) => s.toLowerCase()),
  fn('UPPER', 1, 1, 's', ([s]) => s.toUpperCase()),
  fn('PROPER', 1, 1, 's', ([s]) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_m: string, a: string, b: string) => a + b.toUpperCase())),
  fn('TRIM', 1, 1, 's', ([s]) => s.replace(/ +/g, ' ').replace(/^ | $/g, '')),
  fn('CLEAN', 1, 1, 's', ([s]) => s.replace(/[\x00-\x1f]/g, '')),
  fn('LEFT', 1, 2, 'sn', ([s, n = 1]) => (n < 0 ? ERR.VALUE : s.slice(0, Math.trunc(n)))),
  fn('LEFTB', 1, 2, 'sn', ([s, n = 1]) => (n < 0 ? ERR.VALUE : s.slice(0, Math.trunc(n)))),
  fn('RIGHT', 1, 2, 'sn', ([s, n = 1]) => (n < 0 ? ERR.VALUE : Math.trunc(n) === 0 ? '' : s.slice(-Math.trunc(n)))),
  fn('RIGHTB', 1, 2, 'sn', ([s, n = 1]) => (n < 0 ? ERR.VALUE : Math.trunc(n) === 0 ? '' : s.slice(-Math.trunc(n)))),
  fn('MID', 3, 3, 'snn', ([s, start, n]) => (start < 1 || n < 0 ? ERR.VALUE : s.substr(Math.trunc(start) - 1, Math.trunc(n)))),
  fn('MIDB', 3, 3, 'snn', ([s, start, n]) => (start < 1 || n < 0 ? ERR.VALUE : s.substr(Math.trunc(start) - 1, Math.trunc(n)))),
  fn('FIND', 2, 3, 'ssn', ([a, b, start = 1]) => find(a, b, start, false)),
  fn('FINDB', 2, 3, 'ssn', ([a, b, start = 1]) => find(a, b, start, false)),
  fn('SEARCH', 2, 3, 'ssn', ([a, b, start = 1]) => find(a, b, start, true)),
  fn('SEARCHB', 2, 3, 'ssn', ([a, b, start = 1]) => find(a, b, start, true)),
  fn('REPLACE', 4, 4, 'snns', ([s, start, n, rep]) => (start < 1 || n < 0 ? ERR.VALUE : tooLong(s.slice(0, Math.trunc(start) - 1) + rep + s.slice(Math.trunc(start) - 1 + Math.trunc(n))))),
  fn('REPLACEB', 4, 4, 'snns', ([s, start, n, rep]) => (start < 1 || n < 0 ? ERR.VALUE : tooLong(s.slice(0, Math.trunc(start) - 1) + rep + s.slice(Math.trunc(start) - 1 + Math.trunc(n))))),
  fn('SUBSTITUTE', 3, 4, 'sssn', ([s, old, rep, inst]) => {
    if (old === '') return s
    if (inst === undefined) return tooLong(s.split(old).join(rep))
    const k = Math.trunc(inst)
    if (k < 1) return ERR.VALUE
    let idx = -1
    for (let i = 0; i < k; i++) {
      idx = s.indexOf(old, idx + 1)
      if (idx < 0) return s
    }
    return tooLong(s.slice(0, idx) + rep + s.slice(idx + old.length))
  }),
  fn('REPT', 2, 2, 'sn', ([s, n]) => {
    n = Math.trunc(n)
    if (n < 0 || s.length * n > MAX_TEXT) return ERR.VALUE
    return s.repeat(n)
  }),
  fn('EXACT', 2, 2, 's', ([a, b]) => a === b),
  fn('CHAR', 1, 1, 'n', ([n]) => {
    n = Math.trunc(n)
    return n < 1 || n > 255 ? ERR.VALUE : CP1252[n]
  }),
  fn('CODE', 1, 1, 's', ([s]) => (s === '' ? ERR.VALUE : (CP1252_CODE.get(s[0]) ?? 63))),
  fn('UNICHAR', 1, 1, 'n', ([n]) => {
    n = Math.trunc(n)
    if (n < 1 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) return ERR.VALUE
    return String.fromCodePoint(n)
  }),
  fn('UNICODE', 1, 1, 's', ([s]) => (s === '' ? ERR.VALUE : s.codePointAt(0)!)),
  fn('T', 1, 1, 'v', ([v]) => (typeof v === 'string' ? v : v instanceof CellError ? v : '')),
  fn('VALUE', 1, 1, 'v', ([v], ctx) => {
    if (typeof v === 'number') return v
    if (v instanceof CellError) return v
    if (typeof v === 'boolean') return ERR.VALUE
    return valueOf(v ?? '', ctx)
  }),
  fn('NUMBERVALUE', 1, 3, 'sss', ([s, dec, group]) => {
    const d = dec === undefined ? '.' : dec[0]
    const g = group === undefined ? ',' : group[0]
    if (dec !== undefined && dec === '') return ERR.VALUE
    let t = s.replace(/\s/g, '')
    if (t === '') return 0
    let pct = 0
    while (t.endsWith('%')) {
      pct++
      t = t.slice(0, -1)
    }
    const di = t.indexOf(d)
    if (di >= 0 && t.slice(di + 1).includes(g ?? '\u0000')) return ERR.VALUE
    const parts = di >= 0 ? [t.slice(0, di), t.slice(di + 1)] : [t]
    if (parts[1]?.includes(d)) return ERR.VALUE
    const int = g ? parts[0].split(g).join('') : parts[0]
    const n = Number(int + (parts.length > 1 ? '.' + parts[1] : ''))
    if (!/^[+-]?\d*\.?\d*(e[+-]?\d+)?$/i.test(int + (parts.length > 1 ? '.' + parts[1] : '')) || Number.isNaN(n)) return ERR.VALUE
    return n / Math.pow(100, pct)
  }),
  fn('TEXT', 2, 2, 'vs', ([v, code], ctx) => {
    if (v instanceof CellError) return v
    let x: Scalar = v
    if (typeof v === 'string') {
      const n = parseNumber(v, (s) => parseDateTime(s, ctx.wb.date1904))
      if (n !== null) x = n
    }
    if (typeof x === 'boolean') return x ? 'TRUE' : 'FALSE'
    return formatValue(x, code, ctx.wb.date1904, 255).text
  }),
  fn('FIXED', 1, 3, 'nnb', ([x, d = 2, noCommas = false]) => fixed(x, d, noCommas)),
  fn('DOLLAR', 1, 2, 'nn', ([x, d = 2]) => {
    const s = fixed(Math.abs(x), d, false)
    if (isError(s)) return s
    return x < 0 && Number(s.replace(/,/g, '')) !== 0 ? `($${s})` : `$${s}`
  }),
  fn('CONCAT', 1, MANY, 'r', (args, ctx) => {
    const vals = allValues(args, ctx)
    if (isError(vals)) return vals
    let out = ''
    for (const v of vals) {
      if (v instanceof CellError) return v
      out += toText(v) as string
    }
    return tooLong(out)
  }),
  fn('CONCATENATE', 1, MANY, 's', (args) => tooLong(args.join(''))),
  fn('TEXTJOIN', 3, MANY, 'rb|r', ([delim, ignoreEmpty, ...rest], ctx) => {
    const ds = delimiters(delim, ctx)
    if (isError(ds)) return ds
    const vals = allValues(rest, ctx, true)
    if (isError(vals)) return vals
    const parts: string[] = []
    for (const v of vals) {
      if (v instanceof CellError) return v
      const t = toText(v) as string
      if (ignoreEmpty && t === '') continue
      parts.push(t)
    }
    let out = ''
    parts.forEach((p, i) => (out += (i ? (ds.length ? ds[(i - 1) % ds.length] : '') : '') + p))
    return tooLong(out)
  }),
  fn('TEXTBEFORE', 2, 6, 'srnnnr', ([s, delim, inst = 1, mode = 0, end = 0, notFound], ctx) => {
    const ds = delimiters(delim, ctx)
    if (isError(ds)) return ds
    inst = Math.trunc(inst)
    if (inst === 0 || Math.abs(inst) > s.length + 1) return ERR.VALUE
    const r = splitAround(s, ds, inst, mode === 1, end === 1)
    return r ? r.before : notFound === undefined ? ERR.NA : toScalar(notFound, ctx)
  }),
  fn('TEXTAFTER', 2, 6, 'srnnnr', ([s, delim, inst = 1, mode = 0, end = 0, notFound], ctx) => {
    const ds = delimiters(delim, ctx)
    if (isError(ds)) return ds
    inst = Math.trunc(inst)
    if (inst === 0 || Math.abs(inst) > s.length + 1) return ERR.VALUE
    const r = splitAround(s, ds, inst, mode === 1, end === 1)
    return r ? r.after : notFound === undefined ? ERR.NA : toScalar(notFound, ctx)
  }),
  fn('TEXTSPLIT', 2, 6, 'srrbnr', ([s, colD, rowD, ignoreEmpty = false, mode = 0, pad], ctx) => {
    const cd = delimiters(colD, ctx)
    const rd = delimiters(rowD, ctx)
    if (isError(cd)) return cd
    if (isError(rd)) return rd
    if (!cd.length && !rd.length) return ERR.VALUE
    const split = (t: string, ds: string[]) => {
      const live = ds.filter((d) => d !== '')
      if (!live.length) return [t]
      const re = new RegExp(live.sort((a, b) => b.length - a.length).map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), mode === 1 ? 'i' : '')
      return t.split(re)
    }
    let rows = rd.length ? split(s, rd) : [s]
    if (ignoreEmpty) rows = rows.filter((r) => r !== '')
    const cells = rows.map((r) => {
      const cs = cd.length ? split(r, cd) : [r]
      return ignoreEmpty ? cs.filter((x) => x !== '') : cs
    })
    const cols = Math.max(1, ...cells.map((r) => r.length))
    const padV: Scalar = pad === undefined ? ERR.NA : toScalar(pad, ctx)
    const data: Scalar[] = []
    for (const r of cells) for (let c = 0; c < cols; c++) data.push(c < r.length ? r[c] : padV)
    return data.length ? makeArray(cells.length || 1, cols, data.length ? data : ['']) : ERR.CALC
  }),
  fn('VALUETOTEXT', 1, 2, 'vn', ([v, f = 0]) => (f !== 0 && f !== 1 ? ERR.VALUE : valueToText(v, f === 1))),
  fn('ARRAYTOTEXT', 1, 2, 'rn', ([v, f = 0], ctx) => {
    if (f !== 0 && f !== 1) return ERR.VALUE
    const a = toArray(v, ctx)
    if (isError(a)) return a
    if (f === 0) return a.data.map((x) => valueToText(x, false)).join(', ')
    const rows: string[] = []
    for (let r = 0; r < a.rows; r++) rows.push(a.data.slice(r * a.cols, (r + 1) * a.cols).map((x) => valueToText(x, true)).join(','))
    return `{${rows.join(';')}}`
  }),
  fn('REGEXTEST', 2, 3, 'ssn', ([s, p, cs = 0]) => {
    const re = regex(p, cs === 1)
    return isError(re) ? re : re.test(s)
  }),
  fn('REGEXEXTRACT', 2, 4, 'ssnn', ([s, p, mode = 0, cs = 0]) => {
    const re = regex(p, cs === 1, mode === 1)
    if (isError(re)) return re
    if (mode === 0) {
      const m = re.exec(s)
      return m ? m[0] : ERR.NA
    }
    if (mode === 1) {
      const all = [...s.matchAll(re)].map((m) => m[0])
      return all.length ? makeArray(all.length, 1, all) : ERR.NA
    }
    if (mode === 2) {
      const m = re.exec(s)
      if (!m) return ERR.NA
      const groups = m.slice(1).map((g) => g ?? '')
      return groups.length ? makeArray(1, groups.length, groups) : ERR.NA
    }
    return ERR.VALUE
  }),
  fn('REGEXREPLACE', 3, 5, 'sssnn', ([s, p, rep, occ = 0, cs = 0]) => {
    const re = regex(p, cs === 1, true)
    if (isError(re)) return re
    const r = rep.replace(/\$(\d+)/g, '$$$1')
    if (occ === 0) return tooLong(s.replace(re, r))
    const all = [...s.matchAll(re)]
    const k = occ > 0 ? occ - 1 : all.length + occ
    const m = all[k]
    if (!m) return s
    return tooLong(s.slice(0, m.index) + m[0].replace(new RegExp(re.source, re.flags.replace('g', '')), r) + s.slice(m.index! + m[0].length))
  }),
  fn('ASC', 1, 1, 's', ([s]) => toHalf(s)),
  fn('DBCS', 1, 1, 's', ([s]) => toFull(s)),
  fn('JIS', 1, 1, 's', ([s]) => toFull(s)),
  fn('PHONETIC', 1, 1, 'r', ([v], ctx) => {
    // ponytail: no furigana data is stored, so PHONETIC returns the text itself (Excel does the same when there is none).
    const s = toScalar(v, ctx)
    return s instanceof CellError ? s : typeof s === 'string' ? s : ''
  }),
  fn('BAHTTEXT', 1, 1, 'n', ([x]) => bahtText(x))
])

// BAHTTEXT: the amount in Thai words, baht and satang.
const TH_DIGITS = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า']
const TH_PLACES = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน']
function thaiInt(n: number): string {
  if (n === 0) return ''
  if (n >= 1_000_000) return thaiInt(Math.floor(n / 1_000_000)) + 'ล้าน' + thaiInt(n % 1_000_000)
  const s = String(n)
  let out = ''
  for (let i = 0; i < s.length; i++) {
    const d = Number(s[i])
    const place = s.length - 1 - i
    if (d === 0) continue
    if (place === 1 && d === 1) out += 'สิบ'
    else if (place === 1 && d === 2) out += 'ยี่สิบ'
    else if (place === 0 && d === 1 && s.length > 1) out += 'เอ็ด'
    else out += TH_DIGITS[d] + TH_PLACES[place]
  }
  return out
}
function bahtText(x: number): string {
  const neg = x < 0
  const v = Math.round(Math.abs(x) * 100)
  const baht = Math.floor(v / 100)
  const satang = v % 100
  let out = ''
  if (baht === 0 && satang === 0) out = 'ศูนย์บาทถ้วน'
  else {
    if (baht) out += thaiInt(baht) + 'บาท'
    out += satang ? thaiInt(satang) + 'สตางค์' : 'ถ้วน'
  }
  return (neg ? 'ลบ' : '') + out
}
