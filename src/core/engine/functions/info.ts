// Information functions.
import { defineFunctions, toScalar } from '../evaluate'
import { CellError, ERR, isArray, isError, isLambda, isRef, type ErrorCode, type Value } from '../values'
import { colName } from '../../refs'
import { fn } from './helpers'

const ERROR_TYPE: Partial<Record<ErrorCode, number>> = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#GETTING_DATA': 8, '#SPILL!': 9, '#CONNECT!': 10, '#BLOCKED!': 11, '#UNKNOWN!': 12, '#FIELD!': 13, '#CALC!': 14 }

/** The format code CELL("format") reports: G, F0, ,2, C2, P0, S2, D1… */
function formatClass(code: string | undefined): string {
  if (!code || code === 'General') return 'G'
  const c = code.split(';')
  const first = c[0]
  const dec = (/\.(0+)/.exec(first)?.[1].length ?? 0).toString()
  const neg = c.length > 1 && /[\[(]/.test(c[1]) ? (/\[Red\]/i.test(c[1]) ? '-' : '') : ''
  const paren = c.length > 1 && /\(/.test(c[1]) ? '()' : ''
  if (/[yd]/i.test(first) && /m/i.test(first)) {
    if (/d.*m.*y|m.*d.*y/i.test(first)) return 'D1'
    if (/d.*mmm/i.test(first)) return 'D2'
    if (/mmm.*y/i.test(first)) return 'D3'
    return 'D4'
  }
  if (/h/i.test(first)) return /s/i.test(first) ? (/AM\/PM/i.test(first) ? 'D6' : 'D8') : /AM\/PM/i.test(first) ? 'D7' : 'D9'
  if (/%/.test(first)) return 'P' + dec
  if (/E\+/i.test(first)) return 'S' + dec
  if (/\$|€|£/.test(first)) return 'C' + dec + neg + paren
  if (/,/.test(first)) return ',' + dec + neg + paren
  return 'F' + dec + neg + paren
}

defineFunctions([
  fn('ISBLANK', 1, 1, 'v', ([v]) => v === null),
  fn('ISERR', 1, 1, 'v', ([v]) => v instanceof CellError && v !== ERR.NA),
  fn('ISERROR', 1, 1, 'v', ([v]) => v instanceof CellError),
  fn('ISNA', 1, 1, 'v', ([v]) => v === ERR.NA),
  fn('ISNUMBER', 1, 1, 'v', ([v]) => typeof v === 'number'),
  fn('ISTEXT', 1, 1, 'v', ([v]) => typeof v === 'string'),
  fn('ISNONTEXT', 1, 1, 'v', ([v]) => typeof v !== 'string'),
  fn('ISLOGICAL', 1, 1, 'v', ([v]) => typeof v === 'boolean'),
  fn('ISEVEN', 1, 1, 'n', ([x]) => Math.trunc(x) % 2 === 0),
  fn('ISODD', 1, 1, 'n', ([x]) => Math.abs(Math.trunc(x)) % 2 === 1),
  fn('ISREF', 1, 1, 'r', ([v]) => isRef(v)),
  fn('ISFORMULA', 1, 1, 'r', ([v], ctx) => {
    if (!isRef(v)) return isError(v) ? v : ERR.VALUE
    const a = v.areas[0]
    return !!ctx.wb.sheetById(a.sheet)?.get(a.r0, a.c0)?.f
  }),
  fn('N', 1, 1, 'v', ([v]) => (typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : v instanceof CellError ? v : 0)),
  fn('NA', 0, 0, '', () => ERR.NA),
  fn('ERROR.TYPE', 1, 1, 'v', ([v]) => (v instanceof CellError ? (ERROR_TYPE[v.code] ?? ERR.NA) : ERR.NA)),
  fn('TYPE', 1, 1, 'r', ([v], ctx) => {
    if (isLambda(v)) return 128
    if (isArray(v)) return 64
    if (isRef(v)) {
      const a = v.areas[0]
      if (v.areas.length > 1 || a.r0 !== a.r1 || a.c0 !== a.c1) return 64
    }
    const s = toScalar(v as Value, ctx)
    if (typeof s === 'number' || s === null) return 1
    if (typeof s === 'string') return 2
    if (typeof s === 'boolean') return 4
    return 16
  }),
  fn('SHEET', 0, 1, 'r', ([v], ctx) => {
    if (v === undefined) return ctx.wb.sheetIndex(ctx.sheet) + 1
    if (isRef(v)) return ctx.wb.sheetIndex(v.areas[0].sheet) + 1
    if (typeof v === 'string') {
      const s = ctx.wb.sheetByName(v)
      return s ? ctx.wb.sheetIndex(s.id) + 1 : ERR.NA
    }
    return ERR.VALUE
  }),
  fn('SHEETS', 0, 1, 'r', ([v], ctx) => {
    if (v === undefined) return ctx.wb.sheets.length
    if (isRef(v)) return new Set(v.areas.map((a) => a.sheet)).size
    return ERR.VALUE
  }),
  fn('CELL', 1, 2, 'sr', ([type, ref], ctx) => {
    const a = ref === undefined ? { sheet: ctx.sheet, r0: ctx.row, c0: ctx.col, r1: ctx.row, c1: ctx.col } : isRef(ref) ? ref.areas[0] : null
    if (!a) return ERR.VALUE
    const sheet = ctx.wb.sheetById(a.sheet)
    const cell = sheet?.get(a.r0, a.c0)
    const style = ctx.wb.styles.get(cell?.s)
    switch (type.toLowerCase()) {
      case 'address':
        return `$${colName(a.c0)}$${a.r0 + 1}`
      case 'col':
        return a.c0 + 1
      case 'row':
        return a.r0 + 1
      case 'contents': {
        const v = ctx.read(a.sheet, a.r0, a.c0)
        return v === null ? 0 : v
      }
      case 'type': {
        const v = ctx.read(a.sheet, a.r0, a.c0)
        return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v'
      }
      case 'filename':
        return ctx.wb.props.path ? `${ctx.wb.props.path.replace(/[^\\/]+$/, '')}[${ctx.wb.props.path.replace(/^.*[\\/]/, '')}]${sheet?.name ?? ''}` : ''
      case 'format':
        return formatClass(style.numFmt)
      case 'color':
        return /\[(red|color)/i.test(style.numFmt?.split(';')[1] ?? '') ? 1 : 0
      case 'parentheses':
        return /\(/.test(style.numFmt?.split(';')[0] ?? '') ? 1 : 0
      case 'prefix': {
        const v = ctx.read(a.sheet, a.r0, a.c0)
        if (typeof v !== 'string') return ''
        return style.hAlign === 'right' ? '"' : style.hAlign === 'center' ? '^' : style.hAlign === 'fill' ? '\\' : "'"
      }
      case 'protect':
        return style.locked === false ? 0 : 1
      case 'width':
        return Math.round((sheet?.colWidth(a.c0) ?? 64) / 7)
    }
    return ERR.VALUE
  }, { volatile: true }),
  fn('INFO', 1, 1, 's', ([t], ctx) => {
    switch (t.toLowerCase()) {
      case 'directory':
        return ctx.wb.props.path ? ctx.wb.props.path.replace(/[^\\/]+$/, '') : ''
      case 'numfile':
        return 1
      case 'origin':
        return '$A:$A$1'
      case 'osversion':
        return typeof navigator !== 'undefined' ? navigator.platform : 'Simple Cell'
      case 'recalc':
        return ctx.wb.calc.mode === 'auto' ? 'Automatic' : 'Manual'
      case 'release':
        return '16.0'
      case 'system':
        return typeof navigator !== 'undefined' && /Mac/.test(navigator.platform) ? 'mac' : 'pcdos'
    }
    return ERR.VALUE
  })
])
