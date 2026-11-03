// Engineering functions: Bessel functions, number bases, bit operations, complex numbers, CONVERT.
import { defineFunctions, toArray, type EvalContext } from '../evaluate'
import { CellError, ERR, isError, isRef, isArray, numberToText, toText, type Scalar, type Value } from '../values'
import { checkNum, fn, MANY } from './helpers'
import { erf, erfc } from './stats-core'

// ---------- Bessel functions (Numerical Recipes approximations, as Excel's accuracy) ----------

function j0(x: number) {
  const ax = Math.abs(x)
  if (ax < 8) {
    const y = x * x
    const a = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * -184.9052456))))
    const b = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y))))
    return a / b
  }
  const z = 8 / ax
  const y = z * z
  const xx = ax - 0.785398164
  const a = 1 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)))
  const b = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 - y * 0.934935152e-7)))
  return Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * a - z * Math.sin(xx) * b)
}
function j1(x: number) {
  const ax = Math.abs(x)
  if (ax < 8) {
    const y = x * x
    const a = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.4826 + y * -30.16036606)))))
    const b = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y))))
    return a / b
  }
  const z = 8 / ax
  const y = z * z
  const xx = ax - 2.356194491
  const a = 1 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)))
  const b = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)))
  const ans = Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * a - z * Math.sin(xx) * b)
  return x < 0 ? -ans : ans
}
function jn(n: number, x: number): number {
  if (n === 0) return j0(x)
  if (n === 1) return j1(x)
  const ax = Math.abs(x)
  if (ax === 0) return 0
  let ans: number
  if (ax > n) {
    let bjm = j0(ax)
    let bj = j1(ax)
    for (let j = 1; j < n; j++) {
      const bjp = (j * 2) / ax * bj - bjm
      bjm = bj
      bj = bjp
    }
    ans = bj
  } else {
    const tox = 2 / ax
    const m = 2 * Math.floor((n + Math.floor(Math.sqrt(40 * n))) / 2)
    let jsum = false
    let bjp = 0
    let s = 0
    let bj = 1
    ans = 0
    for (let j = m; j > 0; j--) {
      const bjm = j * tox * bj - bjp
      bjp = bj
      bj = bjm
      if (Math.abs(bj) > 1e10) {
        bj *= 1e-10
        bjp *= 1e-10
        ans *= 1e-10
        s *= 1e-10
      }
      if (jsum) s += bj
      jsum = !jsum
      if (j === n) ans = bjp
    }
    s = 2 * s - bj
    ans /= s
  }
  return x < 0 && n % 2 ? -ans : ans
}
function y0(x: number) {
  if (x < 8) {
    const y = x * x
    const a = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))))
    const b = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.2647 + y * (226.1030244 + y))))
    return a / b + 0.636619772 * j0(x) * Math.log(x)
  }
  const z = 8 / x
  const y = z * z
  const xx = x - 0.785398164
  const a = 1 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)))
  const b = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * -0.934945152e-7)))
  return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * a + z * Math.cos(xx) * b)
}
function y1(x: number) {
  if (x < 8) {
    const y = x * x
    const a = x * (-0.4900604943e13 + y * (0.127527439e13 + y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))))
    const b = 0.249958057e14 + y * (0.4244419664e12 + y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.1020426050e6 + y * (0.3549632885e3 + y)))))
    return a / b + 0.636619772 * (j1(x) * Math.log(x) - 1 / x)
  }
  const z = 8 / x
  const y = z * z
  const xx = x - 2.356194491
  const a = 1 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)))
  const b = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)))
  return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * a + z * Math.cos(xx) * b)
}
function yn(n: number, x: number) {
  if (n === 0) return y0(x)
  if (n === 1) return y1(x)
  const tox = 2 / x
  let bym = y0(x)
  let by = y1(x)
  for (let j = 1; j < n; j++) {
    const byp = j * tox * by - bym
    bym = by
    by = byp
  }
  return by
}
function i0(x: number) {
  const ax = Math.abs(x)
  if (ax < 3.75) {
    const y = (x / 3.75) ** 2
    return 1 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))))
  }
  const y = 3.75 / ax
  return (Math.exp(ax) / Math.sqrt(ax)) * (0.39894228 + y * (0.1328592e-1 + y * (0.225319e-2 + y * (-0.157565e-2 + y * (0.916281e-2 + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1 + y * 0.392377e-2))))))))
}
function i1(x: number) {
  const ax = Math.abs(x)
  let ans: number
  if (ax < 3.75) {
    const y = (x / 3.75) ** 2
    ans = ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))))
  } else {
    const y = 3.75 / ax
    ans = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1 - y * 0.420059e-2))
    ans = 0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2 + y * (0.163801e-2 + y * (-0.1031555e-1 + y * ans))))
    ans *= Math.exp(ax) / Math.sqrt(ax)
  }
  return x < 0 ? -ans : ans
}
function inBessel(n: number, x: number) {
  if (n === 0) return i0(x)
  if (n === 1) return i1(x)
  if (x === 0) return 0
  const tox = 2 / Math.abs(x)
  let bip = 0
  let bi = 1
  let ans = 0
  const m = 2 * (n + Math.floor(Math.sqrt(40 * n)))
  for (let j = m; j > 0; j--) {
    const bim = bip + j * tox * bi
    bip = bi
    bi = bim
    if (Math.abs(bi) > 1e10) {
      ans *= 1e-10
      bi *= 1e-10
      bip *= 1e-10
    }
    if (j === n) ans = bip
  }
  ans *= i0(x) / bi
  return x < 0 && n % 2 ? -ans : ans
}
function k0(x: number) {
  if (x <= 2) {
    const y = (x * x) / 4
    return -Math.log(x / 2) * i0(x) + (-0.57721566 + y * (0.4227842 + y * (0.23069756 + y * (0.348859e-1 + y * (0.262698e-2 + y * (0.1075e-3 + y * 0.74e-5))))))
  }
  const y = 2 / x
  return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (-0.7832358e-1 + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.25154e-2 + y * 0.53208e-3))))))
}
function k1(x: number) {
  if (x <= 2) {
    const y = (x * x) / 4
    return Math.log(x / 2) * i1(x) + (1 / x) * (1 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * -0.4686e-4))))))
  }
  const y = 2 / x
  return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (0.23498619 + y * (-0.365562e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * -0.68245e-3))))))
}
function kn(n: number, x: number) {
  if (n === 0) return k0(x)
  if (n === 1) return k1(x)
  const tox = 2 / x
  let bkm = k0(x)
  let bk = k1(x)
  for (let j = 1; j < n; j++) {
    const bkp = bkm + j * tox * bk
    bkm = bk
    bk = bkp
  }
  return bk
}

// ---------- number bases ----------

const BASES = { 2: { bits: 10, re: /^[01]*$/ }, 8: { bits: 30, re: /^[0-7]*$/ }, 16: { bits: 40, re: /^[0-9a-f]*$/i } } as const

function argText(v: Scalar): string | CellError {
  if (typeof v === 'boolean') return ERR.VALUE
  const t = toText(v)
  return t
}

/** Two's-complement text of the base (10 digits max) → number. */
function fromBase(text: string, base: 2 | 8 | 16): number | CellError {
  const t = text.trim()
  if (t.length > 10 || !BASES[base].re.test(t)) return ERR.NUM
  if (!t) return 0
  const v = parseInt(t, base)
  const bits = BASES[base].bits
  return t.length === 10 && v >= 2 ** (bits - 1) ? v - 2 ** bits : v
}

function toBase(n: number, base: 2 | 8 | 16, places: number | undefined): string | CellError {
  const bits = BASES[base].bits
  const min = -(2 ** (bits - 1))
  const max = 2 ** (bits - 1) - 1
  n = Math.trunc(n)
  if (n < min || n > max) return ERR.NUM
  if (n < 0) return (n + 2 ** bits).toString(base).toUpperCase()
  const s = n.toString(base).toUpperCase()
  if (places === undefined) return s
  const p = Math.trunc(places)
  if (p < 0 || p > 10 || s.length > p) return ERR.NUM
  return s.padStart(p, '0')
}

function convertBase(from: 2 | 8 | 16, to: 2 | 8 | 16 | 10) {
  return (args: Scalar[]) => {
    const t = argText(args[0])
    if (isError(t)) return t
    const n = fromBase(t, from)
    if (isError(n)) return n
    if (to === 10) return n
    const places = args[1] === undefined || args[1] === null ? undefined : Number(args[1])
    return toBase(n, to, places)
  }
}

// ---------- bits ----------

const MAX48 = 2 ** 48
function bitArg(x: number): bigint | CellError {
  if (x < 0 || x >= MAX48 || !Number.isInteger(x)) return ERR.NUM
  return BigInt(x)
}

// ---------- complex numbers ----------

interface Cx {
  re: number
  im: number
  sfx: 'i' | 'j'
}

const NUM = '(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][+-]?\\d+)?'
const CX_RE = new RegExp(`^([+-]?${NUM})?(?:([+-])(${NUM})?([ij]))?$|^([+-]?${NUM})?([ij])$`)

function parseCx(v: Scalar): Cx | CellError {
  if (v instanceof CellError) return v
  if (typeof v === 'number') return { re: v, im: 0, sfx: 'i' }
  if (typeof v === 'boolean') return ERR.VALUE
  const t = (v ?? '').trim()
  if (t === '') return { re: 0, im: 0, sfx: 'i' }
  const m = CX_RE.exec(t)
  if (!m) return ERR.NUM
  if (m[6]) {
    // pure imaginary: "3i", "-i", "i"
    const coef = m[5] === undefined || m[5] === '+' ? 1 : m[5] === '-' ? -1 : Number(m[5])
    return { re: 0, im: coef, sfx: m[6] as 'i' | 'j' }
  }
  const re = m[1] !== undefined ? Number(m[1]) : 0
  if (!m[2]) return { re, im: 0, sfx: 'i' }
  const im = (m[2] === '-' ? -1 : 1) * (m[3] !== undefined ? Number(m[3]) : 1)
  return { re, im, sfx: m[4] as 'i' | 'j' }
}

function fmtCx(z: Cx): string | CellError {
  if (!Number.isFinite(z.re) || !Number.isFinite(z.im)) return ERR.NUM
  const r = Number(z.re.toPrecision(15))
  const i = Number(z.im.toPrecision(15))
  if (i === 0) return numberToText(r)
  const imPart = i === 1 ? '' : i === -1 ? '-' : numberToText(i)
  if (r === 0) return imPart + z.sfx
  return numberToText(r) + (i > 0 ? '+' : '') + imPart + z.sfx
}

const cx = (re: number, im: number, sfx: Cx['sfx'] = 'i'): Cx => ({ re, im, sfx })
const cmul = (a: Cx, b: Cx) => cx(a.re * b.re - a.im * b.im, a.re * b.im + a.im * b.re, a.sfx)
const cdiv = (a: Cx, b: Cx): Cx | CellError => {
  const d = b.re * b.re + b.im * b.im
  if (d === 0) return ERR.NUM
  return cx((a.re * b.re + a.im * b.im) / d, (a.im * b.re - a.re * b.im) / d, a.sfx)
}
const cabs = (z: Cx) => Math.hypot(z.re, z.im)
const cexp = (z: Cx) => cx(Math.exp(z.re) * Math.cos(z.im), Math.exp(z.re) * Math.sin(z.im), z.sfx)
const cln = (z: Cx): Cx | CellError => (z.re === 0 && z.im === 0 ? ERR.NUM : cx(Math.log(cabs(z)), Math.atan2(z.im, z.re), z.sfx))
const csin = (z: Cx) => cx(Math.sin(z.re) * Math.cosh(z.im), Math.cos(z.re) * Math.sinh(z.im), z.sfx)
const ccos = (z: Cx) => cx(Math.cos(z.re) * Math.cosh(z.im), -Math.sin(z.re) * Math.sinh(z.im), z.sfx)
const csinh = (z: Cx) => cx(Math.sinh(z.re) * Math.cos(z.im), Math.cosh(z.re) * Math.sin(z.im), z.sfx)
const ccosh = (z: Cx) => cx(Math.cosh(z.re) * Math.cos(z.im), Math.sinh(z.re) * Math.sin(z.im), z.sfx)
const one = (z: Cx) => cx(1, 0, z.sfx)

function cx1(name: string, f: (z: Cx) => Cx | CellError | number) {
  return fn(name, 1, 1, 'v', ([v]) => {
    const z = parseCx(v)
    if (isError(z)) return z
    const r = f(z)
    if (typeof r === 'number') return checkNum(r)
    return isError(r) ? r : fmtCx(r)
  })
}

function cxList(args: Value[], ctx: EvalContext): Cx[] | CellError {
  const out: Cx[] = []
  for (const a of args) {
    const vals = isRef(a) || isArray(a) ? (toArray(a, ctx) as { data: Scalar[] }).data : [a as Scalar]
    for (const v of vals) {
      if (v === null && (isRef(a) || isArray(a))) continue
      const z = parseCx(v)
      if (isError(z)) return z
      out.push(z)
    }
  }
  if (out.some((z) => z.re !== 0 || z.im !== 0) && new Set(out.filter((z) => z.im !== 0).map((z) => z.sfx)).size > 1) return ERR.VALUE
  return out
}

// ---------- CONVERT ----------

type Unit = { g: string; f: number; prefix?: boolean; binary?: boolean }
const U: Record<string, Unit> = {}
const add = (g: string, names: string[], f: number, prefix = true, binary = false) => names.forEach((n) => (U[n] = { g, f, prefix, binary }))
// Mass (kg base: g = 1e-3)
add('mass', ['g'], 1e-3)
add('mass', ['sg'], 14.59390294, false)
add('mass', ['lbm'], 0.45359237, false)
add('mass', ['u'], 1.66053906660e-27)
add('mass', ['ozm'], 0.028349523125, false)
add('mass', ['grain'], 6.479891e-5, false)
add('mass', ['cwt', 'shweight'], 45.359237, false)
add('mass', ['uk_cwt', 'lcwt', 'hweight'], 50.80234544, false)
add('mass', ['stone'], 6.35029318, false)
add('mass', ['ton'], 907.18474, false)
add('mass', ['uk_ton', 'LTON', 'brton'], 1016.0469088, false)
// Distance (m)
add('dist', ['m'], 1)
add('dist', ['mi'], 1609.344, false)
add('dist', ['Nmi'], 1852, false)
add('dist', ['in'], 0.0254, false)
add('dist', ['ft'], 0.3048, false)
add('dist', ['yd'], 0.9144, false)
add('dist', ['ang'], 1e-10)
add('dist', ['ell'], 1.143, false)
add('dist', ['ly'], 9460730472580800)
add('dist', ['parsec', 'pc'], 3.08567758149137e16)
add('dist', ['Picapt', 'Pica'], 0.0254 / 72, false)
add('dist', ['pica'], 0.0254 / 6, false)
add('dist', ['survey_mi'], 1609.3472186944, false)
// Time (s)
add('time', ['yr'], 365.25 * 86400)
add('time', ['day', 'd'], 86400)
add('time', ['hr'], 3600)
add('time', ['mn', 'min'], 60)
add('time', ['sec', 's'], 1)
// Pressure (Pa)
add('press', ['Pa', 'p'], 1)
add('press', ['atm', 'at'], 101325)
add('press', ['mmHg'], 133.322387415)
add('press', ['psi'], 6894.75729316836, false)
add('press', ['Torr'], 101325 / 760)
// Force (N)
add('force', ['N'], 1)
add('force', ['dyn', 'dy'], 1e-5)
add('force', ['lbf'], 4.4482216152605, false)
add('force', ['pond'], 9.80665e-3)
// Energy (J)
add('energy', ['J'], 1)
add('energy', ['e'], 1e-7)
add('energy', ['c'], 4.184)
add('energy', ['cal'], 4.1868)
add('energy', ['eV', 'ev'], 1.602176634e-19)
add('energy', ['HPh', 'hh'], 2684519.537696, false)
add('energy', ['Wh', 'wh'], 3600)
add('energy', ['flb'], 1.3558179483314, false)
add('energy', ['BTU', 'btu'], 1055.05585262, false)
// Power (W)
add('power', ['HP', 'h'], 745.69987158227, false)
add('power', ['PS'], 735.49875, false)
add('power', ['W', 'w'], 1)
// Magnetism (T)
add('mag', ['T'], 1)
add('mag', ['ga'], 1e-4)
// Volume (m³)
add('vol', ['tsp'], 4.92892159375e-6, false)
add('vol', ['tspm'], 5e-6, false)
add('vol', ['tbs'], 1.478676478125e-5, false)
add('vol', ['oz'], 2.95735295625e-5, false)
add('vol', ['cup'], 2.365882365e-4, false)
add('vol', ['pt', 'us_pt'], 4.73176473e-4, false)
add('vol', ['uk_pt'], 5.6826125e-4, false)
add('vol', ['qt'], 9.46352946e-4, false)
add('vol', ['uk_qt'], 1.1365225e-3, false)
add('vol', ['gal'], 3.785411784e-3, false)
add('vol', ['uk_gal'], 4.54609e-3, false)
add('vol', ['l', 'L', 'lt'], 1e-3)
add('vol', ['ang3', 'ang^3'], 1e-30)
add('vol', ['barrel'], 0.158987294928, false)
add('vol', ['bushel'], 0.03523907016688, false)
add('vol', ['ft3', 'ft^3'], 0.028316846592, false)
add('vol', ['in3', 'in^3'], 1.6387064e-5, false)
add('vol', ['ly3', 'ly^3'], 9460730472580800 ** 3)
add('vol', ['m3', 'm^3'], 1)
add('vol', ['mi3', 'mi^3'], 1609.344 ** 3, false)
add('vol', ['yd3', 'yd^3'], 0.9144 ** 3, false)
add('vol', ['Nmi3', 'Nmi^3'], 1852 ** 3, false)
add('vol', ['Picapt3', 'Picapt^3', 'Pica3', 'Pica^3'], (0.0254 / 72) ** 3, false)
add('vol', ['GRT', 'regton'], 2.8316846592, false)
add('vol', ['MTON'], 1.13267386368, false)
// Area (m²)
add('area', ['uk_acre'], 4046.8564224, false)
add('area', ['us_acre'], 4046.87260987425, false)
add('area', ['ang2', 'ang^2'], 1e-20)
add('area', ['ar'], 100)
add('area', ['ft2', 'ft^2'], 0.09290304, false)
add('area', ['ha'], 10000)
add('area', ['in2', 'in^2'], 6.4516e-4, false)
add('area', ['ly2', 'ly^2'], 9460730472580800 ** 2)
add('area', ['m2', 'm^2'], 1)
add('area', ['Morgen'], 2500, false)
add('area', ['mi2', 'mi^2'], 1609.344 ** 2, false)
add('area', ['Nmi2', 'Nmi^2'], 1852 ** 2, false)
add('area', ['Picapt2', 'Pica2', 'Pica^2', 'Picapt^2'], (0.0254 / 72) ** 2, false)
add('area', ['yd2', 'yd^2'], 0.83612736, false)
// Information (bit)
add('info', ['bit'], 1, true, true)
add('info', ['byte'], 8, true, true)
// Speed (m/s)
add('speed', ['admkn'], 0.514773333333333, false)
add('speed', ['kn'], 0.514444444444444, false)
add('speed', ['m/h', 'm/hr'], 1 / 3600)
add('speed', ['m/s', 'm/sec'], 1)
add('speed', ['mph'], 0.44704, false)

const PREFIX: Record<string, number> = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 1e1, e: 1e1, d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24 }
const BIN_PREFIX: Record<string, number> = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 }
const TEMPS = new Set(['C', 'cel', 'F', 'fah', 'K', 'kel', 'Rank', 'Reau'])

function unit(name: string): { g: string; f: number; temp?: string } | null {
  if (TEMPS.has(name)) return { g: 'temp', f: 1, temp: name }
  if (U[name]) return U[name]
  for (const [p, f] of Object.entries(BIN_PREFIX)) if (name.startsWith(p) && U[name.slice(p.length)]?.binary) return { g: U[name.slice(p.length)].g, f: U[name.slice(p.length)].f * f }
  for (const [p, f] of Object.entries(PREFIX).sort((a, b) => b[0].length - a[0].length)) {
    const rest = name.slice(p.length)
    if (name.startsWith(p) && U[rest]?.prefix) {
      // Squared/cubed units take the prefix to that power (km2 = 1e6 m²)
      const pow = /[2^]2$/.test(rest) || /\^2$/.test(rest) ? 2 : /[3^]3$/.test(rest) || /\^3$/.test(rest) ? 3 : 1
      return { g: U[rest].g, f: U[rest].f * Math.pow(f, pow) }
    }
    if (name.startsWith(p) && (rest === 'K' || rest === 'kel')) return { g: 'temp', f, temp: 'K' }
  }
  return null
}

function toKelvin(v: number, t: string) {
  if (t === 'C' || t === 'cel') return v + 273.15
  if (t === 'F' || t === 'fah') return ((v - 32) * 5) / 9 + 273.15
  if (t === 'Rank') return (v * 5) / 9
  if (t === 'Reau') return v * 1.25 + 273.15
  return v
}
function fromKelvin(k: number, t: string) {
  if (t === 'C' || t === 'cel') return k - 273.15
  if (t === 'F' || t === 'fah') return ((k - 273.15) * 9) / 5 + 32
  if (t === 'Rank') return (k * 9) / 5
  if (t === 'Reau') return (k - 273.15) * 0.8
  return k
}

defineFunctions([
  fn('BESSELJ', 2, 2, 'n', ([x, n]) => (Math.trunc(n) < 0 ? ERR.NUM : checkNum(jn(Math.trunc(n), x)))),
  fn('BESSELY', 2, 2, 'n', ([x, n]) => (Math.trunc(n) < 0 || x <= 0 ? ERR.NUM : checkNum(yn(Math.trunc(n), x)))),
  fn('BESSELI', 2, 2, 'n', ([x, n]) => (Math.trunc(n) < 0 ? ERR.NUM : checkNum(inBessel(Math.trunc(n), x)))),
  fn('BESSELK', 2, 2, 'n', ([x, n]) => (Math.trunc(n) < 0 || x <= 0 ? ERR.NUM : checkNum(kn(Math.trunc(n), x)))),
  fn('BIN2DEC', 1, 1, 'v', convertBase(2, 10)),
  fn('BIN2HEX', 1, 2, 'vv', convertBase(2, 16)),
  fn('BIN2OCT', 1, 2, 'vv', convertBase(2, 8)),
  fn('HEX2BIN', 1, 2, 'vv', convertBase(16, 2)),
  fn('HEX2DEC', 1, 1, 'v', convertBase(16, 10)),
  fn('HEX2OCT', 1, 2, 'vv', convertBase(16, 8)),
  fn('OCT2BIN', 1, 2, 'vv', convertBase(8, 2)),
  fn('OCT2DEC', 1, 1, 'v', convertBase(8, 10)),
  fn('OCT2HEX', 1, 2, 'vv', convertBase(8, 16)),
  fn('DEC2BIN', 1, 2, 'nn', ([n, p]) => toBase(n, 2, p)),
  fn('DEC2HEX', 1, 2, 'nn', ([n, p]) => toBase(n, 16, p)),
  fn('DEC2OCT', 1, 2, 'nn', ([n, p]) => toBase(n, 8, p)),
  ...(['BITAND', 'BITOR', 'BITXOR'] as const).map((name) =>
    fn(name, 2, 2, 'n', ([a, b]) => {
      const x = bitArg(a)
      const y = bitArg(b)
      if (isError(x)) return x
      if (isError(y)) return y
      return Number(name === 'BITAND' ? x & y : name === 'BITOR' ? x | y : x ^ y)
    })
  ),
  ...(['BITLSHIFT', 'BITRSHIFT'] as const).map((name) =>
    fn(name, 2, 2, 'n', ([a, s]) => {
      const x = bitArg(a)
      if (isError(x)) return x
      s = Math.trunc(s)
      if (Math.abs(s) > 53) return ERR.NUM
      const left = name === 'BITLSHIFT' ? s : -s
      const r = left >= 0 ? x << BigInt(left) : x >> BigInt(-left)
      return r >= BigInt(MAX48) ? ERR.NUM : Number(r)
    })
  ),
  fn('DELTA', 1, 2, 'n', ([a, b = 0]) => (a === b ? 1 : 0)),
  fn('GESTEP', 1, 2, 'n', ([a, s = 0]) => (a >= s ? 1 : 0)),
  fn('ERF', 1, 2, 'n', ([lo, hi]) => (hi === undefined ? erf(lo) : erf(hi) - erf(lo))),
  fn('ERF.PRECISE', 1, 1, 'n', ([x]) => erf(x)),
  fn('ERFC', 1, 1, 'n', ([x]) => erfc(x)),
  fn('ERFC.PRECISE', 1, 1, 'n', ([x]) => erfc(x)),
  fn('COMPLEX', 2, 3, 'nns', ([re, im, sfx = 'i']) => (sfx !== 'i' && sfx !== 'j' && sfx !== '' ? ERR.VALUE : fmtCx(cx(re, im, (sfx || 'i') as 'i' | 'j')))),
  fn('IMREAL', 1, 1, 'v', ([v]) => {
    const z = parseCx(v)
    return isError(z) ? z : z.re
  }),
  fn('IMAGINARY', 1, 1, 'v', ([v]) => {
    const z = parseCx(v)
    return isError(z) ? z : z.im
  }),
  fn('IMABS', 1, 1, 'v', ([v]) => {
    const z = parseCx(v)
    return isError(z) ? z : cabs(z)
  }),
  fn('IMARGUMENT', 1, 1, 'v', ([v]) => {
    const z = parseCx(v)
    if (isError(z)) return z
    return z.re === 0 && z.im === 0 ? ERR.DIV0 : Math.atan2(z.im, z.re)
  }),
  cx1('IMCONJUGATE', (z) => cx(z.re, -z.im, z.sfx)),
  cx1('IMEXP', cexp),
  cx1('IMLN', cln),
  cx1('IMLOG10', (z) => {
    const l = cln(z)
    return isError(l) ? l : cx(l.re / Math.LN10, l.im / Math.LN10, z.sfx)
  }),
  cx1('IMLOG2', (z) => {
    const l = cln(z)
    return isError(l) ? l : cx(l.re / Math.LN2, l.im / Math.LN2, z.sfx)
  }),
  cx1('IMSIN', csin),
  cx1('IMCOS', ccos),
  cx1('IMTAN', (z) => cdiv(csin(z), ccos(z))),
  cx1('IMCOT', (z) => cdiv(ccos(z), csin(z))),
  cx1('IMSEC', (z) => cdiv(one(z), ccos(z))),
  cx1('IMCSC', (z) => cdiv(one(z), csin(z))),
  cx1('IMSINH', csinh),
  cx1('IMCOSH', ccosh),
  cx1('IMSECH', (z) => cdiv(one(z), ccosh(z))),
  cx1('IMCSCH', (z) => cdiv(one(z), csinh(z))),
  cx1('IMSQRT', (z) => {
    const r = Math.sqrt(cabs(z))
    const t = Math.atan2(z.im, z.re) / 2
    return cx(r * Math.cos(t), r * Math.sin(t), z.sfx)
  }),
  fn('IMPOWER', 2, 2, 'vn', ([v, n]) => {
    const z = parseCx(v)
    if (isError(z)) return z
    if (z.re === 0 && z.im === 0) return n > 0 ? '0' : ERR.NUM
    const r = Math.pow(cabs(z), n)
    const t = Math.atan2(z.im, z.re) * n
    return fmtCx(cx(r * Math.cos(t), r * Math.sin(t), z.sfx))
  }),
  fn('IMDIV', 2, 2, 'v', ([a, b]) => {
    const x = parseCx(a)
    const y = parseCx(b)
    if (isError(x)) return x
    if (isError(y)) return y
    const r = cdiv(x, y)
    return isError(r) ? r : fmtCx({ ...r, sfx: x.im ? x.sfx : y.sfx })
  }),
  fn('IMSUB', 2, 2, 'v', ([a, b]) => {
    const x = parseCx(a)
    const y = parseCx(b)
    if (isError(x)) return x
    if (isError(y)) return y
    return fmtCx(cx(x.re - y.re, x.im - y.im, x.im ? x.sfx : y.sfx))
  }),
  fn('IMSUM', 1, MANY, 'r', (args, ctx) => {
    const zs = cxList(args, ctx)
    if (isError(zs)) return zs
    const sfx = zs.find((z) => z.im)?.sfx ?? 'i'
    return fmtCx(zs.reduce((a, z) => cx(a.re + z.re, a.im + z.im, sfx), cx(0, 0, sfx)))
  }),
  fn('IMPRODUCT', 1, MANY, 'r', (args, ctx) => {
    const zs = cxList(args, ctx)
    if (isError(zs)) return zs
    const sfx = zs.find((z) => z.im)?.sfx ?? 'i'
    return fmtCx({ ...zs.reduce((a, z) => cmul(a, z), cx(1, 0, sfx)), sfx })
  }),
  fn('CONVERT', 3, 3, 'nss', ([v, from, to]) => {
    const a = unit(from)
    const b = unit(to)
    if (!a || !b || a.g !== b.g) return ERR.NA
    if (a.g === 'temp') return fromKelvin(toKelvin(v * (a.temp === 'K' ? a.f : 1), a.temp!), b.temp!) / (b.temp === 'K' ? b.f : 1)
    return checkNum((v * a.f) / b.f)
  })
])
