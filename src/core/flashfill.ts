// Flash Fill: learn a text transformation from examples and apply it to the other rows.
// ponytail: a small token-based learner (substrings by word position, initials, case changes and literal text) —
// covers splitting names, joining columns, reformatting codes and initials; not Excel's full PROSE engine.

type Case = 'same' | 'upper' | 'lower' | 'proper'
type Seg = { k: 'lit'; s: string } | { k: 'cut'; src: number; side: 'before' | 'after'; delim: string; nth: number; fromEnd: boolean; case: Case } | { k: 'tok'; src: number; from: number; to: number; fromEnd: boolean; case: Case } | { k: 'init'; src: number; tok: number; fromEnd: boolean; case: Case }
export type Program = Seg[]

/** Words (letters), numbers and single other characters; spaces are kept as tokens too. */
export function tokens(s: string): string[] {
  return s.match(/[\p{L}\p{M}']+|\d+|\s+|[^\p{L}\p{M}\d\s]/gu) ?? []
}

const words = (s: string) => tokens(s).filter((t) => /[\p{L}\d]/u.test(t))

function cased(s: string, c: Case): string {
  if (c === 'upper') return s.toUpperCase()
  if (c === 'lower') return s.toLowerCase()
  if (c === 'proper') return s.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase())
  return s
}

function evalSeg(seg: Seg, inputs: string[]): string | null {
  if (seg.k === 'lit') return seg.s
  const src = inputs[seg.src]
  if (src === undefined) return null
  if (seg.k === 'cut') {
    // The text before/after the nth occurrence of a delimiter character.
    const at: number[] = []
    for (let i = 0; i < src.length; i++) if (src[i] === seg.delim) at.push(i)
    const i = seg.fromEnd ? at[at.length - 1 - seg.nth] : at[seg.nth]
    if (i === undefined) return null
    return cased(seg.side === 'before' ? src.slice(0, i) : src.slice(i + 1), seg.case)
  }
  if (seg.k === 'init') {
    const w = words(src)
    const t = seg.fromEnd ? w[w.length - 1 - seg.tok] : w[seg.tok]
    return t ? cased(t[0], seg.case) : null
  }
  const t = tokens(src)
  const from = seg.fromEnd ? t.length - 1 - seg.from : seg.from
  const to = seg.fromEnd ? t.length - 1 - seg.to : seg.to
  if (from < 0 || to >= t.length || from > to) return null
  return cased(t.slice(from, to + 1).join(''), seg.case)
}

export function run(p: Program, inputs: string[]): string | null {
  let out = ''
  for (const seg of p) {
    const v = evalSeg(seg, inputs)
    if (v === null) return null
    out += v
  }
  return out
}

const CASES: Case[] = ['same', 'upper', 'lower', 'proper']
// Simpler programs win; letters and digits as literal text are suspect (they usually come from an input); a piece anchored at the nearer end of its input (first or last word) beats the other.
const cost = (p: Program) =>
  p.reduce((n, s) => n + (s.k === 'lit' ? 2 + s.s.length * 0.5 + (s.s.match(/[\p{L}\d]/gu)?.length ?? 0) * 3 : s.k === 'init' ? 1.5 + s.tok * 0.05 : s.k === 'cut' ? 1.05 + s.nth * 0.05 + (s.case !== 'same' ? 0.2 : 0) : 1 + Math.abs(s.to - s.from) * 0.1 + (s.case !== 'same' ? 0.2 : 0) + Math.min(s.from, s.to) * 0.05 + (s.fromEnd ? 0.01 : 0)), 0)

/** Programs (up to a limit) that turn `inputs` into `output`. */
function candidates(inputs: string[], output: string, limit = 2000): Program[] {
  // Every way to produce output[p..q) from one piece.
  const pieces: { at: number; len: number; seg: Seg }[] = []
  inputs.forEach((src, si) => {
    const t = tokens(src)
    for (let i = 0; i < t.length; i++) {
      let text = ''
      for (let j = i; j < t.length; j++) {
        text += t[j]
        if (!/[\p{L}\d]/u.test(t[i]) || !/[\p{L}\d]/u.test(t[j])) continue
        for (const c of CASES) {
          const v = cased(text, c)
          for (let at = output.indexOf(v); at >= 0; at = output.indexOf(v, at + 1)) {
            pieces.push({ at, len: v.length, seg: { k: 'tok', src: si, from: i, to: j, fromEnd: false, case: c } })
            pieces.push({ at, len: v.length, seg: { k: 'tok', src: si, from: t.length - 1 - i, to: t.length - 1 - j, fromEnd: true, case: c } })
          }
        }
      }
    }
    // Before / after a delimiter ("everything before the @").
    const seen = new Map<string, number[]>()
    for (let i = 0; i < src.length; i++) if (/[^\p{L}\p{M}\d\s]/u.test(src[i])) seen.set(src[i], [...(seen.get(src[i]) ?? []), i])
    for (const [delim, at] of seen)
      at.forEach((i, n) => {
        for (const side of ['before', 'after'] as const) {
          const part = side === 'before' ? src.slice(0, i) : src.slice(i + 1)
          if (!part) continue
          for (const c of CASES) {
            const v = cased(part, c)
            for (let pos = output.indexOf(v); pos >= 0; pos = output.indexOf(v, pos + 1)) {
              pieces.push({ at: pos, len: v.length, seg: { k: 'cut', src: si, side, delim, nth: n, fromEnd: false, case: c } })
              pieces.push({ at: pos, len: v.length, seg: { k: 'cut', src: si, side, delim, nth: at.length - 1 - n, fromEnd: true, case: c } })
            }
          }
        }
      })
    const w = words(src)
    w.forEach((word, wi) => {
      for (const c of CASES) {
        const v = cased(word[0], c)
        for (let at = output.indexOf(v); at >= 0; at = output.indexOf(v, at + 1)) {
          pieces.push({ at, len: 1, seg: { k: 'init', src: si, tok: wi, fromEnd: false, case: c } })
          pieces.push({ at, len: 1, seg: { k: 'init', src: si, tok: w.length - 1 - wi, fromEnd: true, case: c } })
        }
      }
    })
  })
  const byStart = new Map<number, typeof pieces>()
  for (const p of pieces) byStart.set(p.at, [...(byStart.get(p.at) ?? []), p])
  // Paths through the output, literals filling gaps; shortest programs first.
  const memo = new Map<number, Program[]>()
  const from = (pos: number): Program[] => {
    if (pos === output.length) return [[]]
    const hit = memo.get(pos)
    if (hit) return hit
    const out: Program[] = []
    for (const p of byStart.get(pos) ?? []) for (const rest of from(pos + p.len)) if (out.length < limit) out.push([p.seg, ...rest])
    // A literal up to the next place a piece can start (or the end).
    let next = pos + 1
    while (next < output.length && !byStart.has(next)) next++
    for (const rest of from(next)) {
      if (out.length >= limit) break
      const lit = output.slice(pos, next)
      out.push(rest[0]?.k === 'lit' ? [{ k: 'lit', s: lit + rest[0].s }, ...rest.slice(1)] : [{ k: 'lit', s: lit }, ...rest])
    }
    out.sort((a, b) => cost(a) - cost(b))
    const kept = out.slice(0, limit)
    memo.set(pos, kept)
    return kept
  }
  return from(0)
}

/** The simplest program consistent with every example, or null. */
export function learn(examples: { inputs: string[]; output: string }[]): Program | null {
  if (!examples.length) return null
  const [first, ...rest] = examples
  const progs = candidates(first.inputs, first.output).filter((p) => rest.every((e) => run(p, e.inputs) === e.output))
  // A program that is all literal text would just repeat the example: not a pattern.
  const real = progs.filter((p) => p.some((s) => s.k !== 'lit'))
  real.sort((a, b) => cost(a) - cost(b))
  return real[0] ?? null
}
