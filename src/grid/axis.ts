// Row or column positions: a default size plus sparse overrides (hidden = size 0). Offsets in unzoomed pixels.

export class Axis {
  private keys: number[] = []
  /** prefix[j] = sum of (size - default) over keys[0..j-1] */
  private prefix: number[] = [0]
  private sizes = new Map<number, number>()

  constructor(
    readonly def: number,
    overrides: [number, number][],
    readonly count: number
  ) {
    const sorted = [...overrides].sort((a, b) => a[0] - b[0])
    for (const [i, s] of sorted) {
      if (s === def) continue
      this.keys.push(i)
      this.sizes.set(i, s)
      this.prefix.push(this.prefix[this.prefix.length - 1] + (s - def))
    }
  }

  size(i: number): number {
    return this.sizes.get(i) ?? this.def
  }

  /** Number of override keys below i. */
  private below(i: number): number {
    let lo = 0
    let hi = this.keys.length
    while (lo < hi) {
      const m = (lo + hi) >> 1
      if (this.keys[m] < i) lo = m + 1
      else hi = m
    }
    return lo
  }

  offset(i: number): number {
    return i * this.def + this.prefix[this.below(i)]
  }

  /** Index whose span contains px (clamped). Hidden indexes are skipped. */
  indexAt(px: number): number {
    if (px <= 0) return 0
    let lo = 0
    let hi = this.count - 1
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1
      if (this.offset(m) <= px) lo = m
      else hi = m - 1
    }
    while (lo < this.count - 1 && this.size(lo) === 0) lo++
    return lo
  }

  /** Next visible index in a direction (skips hidden rows/columns). */
  visible(i: number, step: 1 | -1): number {
    let j = i
    while (j >= 0 && j < this.count && this.size(j) === 0) j += step
    return Math.max(0, Math.min(this.count - 1, j))
  }
}
