// Conditional formatting icon sets, drawn as small vector shapes (index 0 = lowest value).
const GREEN = '#2e9e4f'
const YELLOW = '#e8b100'
const RED = '#d9372b'
const GRAY = '#8c8c8c'
const BLACK = '#3a3a3a'
const PINK = '#f2a6a6'

type Shape = { k: 'arrow'; dir: number; color: string } | { k: 'dot'; color: string; ring?: boolean } | { k: 'flag'; color: string } | { k: 'symbol'; sym: 'check' | 'bang' | 'cross'; color: string; plain?: boolean } | { k: 'bars'; n: number; of: number } | { k: 'pie'; f: number } | { k: 'star'; f: number } | { k: 'tri'; dir: 1 | 0 | -1; color: string } | { k: 'boxes'; n: number } | { k: 'sign'; i: number }

const arrows = (n: number, gray: boolean): Shape[] => {
  const dirs = n === 3 ? [-90, 0, 90] : n === 4 ? [-90, -45, 45, 90] : [-90, -45, 0, 45, 90]
  const colors = gray ? dirs.map(() => GRAY) : n === 3 ? [RED, YELLOW, GREEN] : n === 4 ? [RED, YELLOW, YELLOW, GREEN] : [RED, YELLOW, YELLOW, YELLOW, GREEN]
  return dirs.map((dir, i) => ({ k: 'arrow', dir, color: colors[i] }))
}

export function iconShape(set: string, i: number): Shape | null {
  const list: Shape[] | undefined = (
    {
      '3Arrows': arrows(3, false),
      '3ArrowsGray': arrows(3, true),
      '4Arrows': arrows(4, false),
      '4ArrowsGray': arrows(4, true),
      '5Arrows': arrows(5, false),
      '5ArrowsGray': arrows(5, true),
      '3Flags': [RED, YELLOW, GREEN].map((color) => ({ k: 'flag', color })),
      '3TrafficLights1': [RED, YELLOW, GREEN].map((color) => ({ k: 'dot', color })),
      '3TrafficLights2': [RED, YELLOW, GREEN].map((color) => ({ k: 'dot', color, ring: true })),
      '4TrafficLights': [BLACK, RED, YELLOW, GREEN].map((color) => ({ k: 'dot', color })),
      '4RedToBlack': [BLACK, GRAY, PINK, RED].map((color) => ({ k: 'dot', color })),
      '3Signs': [0, 1, 2].map((i) => ({ k: 'sign', i })),
      '3Symbols': [
        { k: 'symbol', sym: 'cross', color: RED },
        { k: 'symbol', sym: 'bang', color: YELLOW },
        { k: 'symbol', sym: 'check', color: GREEN }
      ],
      '3Symbols2': [
        { k: 'symbol', sym: 'cross', color: RED, plain: true },
        { k: 'symbol', sym: 'bang', color: YELLOW, plain: true },
        { k: 'symbol', sym: 'check', color: GREEN, plain: true }
      ],
      '3Stars': [0, 0.5, 1].map((f) => ({ k: 'star', f })),
      '3Triangles': [
        { k: 'tri', dir: -1, color: RED },
        { k: 'tri', dir: 0, color: YELLOW },
        { k: 'tri', dir: 1, color: GREEN }
      ],
      '4Rating': [1, 2, 3, 4].map((n) => ({ k: 'bars', n, of: 4 })),
      '5Rating': [0, 1, 2, 3, 4].map((n) => ({ k: 'bars', n, of: 4 })),
      '5Quarters': [0, 0.25, 0.5, 0.75, 1].map((f) => ({ k: 'pie', f })),
      '5Boxes': [0, 1, 2, 3, 4].map((n) => ({ k: 'boxes', n }))
    } as Record<string, Shape[]>
  )[set]
  return list?.[Math.max(0, Math.min(list.length - 1, i))] ?? null
}

/** Draw icon `i` of `set` in a square of side `s` at (x, y). */
export function drawIcon(ctx: CanvasRenderingContext2D, set: string, i: number, x: number, y: number, s: number) {
  const shape = iconShape(set, i)
  if (!shape) return
  const cx = x + s / 2
  const cy = y + s / 2
  const r = s * 0.42
  ctx.save()
  ctx.lineJoin = 'round'
  ctx.lineCap = 'round'
  switch (shape.k) {
    case 'arrow': {
      ctx.translate(cx, cy)
      ctx.rotate((-shape.dir * Math.PI) / 180)
      ctx.fillStyle = shape.color
      ctx.beginPath()
      ctx.moveTo(r, 0)
      ctx.lineTo(0, -r * 0.85)
      ctx.lineTo(0, -r * 0.35)
      ctx.lineTo(-r, -r * 0.35)
      ctx.lineTo(-r, r * 0.35)
      ctx.lineTo(0, r * 0.35)
      ctx.lineTo(0, r * 0.85)
      ctx.closePath()
      ctx.fill()
      break
    }
    case 'dot':
      ctx.fillStyle = shape.color
      ctx.beginPath()
      ctx.arc(cx, cy, r, 0, Math.PI * 2)
      ctx.fill()
      if (shape.ring) {
        ctx.strokeStyle = BLACK
        ctx.lineWidth = Math.max(1, s * 0.08)
        ctx.stroke()
      }
      break
    case 'flag':
      ctx.strokeStyle = BLACK
      ctx.lineWidth = Math.max(1, s * 0.08)
      ctx.beginPath()
      ctx.moveTo(x + s * 0.25, y + s * 0.9)
      ctx.lineTo(x + s * 0.25, y + s * 0.1)
      ctx.stroke()
      ctx.fillStyle = shape.color
      ctx.beginPath()
      ctx.moveTo(x + s * 0.25, y + s * 0.1)
      ctx.lineTo(x + s * 0.9, y + s * 0.3)
      ctx.lineTo(x + s * 0.25, y + s * 0.52)
      ctx.closePath()
      ctx.fill()
      break
    case 'symbol': {
      if (!shape.plain) {
        ctx.fillStyle = shape.color
        ctx.beginPath()
        ctx.arc(cx, cy, r, 0, Math.PI * 2)
        ctx.fill()
      }
      ctx.strokeStyle = shape.plain ? shape.color : '#fff'
      ctx.lineWidth = Math.max(1.5, s * (shape.plain ? 0.18 : 0.13))
      ctx.beginPath()
      const q = r * 0.5
      if (shape.sym === 'check') {
        ctx.moveTo(cx - q, cy)
        ctx.lineTo(cx - q * 0.2, cy + q * 0.75)
        ctx.lineTo(cx + q, cy - q * 0.7)
      } else if (shape.sym === 'cross') {
        ctx.moveTo(cx - q, cy - q)
        ctx.lineTo(cx + q, cy + q)
        ctx.moveTo(cx + q, cy - q)
        ctx.lineTo(cx - q, cy + q)
      } else {
        ctx.moveTo(cx, cy - q * 1.1)
        ctx.lineTo(cx, cy + q * 0.35)
        ctx.moveTo(cx, cy + q * 0.95)
        ctx.lineTo(cx, cy + q * 1.0)
      }
      ctx.stroke()
      break
    }
    case 'sign': {
      // Red diamond, yellow triangle, green circle.
      ctx.fillStyle = [RED, YELLOW, GREEN][shape.i]
      ctx.beginPath()
      if (shape.i === 0) {
        ctx.moveTo(cx, cy - r)
        ctx.lineTo(cx + r, cy)
        ctx.lineTo(cx, cy + r)
        ctx.lineTo(cx - r, cy)
      } else if (shape.i === 1) {
        ctx.moveTo(cx, cy - r)
        ctx.lineTo(cx + r, cy + r * 0.8)
        ctx.lineTo(cx - r, cy + r * 0.8)
      } else ctx.arc(cx, cy, r, 0, Math.PI * 2)
      ctx.closePath()
      ctx.fill()
      break
    }
    case 'tri':
      ctx.fillStyle = shape.color
      ctx.beginPath()
      if (shape.dir === 0) ctx.fillRect(x + s * 0.15, cy - s * 0.1, s * 0.7, s * 0.2)
      else {
        const d = shape.dir
        ctx.moveTo(cx, cy - d * r * 0.8)
        ctx.lineTo(cx + r, cy + d * r * 0.6)
        ctx.lineTo(cx - r, cy + d * r * 0.6)
        ctx.closePath()
        ctx.fill()
      }
      break
    case 'bars': {
      const w = s / 5.5
      for (let k = 0; k < shape.of; k++) {
        const h = (s * 0.85 * (k + 1)) / shape.of
        ctx.fillStyle = k < shape.n ? '#2f6fbf' : '#c8d3e3'
        ctx.fillRect(x + s * 0.08 + k * w * 1.3, y + s * 0.92 - h, w, h)
      }
      break
    }
    case 'pie':
      ctx.strokeStyle = BLACK
      ctx.lineWidth = Math.max(1, s * 0.07)
      ctx.beginPath()
      ctx.arc(cx, cy, r, 0, Math.PI * 2)
      ctx.stroke()
      if (shape.f > 0) {
        ctx.fillStyle = BLACK
        ctx.beginPath()
        ctx.moveTo(cx, cy)
        ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + shape.f * Math.PI * 2)
        ctx.closePath()
        ctx.fill()
      }
      break
    case 'star': {
      const path = () => {
        ctx.beginPath()
        for (let k = 0; k < 10; k++) {
          const a = -Math.PI / 2 + (k * Math.PI) / 5
          const rr = k % 2 ? r * 0.45 : r
          ctx.lineTo(cx + rr * Math.cos(a), cy + rr * Math.sin(a))
        }
        ctx.closePath()
      }
      path()
      ctx.strokeStyle = YELLOW
      ctx.lineWidth = Math.max(1, s * 0.07)
      ctx.stroke()
      if (shape.f > 0) {
        ctx.save()
        ctx.beginPath()
        ctx.rect(cx - r, cy - r, 2 * r * shape.f, 2 * r)
        ctx.clip()
        path()
        ctx.fillStyle = YELLOW
        ctx.fill()
        ctx.restore()
      }
      break
    }
    case 'boxes': {
      const g = s * 0.4
      const pos = [
        [0, 1],
        [1, 1],
        [0, 0],
        [1, 0]
      ]
      pos.forEach(([px, py], k) => {
        ctx.fillStyle = k < shape.n ? '#2f6fbf' : '#c8d3e3'
        ctx.fillRect(x + s * 0.08 + px * (g + s * 0.04), y + s * 0.08 + py * (g + s * 0.04), g, g)
      })
      break
    }
  }
  ctx.restore()
}
