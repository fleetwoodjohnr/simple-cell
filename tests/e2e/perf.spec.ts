// Opt-in (SC_PERF=1): grid scrolling frame times in the real app.
import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook } from './helpers'

test.skip(!process.env.SC_PERF, 'perf run only')
test('scrolling a filled sheet: frame times', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    // 20,000 rows × 8 columns of numbers and text through the core.
    await page.evaluate(async () => {
      const sc = (window as any).__sc
      const tab = sc.get().active
      const book = sc.books.get(tab)
      const s = sc.getView(tab).sheet
      const rows = Array.from({ length: 20_000 }, (_, r) => Array.from({ length: 8 }, (_, c) => (c % 2 ? `Row ${r} col ${c}` : r * 8 + c)))
      await book.api.setRangeValues(s, { sheet: s, r0: 0, c0: 0, r1: 19_999, c1: 7 }, rows)
    })
    await page.waitForTimeout(500)
    const box = (await page.getByTestId('grid').boundingBox())!
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
    const frames = await page.evaluate(
      () =>
        new Promise<number[]>((resolve) => {
          const out: number[] = []
          let last = performance.now()
          const tick = (t: number) => {
            out.push(t - last)
            last = t
            if (out.length < 240) requestAnimationFrame(tick)
            else resolve(out.slice(10))
          }
          requestAnimationFrame(tick)
        })
    ).catch(() => [] as number[])
    void frames
    // Measure while wheel-scrolling.
    const measuring = page.evaluate(
      () =>
        new Promise<number[]>((resolve) => {
          const out: number[] = []
          let last = performance.now()
          const tick = (t: number) => {
            out.push(t - last)
            last = t
            if (out.length < 180) requestAnimationFrame(tick)
            else resolve(out.slice(5))
          }
          requestAnimationFrame(tick)
        })
    )
    for (let i = 0; i < 120; i++) await page.mouse.wheel(0, 120)
    const times = (await measuring).sort((a, b) => a - b)
    const p = (q: number) => times[Math.floor(times.length * q)]
    console.log(`scroll frames: median ${p(0.5).toFixed(1)} ms, p95 ${p(0.95).toFixed(1)} ms, max ${times[times.length - 1].toFixed(1)} ms`)
    expect(times.length).toBeGreaterThan(50)
  } finally {
    await quit(app)
  }
})
