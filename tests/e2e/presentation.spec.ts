import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook, goto } from './helpers'

test('border pictograms and persistent arrow/Excel cursor preference', async () => {
  let { app, page, userData } = await launch()
  try {
    await newWorkbook(page)
    const grid = page.getByTestId('grid-scroll')
    await grid.hover({ position: { x: 140, y: 80 } })
    await expect(grid).toHaveCSS('cursor', 'default')
    await page.getByRole('button', { name: 'Borders', exact: true }).click()
    for (const name of ['Bottom border', 'All borders', 'Outside borders', 'Thick outside borders', 'Bottom double border', 'Top and thick bottom border']) {
      await expect(page.getByRole('menuitem', { name: new RegExp(`^${name}(?:\\s|$)`) }).locator('svg')).toHaveCount(1)
    }
    await page.getByRole('menuitem', { name: 'All borders', exact: true }).click()
    await goto(page, 'A1')
    await page.evaluate(() => (window as any).__sc.runCommand('tools.settings'))
    await page.getByRole('radio', { name: 'Excel cross', exact: true }).click()
    await page.getByRole('button', { name: 'Done', exact: true }).click()
    await grid.hover({ position: { x: 140, y: 80 } })
    await expect.poll(() => grid.evaluate((el) => el.style.cursor)).toContain('data:image/svg+xml')
    await quit(app)
    ;({ app, page } = await launch({ userData }))
    await page.evaluate(() => (window as any).__sc.runCommand('tools.settings'))
    await expect(page.getByRole('radio', { name: 'Excel cross', exact: true })).toBeChecked()
    await page.getByRole('radio', { name: 'Arrow', exact: true }).click()
    await page.getByRole('button', { name: 'Done', exact: true }).click()
  } finally { await quit(app) }
})

test('bold uses each font’s real bold face, installed or bundled (not a faked bold of the regular)', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    const same = await page.evaluate(async () => {
      const text = 'Quarterly revenue by region'
      const families = [...new Set([...document.fonts].map((f) => f.family.replace(/"/g, '')))]
      const ctx = document.createElement('canvas').getContext('2d')!
      const width = async (font: string) => {
        await document.fonts.load(font, text)
        ctx.font = font
        return ctx.measureText(text).width
      }
      const out: string[] = []
      // Monospaced bold keeps its advances; the proportional families widen.
      for (const f of families.filter((f) => !/courier|mono/i.test(f))) if ((await width(`700 15px "${f}"`)) <= (await width(`15px "${f}"`))) out.push(f)
      return out
    })
    expect(same).toEqual([])
  } finally {
    await quit(app)
  }
})
