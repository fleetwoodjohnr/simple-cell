import { test, expect, type Page } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto } from './helpers'

const hiddenRows = (page: Page) =>
  page.evaluate(() => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const m = sc.books.get(tab).sheetMeta(sc.getView(tab).sheet)
    return m.rows.filter(([, h]: [number, number]) => h === 0).map(([r]: [number, number]) => r).sort((a: number, b: number) => a - b)
  })

test('group rows (Shift+Alt+Right), collapse and expand with the outline buttons, levels, ungroup', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    for (let i = 1; i <= 6; i++) await typeCell(page, `A${i}`, String(i))
    await goto(page, 'A2:A4')
    await page.keyboard.press('Shift+Alt+ArrowRight')
    const minus = page.getByTestId('outline-row-2-4')
    await expect(minus).toHaveText('−')
    await minus.click()
    await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 3])
    await expect(page.getByTestId('outline-row-2-4')).toHaveText('+')
    await page.getByRole('button', { name: 'Show row level 2' }).click()
    await expect.poll(() => hiddenRows(page)).toEqual([])
    await page.getByRole('button', { name: 'Show row level 1' }).click()
    await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 3])
    await page.screenshot({ path: 'test-results/outline.png', clip: { x: 0, y: 280, width: 500, height: 200 } })
    await page.getByTestId('outline-row-2-4').click()
    await goto(page, 'A2:A4')
    await page.keyboard.press('Shift+Alt+ArrowLeft')
    await expect(page.getByTestId('outline-row-2-4')).toHaveCount(0)
  } finally {
    await quit(app)
  }
})
