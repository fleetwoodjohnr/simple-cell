import { test, expect, type Page } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, MOD } from './helpers'

const charts = (page: Page) =>
  page.evaluate(() => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    return sc.books.get(tab).sheetMeta(sc.getView(tab).sheet).charts
  })

test('charts: Alt+F1 from data, edit, move, delete, undo', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    const data = [['Month', 'North', 'South'], ['Jan', '10', '7'], ['Feb', '12', '9'], ['Mar', '15', '11']]
    for (const [r, row] of data.entries()) for (const [c, v] of row.entries()) await typeCell(page, `${'ABC'[c]}${r + 1}`, v)
    await goto(page, 'B2')
    await page.keyboard.press('Alt+F1')
    await expect.poll(async () => (await charts(page)).length).toBe(1)
    const [chart] = await charts(page)
    const el = page.getByTestId(`chart-${chart.id}`)
    await expect(el).toBeVisible()
    await expect(el.locator('canvas')).toHaveCount(1)
    await page.screenshot({ path: 'test-results/chart.png' })
    // The chart is on the printed page.
    await page.keyboard.press(`${MOD}+p`)
    await expect(page.getByTestId('print-preview').locator('img')).toHaveCount(1)
    await expect.poll(() => page.getByTestId('print-preview').locator('img').evaluate((i: HTMLImageElement) => i.naturalWidth)).toBeGreaterThan(0)
    await page.getByRole('button', { name: 'Cancel' }).click()
    // Double-click opens the editor: make it a line chart with a title.
    await el.dblclick()
    await page.getByRole('radio', { name: 'Line' }).click()
    await page.getByLabel('Chart title').fill('Sales by month')
    await page.getByRole('button', { name: 'OK' }).click()
    await expect.poll(async () => (await charts(page))[0]).toMatchObject({ type: 'line', title: 'Sales by month' })
    // Drag it to another place.
    const box = (await el.boundingBox())!
    await page.mouse.move(box.x + 40, box.y + 40)
    await page.mouse.down()
    await page.mouse.move(box.x + 40, box.y + 140, { steps: 5 })
    await page.mouse.up()
    await expect.poll(async () => (await charts(page))[0].anchor.r0).toBeGreaterThan(chart.anchor.r0 + 2)
    // Delete with the chart selected; undo brings it back.
    await el.click()
    await page.keyboard.press('Delete')
    await expect.poll(async () => (await charts(page)).length).toBe(0)
    await page.keyboard.press(`${MOD}+z`)
    await expect.poll(async () => (await charts(page)).length).toBe(1)
  } finally {
    await quit(app)
  }
})

test('sparklines: Insert ▸ Line from the selection draws one per row', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    const rows = [['3', '5', '2', '8', '6'], ['-1', '4', '-2', '6', '3']]
    for (const [r, row] of rows.entries()) for (const [c, v] of row.entries()) await typeCell(page, `${'ABCDE'[c]}${r + 1}`, v)
    await goto(page, 'A1:E2')
    await page.getByTestId('ribbon-tab-insert').click()
    await page.getByRole('button', { name: 'Line sparklines' }).click()
    await expect(page.getByLabel('Location range')).toHaveValue('F1:F2')
    await page.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sc = (window as any).__sc
          const tab = sc.get().active
          return sc.books.get(tab).cell(sc.getView(tab).sheet, 1, 5)?.spark?.values
        })
      )
      .toEqual([-1, 4, -2, 6, 3])
    await page.getByTestId('ribbon-tab-home').click()
    const box = await page.getByTestId('grid').boundingBox()
    await page.screenshot({ path: 'test-results/sparklines.png', clip: { x: box!.x, y: box!.y, width: 520, height: 80 } })
  } finally {
    await quit(app)
  }
})
