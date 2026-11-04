// Opt-in (SC_SCREENSHOTS=1): the README screenshots in docs/screenshots, taken from the real app.
import { test, expect, type Page } from '@playwright/test'
import { launch, quit, newWorkbook, goto } from './helpers'

test.skip(!process.env.SC_SCREENSHOTS, 'screenshot run only')

/** Type rows into the active sheet from A1, as the user would (formulas and formats parsed). */
async function fill(page: Page, rows: string[][]) {
  await page.evaluate(async (rows) => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const s = sc.getView(tab).sheet
    for (const [r, row] of rows.entries()) for (const [c, t] of row.entries()) await sc.books.get(tab).api.input(s, r, c, t)
  }, rows)
}

/** Quarterly sales: a styled table with a total row, data bars, currency, a chart, fitted columns. */
async function salesWorkbook(page: Page) {
  await newWorkbook(page)
  const regions: [string, number[]][] = [['North', [42150, 45800, 51200, 58900]], ['South', [38400, 36100, 41750, 47300]], ['East', [51900, 55200, 49800, 62400]], ['West', [29700, 33900, 38600, 44100]], ['Central', [35200, 37800, 40900, 43500]]]
  await fill(page, [['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'], ...regions.map(([n, q], i) => [n, ...q.map(String), `=SUM(B${i + 2}:E${i + 2})`])])
  await page.evaluate(async () => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const api = sc.books.get(tab).api
    const s = sc.getView(tab).sheet
    await api.formatCells([{ sheet: s, r0: 1, c0: 1, r1: 6, c1: 5 }], { numFmt: '$#,##0' })
    await api.createTable(s, { sheet: s, r0: 0, c0: 0, r1: 5, c1: 5 }, true, 'TableStyleMedium2')
    await api.addCondFormat(s, [{ sheet: s, r0: 1, c0: 5, r1: 5, c1: 5 }], { type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '#63be7b' } })
    await api.addChartFromRange(s, { sheet: s, r0: 0, c0: 0, r1: 5, c1: 4 }, 'column')
  })
  await goto(page, 'A1:F6')
  await page.evaluate(() => (window as any).__sc.runCommand('fmt.autofitCols'))
  await goto(page, 'F3')
  await page.waitForTimeout(500)
}

// ponytail: light theme only; dark mode paints light text on light fills (table stripes, data bars, charts), add a dark shot once that's fixed
test('README screenshots', async () => {
  const { app, page } = await launch({ settings: { theme: 'light' } })
  try {
    await salesWorkbook(page)
    await page.screenshot({ path: 'docs/screenshots/main.png' })
    // A PivotTable of order lines: Region down, Product across, Sales summed, with the field list open.
    await page.evaluate(() => (window as any).__sc.runCommand('sheet.insert'))
    const products = ['Pens', 'Paper', 'Ink']
    const lines = Array.from({ length: 24 }, (_, i) => [['North', 'South', 'East', 'West'][i % 4], products[i % 3], String(120 + ((i * 37) % 90) * 5)])
    await fill(page, [['Region', 'Product', 'Sales'], ...lines])
    await goto(page, 'A2')
    await page.getByTestId('ribbon-tab-insert').click()
    await page.getByRole('button', { name: 'PivotTable' }).click()
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    const panel = page.getByTestId('pivot-panel')
    for (const f of ['Region', 'Product', 'Sales']) {
      await panel.getByLabel(f, { exact: true }).click()
      await expect(panel.getByLabel(f, { exact: true })).toBeChecked()
    }
    await panel.getByLabel('Move Product', { exact: true }).selectOption('cols')
    await page.waitForTimeout(500)
    await page.screenshot({ path: 'docs/screenshots/pivot.png' })
  } finally {
    await quit(app)
  }
})
