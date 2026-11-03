import { test, expect, type Page } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, cell, MOD } from './helpers'

const hiddenRows = (page: Page) =>
  page.evaluate(() => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const m = sc.books.get(tab).sheetMeta(sc.getView(tab).sheet)
    return m.rows.filter(([, h]: [number, number]) => h === 0).map(([r]: [number, number]) => r).sort((a: number, b: number) => a - b)
  })

test('tables: Ctrl+T, total row, header filter menu, Ctrl+Shift+L', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    const data = [
      ['Region', 'Item', 'Units'],
      ['East', 'Pens', '10'],
      ['West', 'Ink', '5'],
      ['East', 'Paper', '25'],
      ['North', 'Pens', '7']
    ]
    for (const [r, row] of data.entries()) for (const [c, v] of row.entries()) await typeCell(page, `${'ABC'[c]}${r + 1}`, v)
    await goto(page, 'B3')
    await page.keyboard.press(`${MOD}+t`)
    await expect(page.getByLabel('Table range')).toHaveValue('$A$1:$C$5')
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    // Table Design appears while the cursor is in the table.
    await page.getByTestId('ribbon-tab-table').click()
    await page.getByLabel('Total row').click()
    await expect(page.getByLabel('Total row')).toBeChecked()
    await expect.poll(() => cell(page, 'C6')).toBe(47)
    expect(await cell(page, 'A6')).toBe('Total')
    // Filter from the header button: untick West and North.
    await page.getByTestId('filter-button-A1').click()
    const menu = page.getByTestId('filter-menu')
    await expect(menu.getByText('North')).toBeVisible()
    await menu.getByLabel('West').uncheck()
    await menu.getByLabel('North').uncheck()
    await menu.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => hiddenRows(page)).toEqual([2, 4])
    await expect.poll(() => cell(page, 'C6')).toBe(35) // SUBTOTAL skips the filtered rows
    await expect(page.getByTestId('filter-button-A1')).toHaveAccessibleName(/filtered/)
    // A condition on another column.
    await page.getByTestId('filter-button-C1').click()
    await menu.getByRole('button', { name: 'Condition', exact: true }).click()
    await menu.getByLabel('Condition', { exact: true }).selectOption('greaterThan')
    await menu.getByLabel('Condition value').fill('20')
    await menu.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 4])
    await page.keyboard.press(`${MOD}+z`)
    await page.keyboard.press(`${MOD}+z`)
    await expect.poll(() => hiddenRows(page)).toEqual([])
    // Ctrl+Shift+L on plain data outside the table: a sheet AutoFilter.
    await typeCell(page, 'F1', 'Name')
    await typeCell(page, 'F2', 'Ann')
    await typeCell(page, 'F3', 'Bob')
    await goto(page, 'F2')
    await page.keyboard.press(`${MOD}+Shift+l`)
    await expect(page.getByTestId('filter-button-F1')).toBeVisible()
    await page.getByTestId('filter-button-F1').click()
    await menu.getByLabel('Search values').fill('bo')
    await menu.getByLabel('Search values').press('Enter')
    await expect.poll(() => hiddenRows(page)).toEqual([1])
    await page.keyboard.press(`${MOD}+Shift+l`)
    await expect(page.getByTestId('filter-button-F1')).toHaveCount(0)
    await expect.poll(() => hiddenRows(page)).toEqual([])
    await page.screenshot({ path: 'test-results/tables.png' })
  } finally {
    await quit(app)
  }
})
