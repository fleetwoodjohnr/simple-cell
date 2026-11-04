import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, cell, clickMenuPath } from './helpers'

test('PivotTable: insert on a new sheet, tick fields, summarize, refresh', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    const data = [['Region', 'Product', 'Sales'], ['East', 'Pens', '10'], ['West', 'Ink', '5'], ['East', 'Ink', '2']]
    for (const [r, row] of data.entries()) for (const [c, v] of row.entries()) await typeCell(page, `${'ABC'[c]}${r + 1}`, v)
    await goto(page, 'B2')
    await page.getByTestId('ribbon-tab-insert').click()
    await page.getByRole('button', { name: 'PivotTable' }).click()
    await expect(page.getByLabel('PivotTable source')).toHaveValue('Sheet1!$A$1:$C$4')
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    const panel = page.getByTestId('pivot-panel')
    await expect(panel).toBeVisible()
    await expect(page.getByTestId('sheet-tab-Sheet2')).toBeVisible()
    await panel.getByLabel('Region', { exact: true }).click()
    await expect(panel.getByLabel('Region', { exact: true })).toBeChecked()
    await panel.getByLabel('Sales', { exact: true }).click()
    await expect(panel.getByLabel('Sales', { exact: true })).toBeChecked()
    await expect.poll(() => cell(page, 'B4')).toBe(12)
    expect(await cell(page, 'A6')).toBe('Grand Total')
    expect(await cell(page, 'B6')).toBe(17)
    await panel.getByLabel('Summarize Sales by').selectOption('max')
    await expect.poll(() => cell(page, 'B3')).toBe('Max of Sales')
    expect(await cell(page, 'B4')).toBe(10)
    // Product across the top.
    await panel.getByLabel('Product', { exact: true }).click()
    await expect(panel.getByLabel('Product', { exact: true })).toBeChecked()
    await panel.getByLabel('Move Product', { exact: true }).selectOption('cols')
    await expect.poll(() => cell(page, 'B4')).toBe('Ink')
    await page.screenshot({ path: 'test-results/pivot.png' })
    // Change the source, then Refresh All.
    await page.getByTestId('sheet-tab-Sheet1').click()
    await typeCell(page, 'C2', '40')
    await page.getByTestId('sheet-tab-Sheet2').click()
    await page.getByTestId('ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Refresh all' }).click()
    await expect.poll(() => cell(page, 'C5')).toBe(40)
  } finally {
    await quit(app)
  }
})

test('data tools: Remove Duplicates and Goal Seek through their dialogs', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    for (const [r, row] of [['Name'], ['Ann'], ['Bo'], ['ann']].entries()) await typeCell(page, `A${r + 1}`, row[0])
    await goto(page, 'A2')
    await page.getByTestId('ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Remove duplicates' }).click()
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => cell(page, 'A4')).toBeNull()
    expect(await cell(page, 'A3')).toBe('Bo')
    await typeCell(page, 'C1', '100')
    await typeCell(page, 'D1', '=C1*3')
    await goto(page, 'D1')
    await page.getByRole('button', { name: 'What-if analysis' }).click()
    await clickMenuPath(page, 'Goal seek…')
    await page.getByLabel('To value').fill('600')
    await page.getByLabel('By changing cell').fill('C1')
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(page.getByText('found a solution')).toBeVisible()
    expect(await cell(page, 'C1')).toBeCloseTo(200, 6)
  } finally {
    await quit(app)
  }
})
