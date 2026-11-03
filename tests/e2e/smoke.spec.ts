import { test, expect } from '@playwright/test'
import { launch, newWorkbook, quit, typeCell, cell, shown, MOD, goto } from './helpers'

test('create a workbook, enter values and formulas, undo/redo', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', '10')
    await typeCell(page, 'A2', '32')
    await typeCell(page, 'A3', '=SUM(A1:A2)*2')
    expect(await cell(page, 'A3')).toBe(84)
    await typeCell(page, 'A1', '20')
    expect(await cell(page, 'A3')).toBe(104)
    await page.keyboard.press(`${MOD}+z`)
    await page.waitForTimeout(100)
    expect(await cell(page, 'A3')).toBe(84)
    await page.keyboard.press(`${MOD}+y`)
    await page.waitForTimeout(100)
    expect(await cell(page, 'A3')).toBe(104)
    // Formula bar shows the formula of the active cell
    await goto(page, 'A3')
    await expect(page.getByTestId('formula-input')).toHaveValue('=SUM(A1:A2)*2')
    await typeCell(page, 'B1', '25%')
    expect(await shown(page, 'B1')).toBe('25%')
    await page.screenshot({ path: 'test-results/smoke.png' })
  } finally {
    await quit(app)
  }
})
