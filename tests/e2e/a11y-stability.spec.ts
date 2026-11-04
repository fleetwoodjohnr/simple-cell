import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, waitForUiSettled } from './helpers'

test('SQL source targets stay at least 24px tall during the dialog entrance scale', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', 'Units')
    await typeCell(page, 'A2', '7')
    await goto(page, 'A1:A2')
    await page.getByTestId('ribbon-tab-automate').click()
    await page.getByRole('button', { name: 'SQL query', exact: true }).click()
    const dialog = page.getByRole('dialog')
    await expect(dialog.getByRole('button', { name: 'Units', exact: true })).toBeVisible()
    // An accessibility scan must not inspect the intermediate colors of an opening dialog.
    await dialog.evaluate((el) => { el.animate([{ opacity: 0 }, { opacity: 1 }], { duration: 200 }) })
    await waitForUiSettled(page)
    await expect(dialog).toHaveCSS('opacity', '1')
    // Hold the entrance animation's smallest scale, independent of runner speed or OS.
    await dialog.evaluate((el) => {
      for (const animation of el.getAnimations()) animation.finish()
      el.style.scale = '0.97'
    })
    const heights = await dialog.getByRole('list', { name: 'Tables and columns' }).locator('button').evaluateAll((buttons) => buttons.map((button) => button.getBoundingClientRect().height))
    expect(heights.length).toBeGreaterThan(0)
    for (const height of heights) expect(height).toBeGreaterThanOrEqual(24)
  } finally { await quit(app) }
})
