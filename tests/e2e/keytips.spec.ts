import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto } from './helpers'

test.skip(process.platform === 'darwin', 'Excel for Mac has no KeyTips')
test('KeyTips: Alt, a tab letter, then a command letter', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', 'x')
    await goto(page, 'A1')
    await page.keyboard.press('Alt')
    const layer = page.getByTestId('keytips')
    await expect(layer.locator('[data-for]').filter({ hasText: /^A$/ })).toHaveCount(1) // Data tab
    await page.keyboard.press('h')
    const bold = layer.locator('[data-for="Bold"]')
    await expect(bold).toBeVisible()
    const keys = (await bold.textContent())!
    for (const k of keys) await page.keyboard.press(k.toLowerCase())
    await expect(layer).toHaveCount(0)
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sc = (window as any).__sc
          const tab = sc.get().active
          return sc.books.get(tab).api.styleAt(sc.getView(tab).sheet, 0, 0).then((s: { bold?: boolean }) => !!s.bold)
        })
      )
      .toBe(true)
    // Esc steps back to the tabs, then closes.
    await page.keyboard.press('Alt')
    await page.keyboard.press('a')
    await expect(layer.locator('[data-for="Filter"]')).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(layer.locator('[data-for="Filter"]')).toHaveCount(0)
    await page.keyboard.press('Escape')
    await expect(layer).toHaveCount(0)
  } finally {
    await quit(app)
  }
})
