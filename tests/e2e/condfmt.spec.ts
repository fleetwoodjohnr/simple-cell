import { test, expect, type Page } from '@playwright/test'
import path from 'node:path'
import { launch, quit, newWorkbook, typeCell, goto, MOD, tempDir, stubDialogs } from './helpers'

/** What the grid draws for a cell: the cached view cell (conditional style, icon, bar). */
async function drawn(page: Page, ref: string) {
  return page.evaluate((ref) => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const m = /^([A-Z]+)(\d+)$/.exec(ref)!
    return sc.books.get(tab).cell(sc.getView(tab).sheet, Number(m[2]) - 1, m[1].charCodeAt(0) - 65) ?? null
  }, ref)
}

async function menu(page: Page, ...path: string[]) {
  await page.getByTestId('ribbon').getByRole('button', { name: 'Conditional formatting', exact: true }).click()
  for (const p of path) await page.getByRole('menuitem', { name: p, exact: true }).click()
}

test('conditional formatting: highlight rule, icon set, manage, undo, save and reopen', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    for (const [i, v] of ['5', '15', '25', '35'].entries()) await typeCell(page, `A${i + 1}`, v)
    await goto(page, 'A1:A4')
    await menu(page, 'Highlight cells rules', 'Greater than…')
    await page.getByLabel('Value').fill('20')
    await page.getByRole('button', { name: 'OK' }).click()
    await expect.poll(async () => (await drawn(page, 'A3'))?.x?.fill).toBe('#FFC7CE')
    expect((await drawn(page, 'A1'))?.x).toBeUndefined()
    // The format follows the data.
    await typeCell(page, 'A1', '50')
    await expect.poll(async () => (await drawn(page, 'A1'))?.x?.fill).toBe('#FFC7CE')
    await goto(page, 'A1:A4')
    await menu(page, 'Icon sets', '3 arrows')
    await expect.poll(async () => (await drawn(page, 'A1'))?.icon).toEqual({ set: '3Arrows', i: 2 })
    expect((await drawn(page, 'A2'))?.icon).toEqual({ set: '3Arrows', i: 0 })
    // Manage rules: the icon set is first (newest); delete it.
    await menu(page, 'Manage rules…')
    const rows = page.getByRole('option')
    await expect(rows).toHaveCount(2)
    await expect(rows.nth(0)).toContainText('Icon set')
    await expect(rows.nth(1)).toContainText('Cell value > 20')
    await rows.nth(0).click()
    await page.getByRole('button', { name: 'Delete rule' }).click()
    await expect(rows).toHaveCount(1)
    await page.getByRole('button', { name: 'Close' }).first().click()
    await expect.poll(async () => (await drawn(page, 'A1'))?.icon).toBeUndefined()
    await page.keyboard.press(`${MOD}+z`)
    await expect.poll(async () => (await drawn(page, 'A1'))?.icon).toEqual({ set: '3Arrows', i: 2 })
    // Saved as .xlsx and reopened, the rules are still there.
    const file = path.join(tempDir(), 'cf.xlsx')
    await stubDialogs(app, { save: file, open: [file] })
    await page.keyboard.press(`${MOD}+s`)
    await expect(page.getByTestId('save-status')).toHaveText('Saved')
    await page.keyboard.press(`${MOD}+w`)
    await page.keyboard.press(`${MOD}+o`)
    await page.getByTestId('grid').waitFor()
    await expect.poll(async () => (await drawn(page, 'A4'))?.x?.fill).toBe('#FFC7CE')
    expect((await drawn(page, 'A4'))?.icon).toEqual({ set: '3Arrows', i: 1 })
    await page.screenshot({ path: 'test-results/condfmt.png' })
  } finally {
    await quit(app)
  }
})
