import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, cell, MOD } from './helpers'

// Runs against the packaged renderer and its CSP, so it also proves SQLite's WebAssembly is allowed to load.
test('Automate ▸ SQL: preview a query, send its rows to a new sheet, then UPDATE the table in place', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    const data = [
      ['Region', 'Item', 'Units'],
      ['East', 'Pens', '10'],
      ['West', 'Ink', '5'],
      ['East', 'Paper', '25']
    ]
    for (const [r, row] of data.entries()) for (const [c, v] of row.entries()) await typeCell(page, `${'ABC'[c]}${r + 1}`, v)
    await goto(page, 'B2')
    await page.keyboard.press(`${MOD}+t`)
    await page.getByRole('button', { name: 'OK', exact: true }).click()

    await page.getByTestId('ribbon-tab-automate').click()
    await page.getByRole('button', { name: 'SQL query', exact: true }).click()
    const editor = page.getByRole('textbox', { name: 'SQL query' })
    await expect(editor).toHaveValue('SELECT *\nFROM Table1')
    await editor.fill("SELECT Item, Units FROM Table1 WHERE Region = 'East' ORDER BY Units DESC")
    await editor.press(`${MOD}+Enter`)
    await expect(page.getByTestId('sql-summary')).toHaveText('2 rows')
    await page.getByRole('button', { name: 'Run', exact: true }).click()
    await expect.poll(() => cell(page, 'Query!A2').catch(() => undefined)).toBe('Paper')
    expect(await cell(page, 'Query!B3')).toBe(10)
    await page.keyboard.press(`${MOD}+z`)
    await expect(page.getByTestId('sheet-tab-Query')).toHaveCount(0)
    await expect.poll(() => page.evaluate(() => { const sc = (window as any).__sc; const tab = sc.get().active; return sc.books.get(tab).meta.sheets.some((s: { id: number }) => s.id === sc.getView(tab).sheet) })).toBe(true)
    await page.keyboard.press(`${MOD}+y`)
    await expect(page.getByTestId('sheet-tab-Query')).toBeVisible()

    await page.getByRole('button', { name: 'SQL query', exact: true }).click()
    await page.getByRole('textbox', { name: 'SQL query' }).fill("UPDATE Table1 SET Units = Units * 2 WHERE Item = 'Ink'")
    await page.getByRole('button', { name: 'Run', exact: true }).click()
    await expect.poll(() => cell(page, 'Sheet1!C3')).toBe(10)
    await page.keyboard.press(`${MOD}+z`)
    await expect.poll(() => cell(page, 'Sheet1!C3')).toBe(5)
  } finally {
    await quit(app)
  }
})

test('SQL selected range: ordinary range output, source edits, cancellation', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    for (const [ref, v] of [['A1', 'Units'], ['A2', '3'], ['A3', '7']]) await typeCell(page, ref, v)
    await goto(page, 'A1:A3')
    await page.getByTestId('ribbon-tab-automate').click()
    const open = async () => page.getByRole('button', { name: 'SQL query', exact: true }).click()
    await open()
    await expect(page.getByRole('textbox', { name: 'Source range' })).toHaveValue('A1:A3')
    await page.getByRole('textbox', { name: 'SQL query' }).fill('SELECT Units * 2 AS Double FROM Selection')
    await page.getByRole('radio', { name: 'Output range', exact: true }).click()
    await page.getByRole('textbox', { name: 'Top-left cell' }).fill('D1')
    await page.getByRole('button', { name: 'Run', exact: true }).click()
    await expect.poll(() => cell(page, 'D3')).toBe(14)
    expect(await page.evaluate(() => { const sc = (window as any).__sc; return sc.books.get(sc.get().active).meta.tables.length })).toBe(0)
    await goto(page, 'A1:A3'); await open()
    await page.getByRole('textbox', { name: 'SQL query' }).fill('UPDATE Selection SET Units=Units+1')
    await page.getByRole('button', { name: 'Run', exact: true }).click()
    await expect.poll(() => cell(page, 'A2')).toBe(4)
    await page.keyboard.press(`${MOD}+z`)
    await expect.poll(() => cell(page, 'A2')).toBe(3)
    await open()
    await page.getByRole('textbox', { name: 'SQL query' }).fill('WITH RECURSIVE x(n) AS (VALUES(1) UNION ALL SELECT n+1 FROM x) SELECT SUM(n) FROM x')
    await page.getByRole('button', { name: 'Preview', exact: true }).click()
    await page.waitForTimeout(250)
    await page.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect.poll(() => cell(page, 'A2')).toBe(3)
    await open()
    await page.getByRole('textbox', { name: 'SQL query' }).fill('SELECT 42 AS Answer')
    await page.getByRole('button', { name: 'Preview', exact: true }).click()
    await expect(page.getByTestId('sql-summary')).toHaveText('1 row')
  } finally { await quit(app) }
})
