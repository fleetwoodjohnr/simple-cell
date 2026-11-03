import { test, expect } from '@playwright/test'
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'
import { launch, newWorkbook, quit, typeCell, cell, shown, MOD, goto, tempDir, stubDialogs } from './helpers'

test('copy and paste moves relative references; cut moves cells', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', '1')
    await typeCell(page, 'A2', '2')
    await typeCell(page, 'B1', '=A1*10')
    await goto(page, 'B1')
    await page.keyboard.press(`${MOD}+c`)
    await goto(page, 'B2')
    await page.keyboard.press(`${MOD}+v`)
    await page.waitForTimeout(150)
    expect(await cell(page, 'B2')).toBe(20)
    await goto(page, 'B2')
    await expect(page.getByTestId('formula-input')).toHaveValue('=A2*10')
    // Paste from another application: tab-separated text becomes numbers
    await page.evaluate(() => navigator.clipboard.writeText('5\t6\n7\t8'))
    await goto(page, 'D1')
    await page.keyboard.press(`${MOD}+v`)
    await page.waitForTimeout(150)
    expect(await cell(page, 'E2')).toBe(8)
    // Fill down
    await goto(page, 'B1:B4')
    await page.keyboard.press(`${MOD}+d`)
    await page.waitForTimeout(150)
    await goto(page, 'B4')
    await expect(page.getByTestId('formula-input')).toHaveValue('=A4*10')
  } finally {
    await quit(app)
  }
})

test('insert and delete rows, sheets, rename, navigation and status bar', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    for (let i = 1; i <= 5; i++) await typeCell(page, `A${i}`, String(i * 10))
    await typeCell(page, 'C1', '=SUM(A1:A5)')
    expect(await cell(page, 'C1')).toBe(150)
    // Insert a row above row 3 through the command
    await goto(page, 'A3')
    await page.keyboard.press('Shift+Space')
    await page.keyboard.press(`${MOD}+Shift+=`)
    await page.waitForTimeout(150)
    expect(await cell(page, 'A4')).toBe(30)
    await goto(page, 'C1')
    await expect(page.getByTestId('formula-input')).toHaveValue('=SUM(A1:A6)')
    // Delete it again
    await goto(page, 'A3')
    await page.keyboard.press('Shift+Space')
    await page.keyboard.press(`${MOD}+Minus`)
    await page.waitForTimeout(150)
    expect(await cell(page, 'A3')).toBe(30)
    // Ctrl+Plus on a cell asks how to shift; shifting A2 down moves only column A, and the SUM grows.
    await goto(page, 'A2')
    await page.keyboard.press(`${MOD}+Shift+=`)
    await page.getByLabel('Shift cells down').check()
    await page.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => cell(page, 'A3')).toBe(20)
    expect(await cell(page, 'A2')).toBeNull()
    await goto(page, 'C1')
    await expect(page.getByTestId('formula-input')).toHaveValue('=SUM(A1:A6)')
    await page.keyboard.press(`${MOD}+z`)
    await expect.poll(() => cell(page, 'A2')).toBe(20)
    // Ctrl+Down jumps to the end of the data
    await goto(page, 'A1')
    await page.keyboard.press(`${MOD}+ArrowDown`)
    await page.waitForTimeout(100)
    await expect(page.getByTestId('name-box')).toHaveValue('A5')
    // Status bar quick calculations
    await page.keyboard.press(`${MOD}+Shift+ArrowUp`)
    await page.waitForTimeout(200)
    await expect(page.getByTestId('quick-stats')).toContainText('Sum: 150')
    await expect(page.getByTestId('quick-stats')).toContainText('Average: 30')
    // New sheet, rename by double-click, cross-sheet formula
    await page.getByTestId('add-sheet').click()
    await page.waitForTimeout(150)
    await page.getByTestId('sheet-tab-Sheet2').dblclick()
    await page.keyboard.press(`${MOD}+a`)
    await page.keyboard.type('Totals')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(150)
    await expect(page.getByTestId('sheet-tab-Totals')).toBeVisible()
    await typeCell(page, 'A1', '=Sheet1!C1+1')
    expect(await cell(page, 'A1')).toBe(151)
  } finally {
    await quit(app)
  }
})

test('save as CSV, close, reopen: values intact; formula-injection guard', async () => {
  const dir = tempDir()
  const file = path.join(dir, 'data.csv')
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', 'Name')
    await typeCell(page, 'B1', 'Amount')
    await typeCell(page, 'A2', 'Widget, large')
    await typeCell(page, 'B2', '1250.5')
    await typeCell(page, 'A3', "'=HYPERLINK(\"http://x\")")
    await typeCell(page, 'B3', '=B2*2')
    await stubDialogs(app, { save: file })
    await page.keyboard.press(`${MOD}+s`)
    // CSV drops formulas: the warning is confirmed
    await page.getByRole('button', { name: 'Save anyway' }).click()
    await expect(page.getByTestId('save-status')).toHaveText('Saved')
    const text = readFileSync(file, 'utf8')
    expect(text).toContain('"Widget, large",1250.5')
    expect(text).toContain(`"'=HYPERLINK(""http://x"")",2501`)
    await page.keyboard.press(`${MOD}+w`)
    await stubDialogs(app, { open: [file] })
    await page.keyboard.press(`${MOD}+o`)
    await page.getByTestId('grid').waitFor()
    await page.waitForTimeout(200)
    expect(await cell(page, 'B2')).toBe(1250.5)
    expect(await cell(page, 'A2')).toBe('Widget, large')
    expect(await shown(page, 'B3')).toBe('2501')
  } finally {
    await quit(app)
  }
})

test('unsaved work survives a renderer crash', async () => {
  const userData = tempDir('sc-profile-')
  writeFileSync(path.join(userData, 'settings.json'), JSON.stringify({ recoveryIntervalSec: 1 }))
  const { app, page } = await launch({ userData })
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', 'Survives a crash')
    await typeCell(page, 'A2', '=LEN(A1)')
    await expect.poll(() => (existsSync(path.join(userData, 'recovery')) ? readdirSync(path.join(userData, 'recovery')).length : 0), { timeout: 10_000 }).toBeGreaterThan(0)
    await page.waitForTimeout(1500)
    await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.forcefullyCrashRenderer())
    const js = (code: string) => app.evaluate(({ BrowserWindow }, c) => BrowserWindow.getAllWindows()[0].webContents.executeJavaScript(c), code)
    await expect.poll(() => js(`document.querySelectorAll('[data-testid=recovery] button').length`).catch(() => -1), { timeout: 20_000 }).toBeGreaterThan(0)
    await js(`[...document.querySelectorAll('[data-testid=recovery] button')].find((b) => b.textContent.trim() === 'Restore').click()`)
    const value = (ref: [number, number]) => js(`(async () => { const sc = window.__sc; const t = sc.get().active; const b = sc.books.get(t); return b ? b.api.value(sc.getView(t).sheet, ${ref[0]}, ${ref[1]}) : null })()`)
    await expect.poll(() => value([0, 0]).catch(() => null), { timeout: 10_000 }).toBe('Survives a crash')
    expect(await value([1, 0])).toBe(16)
  } finally {
    await quit(app)
  }
})

test('keyboard: editing modes, F2, Escape, Alt+Enter, Ctrl+Enter, F4, point mode, AutoSum', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', '3')
    await typeCell(page, 'A2', '4')
    // Point mode: type "=" then arrow keys insert a reference
    await goto(page, 'B2')
    await page.keyboard.type('=')
    await page.keyboard.press('ArrowLeft')
    await page.keyboard.type('*2')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(100)
    expect(await cell(page, 'B2')).toBe(8)
    // F4 cycles references
    await goto(page, 'C1')
    await page.keyboard.type('=A1')
    await page.keyboard.press('F4')
    await page.keyboard.press('Enter')
    await goto(page, 'C1')
    await expect(page.getByTestId('formula-input')).toHaveValue('=$A$1')
    // Escape cancels
    await goto(page, 'A1')
    await page.keyboard.type('999')
    await page.keyboard.press('Escape')
    expect(await cell(page, 'A1')).toBe(3)
    // F2 edits in place and keeps the content
    await page.keyboard.press('F2')
    await page.keyboard.type('0')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(100)
    expect(await cell(page, 'A1')).toBe(30)
    // Alt+Enter makes a line break
    await goto(page, 'D1')
    await page.keyboard.type('one')
    await page.keyboard.press('Alt+Enter')
    await page.keyboard.type('two')
    await page.keyboard.press('Enter')
    expect(await cell(page, 'D1')).toBe('one\ntwo')
    // Ctrl+Enter fills the selection
    await goto(page, 'E1:E3')
    await page.keyboard.type('x')
    await page.keyboard.press(`${MOD}+Enter`)
    await page.waitForTimeout(100)
    expect(await cell(page, 'E3')).toBe('x')
    // AutoSum below numbers
    await goto(page, 'A3')
    await page.keyboard.press('Alt+=')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(100)
    expect(await cell(page, 'A3')).toBe(34)
    // Bold toggles with Ctrl+B
    await goto(page, 'A1')
    await page.keyboard.press(`${MOD}+b`)
    await page.waitForTimeout(100)
    const bold = await page.evaluate(async () => {
      const sc = (window as any).__sc
      const t = sc.get().active
      return (await sc.books.get(t).api.styleAt(sc.getView(t).sheet, 0, 0)).bold
    })
    expect(bold).toBe(true)
  } finally {
    await quit(app)
  }
})
