import { test, expect } from '@playwright/test'
import { copyFileSync, existsSync, readFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { unzipSync, strFromU8 } from 'fflate'
import { launch, quit, typeCell, cell, MOD, tempDir, stubDialogs, goto } from './helpers'

const fixtures = path.join(__dirname, '../fixtures')
const soffice = ['/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'].find((p) => existsSync(p))

async function openFile(app: Parameters<typeof stubDialogs>[0], page: import('@playwright/test').Page, file: string) {
  await stubDialogs(app, { open: [file] })
  await page.keyboard.press(`${MOD}+o`)
  await page.getByTestId('grid').waitFor()
  await page.waitForTimeout(300)
}

test('open an XLSX, edit, save, reopen: values, formulas and formatting intact', async () => {
  const dir = tempDir()
  const file = path.join(dir, 'features.xlsx')
  copyFileSync(path.join(fixtures, 'features.xlsx'), file)
  const { app, page } = await launch()
  try {
    await openFile(app, page, file)
    expect(await cell(page, 'D2')).toBe(5002)
    await expect(page.getByTestId('sheet-tab-Summary')).toBeVisible()
    await typeCell(page, 'B2', '6')
    expect(await cell(page, 'D2')).toBe(7503)
    await page.keyboard.press(`${MOD}+s`)
    await expect(page.getByTestId('save-status')).toHaveText('Saved')
    const parts = unzipSync(new Uint8Array(readFileSync(file)))
    expect(strFromU8(parts['xl/worksheets/sheet1.xml'])).toContain('<f>B2*C2</f><v>7503</v>')
    await page.keyboard.press(`${MOD}+w`)
    await openFile(app, page, file)
    expect(await cell(page, 'B2')).toBe(6)
    expect(await cell(page, 'Summary!B1')).toBeCloseTo(7704.4)
    await goto(page, 'D2')
    await expect(page.getByTestId('formula-input')).toHaveValue('=B2*C2')
    // Save as ODS, reopen
    const ods = path.join(dir, 'copy.ods')
    await stubDialogs(app, { save: ods })
    await page.keyboard.press(`${MOD}+Shift+s`)
    await expect(page.getByTestId('save-status')).toHaveText('Saved')
    await page.keyboard.press(`${MOD}+w`)
    await openFile(app, page, ods)
    expect(await cell(page, 'D2')).toBe(7503)
    expect(await cell(page, 'Summary!B2')).toBeCloseTo(199.9)
    // Saving a two-sheet workbook as CSV warns; "Save as .xlsx instead" keeps everything.
    const csv = path.join(dir, 'copy.csv')
    const kept = path.join(dir, 'copy.xlsx')
    await stubDialogs(app, { save: [csv, kept] })
    await page.keyboard.press(`${MOD}+Shift+s`)
    await page.getByTestId('save-xlsx-instead').click()
    await expect(page.getByTestId('save-status')).toHaveText('Saved')
    expect(existsSync(csv)).toBe(false)
    expect(Object.keys(unzipSync(new Uint8Array(readFileSync(kept))))).toContain('xl/worksheets/sheet2.xml')
  } finally {
    await quit(app)
  }
})

test.skip(!soffice, 'needs LibreOffice')
test('legacy XLS opens through LibreOffice; templates open untitled', async () => {
  const dir = tempDir()
  execFileSync(soffice!, ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(path.join(dir, 'p')).href}`, '--convert-to', 'xls', '--outdir', dir, path.join(fixtures, 'features.xlsx')], { timeout: 120_000 })
  const { app, page } = await launch()
  try {
    await openFile(app, page, path.join(dir, 'features.xls'))
    expect(await cell(page, 'D2')).toBe(5002)
    await expect(page.getByTestId('compat-badge')).toBeVisible()
    // Saving keeps .xls (written by LibreOffice from an .xlsx copy) after a warning.
    const xls = path.join(dir, 'saved.xls')
    await typeCell(page, 'B2', '7')
    await stubDialogs(app, { save: xls })
    await page.keyboard.press(`${MOD}+s`)
    await page.getByRole('button', { name: 'Save anyway' }).click()
    await expect(page.getByTestId('save-status')).toHaveText('Saved')
    expect([...readFileSync(xls).subarray(0, 4)]).toEqual([0xd0, 0xcf, 0x11, 0xe0]) // OLE compound file (BIFF8)
    await page.keyboard.press(`${MOD}+w`)
    await openFile(app, page, xls)
    expect(await cell(page, 'D2')).toBe(8753.5)
    // An .xltx template opens as a new workbook (no path)
    const tpl = path.join(dir, 'report.xltx')
    copyFileSync(path.join(fixtures, 'features.xlsx'), tpl)
    await openFile(app, page, tpl)
    expect(await cell(page, 'A2')).toBe('Widget')
    await expect(page.getByTestId('save-status')).toHaveText('New')
  } finally {
    await quit(app)
  }
})
