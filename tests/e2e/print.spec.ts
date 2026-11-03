import { test, expect } from '@playwright/test'
import { existsSync, readFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import path from 'node:path'
import { launch, quit, newWorkbook, typeCell, MOD, tempDir, stubDialogs } from './helpers'

const pdftotext = ['/usr/bin/pdftotext', '/usr/local/bin/pdftotext', '/opt/homebrew/bin/pdftotext'].find((p) => existsSync(p))

test('print dialog: pages, footer, Save as PDF; the PDF holds the sheet text', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', 'Quarterly report')
    await typeCell(page, 'A2', 'Revenue')
    await typeCell(page, 'B2', '1250.5')
    // Rows far down make a second page.
    await typeCell(page, 'A80', 'Last line')
    await page.keyboard.press(`${MOD}+p`)
    await expect(page.getByTestId('print-pages')).toHaveText('2 pages')
    await page.getByLabel('Footer').fill('Page &P of &N')
    await page.getByLabel('Footer').blur()
    await page.getByLabel('Scaling').selectOption('sheet')
    await expect(page.getByTestId('print-pages')).toHaveText('1 page')
    const file = path.join(tempDir(), 'report.pdf')
    await stubDialogs(app, { save: file })
    await page.getByRole('button', { name: 'Save as PDF' }).click()
    await expect.poll(() => existsSync(file), { timeout: 15_000 }).toBe(true)
    expect(readFileSync(file).subarray(0, 5).toString()).toBe('%PDF-')
    if (!pdftotext) return
    const text = execFileSync(pdftotext, ['-layout', file, '-']).toString()
    expect(text).toContain('Quarterly report')
    expect(text).toContain('1250.5')
    expect(text).toContain('Last line')
    expect(text).toContain('Page 1 of 1')
    expect(execFileSync('pdfinfo', [file]).toString()).toMatch(/Pages:\s+1/)
  } finally {
    await quit(app)
  }
})
