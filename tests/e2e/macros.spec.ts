import { test, expect, type Page } from '@playwright/test'
import path from 'node:path'
import { launch, quit, newWorkbook, typeCell, cell, MOD, tempDir, stubDialogs } from './helpers'

async function setCode(page: Page, code: string) {
  await page.locator('[data-testid="macro-editor"] .cm-content').click()
  await page.keyboard.press(`${MOD}+a`)
  await page.keyboard.press('Delete')
  await page.keyboard.insertText(code)
}

test('macros: run against the workbook in a sandbox; stop; trust before running a file’s macros', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', '21')
    await page.keyboard.press('Alt+F8')
    await expect(page.locator('[data-testid="macro-editor"] .cm-content')).toContainText('async function main')
    await page.getByRole('button', { name: 'Run' }).click()
    await expect(page.getByTestId('macro-log')).toContainText('✔ Finished')
    expect(await cell(page, 'B1')).toBe(42)
    // The sandbox: no network, no app, no DOM — even if the macro tries to get them back.
    await setCode(page, `async function main(wb) {
  const out = []
  out.push(typeof fetch, typeof XMLHttpRequest, typeof WebSocket, typeof importScripts, typeof document, typeof window, typeof require, typeof process)
  try { self.fetch = () => 1 } catch (e) { out.push('locked') }
  out.push(typeof self.fetch)
  await wb.activeSheet.range('C1').setValue(out.join(','))
}`)
    await page.getByRole('button', { name: 'Run' }).click()
    await expect(page.getByTestId('macro-log')).toContainText('✔ Finished')
    expect(await cell(page, 'C1')).toBe('undefined,undefined,undefined,undefined,undefined,undefined,undefined,undefined,locked,undefined')
    // Errors are reported with their line; an endless loop can be stopped.
    await setCode(page, 'await workbook.activeSheet.range("not a range").getValue()')
    await page.getByRole('button', { name: 'Run' }).click()
    await expect(page.getByTestId('macro-log')).toContainText('isn’t a range address')
    await setCode(page, 'while (true) {}')
    await page.getByRole('button', { name: 'Run' }).click()
    await page.getByRole('button', { name: 'Stop' }).click()
    await expect(page.getByTestId('macro-log')).toContainText('✖ Stopped.')
    await setCode(page, 'await workbook.activeSheet.range("D1").setValue("from file")')
    await page.getByRole('button', { name: 'Close', exact: true }).last().click()
    // Saved and reopened: the macros are there but disabled until the user enables them.
    const file = path.join(tempDir(), 'm.xlsx')
    await stubDialogs(app, { save: file, open: [file] })
    await page.keyboard.press(`${MOD}+s`)
    await expect(page.getByTestId('save-status')).toHaveText('Saved')
    // Forget that this session's edits were trusted, as on another computer.
    await page.evaluate(async () => {
      const sc = (window as any).__sc
      sc.set({ settings: { ...sc.get().settings, trustedMacros: [] } })
    })
    await page.keyboard.press(`${MOD}+w`)
    await page.keyboard.press(`${MOD}+o`)
    await page.getByTestId('grid').waitFor()
    await page.keyboard.press('Alt+F8')
    await expect(page.getByRole('alert')).toContainText('came from a file')
    await expect(page.getByRole('button', { name: 'Run' })).toBeDisabled()
    await page.getByRole('button', { name: 'Enable macros' }).click()
    await page.getByRole('button', { name: 'Run' }).click()
    await expect(page.getByTestId('macro-log')).toContainText('✔ Finished')
    expect(await cell(page, 'D1')).toBe('from file')
  } finally {
    await quit(app)
  }
})
