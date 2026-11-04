import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, cell, MOD, holdWorkerReply } from './helpers'

for (const text of ['Typed while the note loads', '']) {
  test(`a delayed note lookup preserves ${text ? 'new text' : 'a deliberate deletion'}`, async () => {
    const { app, page } = await launch()
    try {
      await newWorkbook(page)
      await page.evaluate(async () => {
        const sc = (window as any).__sc
        const tab = sc.get().active
        await sc.books.get(tab).api.setNote(sc.getView(tab).sheet, 0, 0, 'Original note', 'Tester')
      })
      const release = await holdWorkerReply(page, 'noteAt')
      await page.keyboard.press('Shift+F2')
      const note = page.getByLabel('Note', { exact: true })
      await note.fill('Pending edit')
      await note.fill(text)
      const save = page.getByRole('button', { name: 'Save', exact: true })
      await expect(save).toBeDisabled()
      await release()
      await expect(save).toBeEnabled()
      await expect(note).toHaveValue(text)
      await save.click()
      await expect(page.getByRole('dialog')).toHaveCount(0)
      expect(await page.evaluate(async () => {
        const sc = (window as any).__sc
        const tab = sc.get().active
        return (await sc.books.get(tab).api.noteAt(sc.getView(tab).sheet, 0, 0))?.text ?? null
      })).toBe(text || null)
    } finally {
      await quit(app)
    }
  })
}

test('notes (Shift+F2, hover) and links (Ctrl+K, chip, open, internal jump)', async () => {
  const { app, page } = await launch()
  try {
    await app.evaluate(({ shell }) => {
      ;(globalThis as any).__opened = []
      shell.openExternal = (async (u: string) => void (globalThis as any).__opened.push(u)) as never
    })
    await newWorkbook(page)
    await typeCell(page, 'B2', 'Revenue')
    await goto(page, 'B2')
    await page.keyboard.press('Shift+F2')
    await page.getByLabel('Note', { exact: true }).fill('Includes the March correction')
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    // Hovering the cell shows the note (once the dialog is gone: until then the pointer is over its overlay).
    await expect(page.getByRole('dialog')).toHaveCount(0)
    const box = await page.getByTestId('grid').boundingBox()
    await page.mouse.move(box!.x + 46 + 64 + 30, box!.y + 22 + 20 + 10)
    await expect(page.getByTestId('note-popup')).toContainText('Includes the March correction')
    await page.mouse.move(box!.x + 400, box!.y + 300)
    await expect(page.getByTestId('note-popup')).toHaveCount(0)
    // Ctrl+K on an empty cell: the address becomes the text.
    await goto(page, 'A5')
    await page.keyboard.press(`${MOD}+k`)
    await page.getByLabel('Address').fill('example.com/docs')
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => cell(page, 'A5')).toBe('https://example.com/docs')
    await expect(page.getByTestId('link-chip')).toContainText('https://example.com/docs')
    await page.getByRole('button', { name: 'Open link' }).click()
    await expect.poll(() => app.evaluate(() => (globalThis as any).__opened)).toEqual(['https://example.com/docs'])
    // An internal link jumps to its place.
    await goto(page, 'A6')
    await page.keyboard.press(`${MOD}+k`)
    await page.getByLabel('Text to display').fill('Go to totals')
    await page.getByLabel('Address').fill('Sheet1!D20')
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    expect(await cell(page, 'A6')).toBe('Go to totals')
    await page.getByTestId('link-chip').getByRole('button', { name: 'Sheet1!D20' }).click()
    await expect(page.getByTestId('name-box')).toHaveValue('D20')
    // Remove the link from the chip; undo brings it back.
    await goto(page, 'A5')
    await page.getByRole('button', { name: 'Remove link' }).click()
    await expect(page.getByTestId('link-chip')).toHaveCount(0)
    await page.keyboard.press(`${MOD}+z`)
    await expect(page.getByTestId('link-chip')).toBeVisible()
  } finally {
    await quit(app)
  }
})
