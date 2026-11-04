import { test, expect, type Page } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, cell, MOD, keyFor, clickMenuPath } from './helpers'

const style = (page: Page, r: number, c: number) =>
  page.evaluate(([r, c]) => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    return sc.books.get(tab).api.styleAt(sc.getView(tab).sheet, r, c)
  }, [r, c])

test('paste special (operation, transpose) and format painter', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await typeCell(page, 'A1', '10')
    await typeCell(page, 'A2', '20')
    await typeCell(page, 'C1', '3')
    // Copy C1, Paste Special ▸ Multiply onto A1:A2.
    await goto(page, 'C1')
    await page.keyboard.press(`${MOD}+c`)
    await goto(page, 'A1:A2')
    await page.keyboard.press(keyFor('edit.pasteSpecial'))
    await page.getByLabel('Values', { exact: true }).check()
    await page.getByLabel('Multiply').check()
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => cell(page, 'A2')).toBe(60)
    expect(await cell(page, 'A1')).toBe(30)
    // Transpose the column into a row.
    await goto(page, 'A1:A2')
    await page.keyboard.press(`${MOD}+c`)
    await goto(page, 'E1')
    await page.keyboard.press(keyFor('edit.pasteSpecial'))
    await page.getByLabel('Transpose').check()
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => cell(page, 'F1')).toBe(60)
    expect(await cell(page, 'E2')).toBeNull()
    // Format painter: bold red A1 → click the brush → select C1:C2.
    await goto(page, 'A1')
    await page.keyboard.press(`${MOD}+b`)
    await page.getByTestId('format-painter').click()
    await expect(page.getByTestId('format-painter')).toHaveAttribute('aria-pressed', 'true')
    const box = await page.getByTestId('grid').boundingBox()
    const at = (col: number, row: number) => ({ x: box!.x + 40 + 64 * col + 30, y: box!.y + 26 + 20 * row + 10 })
    const p1 = at(2, 0)
    const p2 = at(2, 1)
    await page.mouse.move(p1.x, p1.y)
    await page.mouse.down()
    await page.mouse.move(p2.x, p2.y, { steps: 3 })
    await page.mouse.up()
    await expect.poll(async () => (await style(page, 1, 2)).bold).toBe(true)
    expect((await style(page, 0, 2)).bold).toBe(true)
    await expect(page.getByTestId('format-painter')).toHaveAttribute('aria-pressed', 'false')
    // Cell styles: Good sets its fill and font colour; Normal clears formatting again.
    await goto(page, 'A2')
    await page.getByRole('button', { name: 'Cell styles' }).click()
    await page.getByRole('button', { name: 'Good', exact: true }).click()
    await expect.poll(async () => (await style(page, 1, 0)).fill).toBe('#C6EFCE')
    await page.getByRole('button', { name: 'Cell styles' }).click()
    await page.getByRole('button', { name: 'Normal', exact: true }).click()
    await expect.poll(async () => (await style(page, 1, 0)).fill).toBeUndefined()
    // Flash Fill: one example, Ctrl+E fills the rest of the column.
    for (const [i, n] of ['Ann Lee', 'Bo Zhang', 'Cy Ray'].entries()) await typeCell(page, `H${i + 1}`, n)
    await typeCell(page, 'I1', 'Lee')
    await goto(page, 'I2')
    await page.keyboard.press(`${MOD}+e`)
    await expect.poll(() => cell(page, 'I3')).toBe('Ray')
    expect(await cell(page, 'I2')).toBe('Zhang')
    // Ctrl+[ selects the cells a formula uses; Ctrl+Shift+U makes the formula bar taller.
    await typeCell(page, 'K1', '=A1+C1')
    await goto(page, 'K1')
    await page.keyboard.press(`${MOD}+BracketLeft`)
    await expect(page.getByTestId('name-box')).toHaveValue('A1')
    await page.keyboard.press(keyFor('view.formulaBarTall'))
    await expect.poll(async () => (await page.getByTestId('formula-bar').boundingBox())!.height).toBeGreaterThan(80)
  } finally {
    await quit(app)
  }
})

test('AutoFit column width and row height (Home ▸ Format, double-click a row border), one undo step each', async () => {
  const { app, page } = await launch()
  const meta = () =>
    page.evaluate(() => {
      const sc = (window as any).__sc
      const tab = sc.get().active
      const m = sc.books.get(tab).sheetMeta(sc.getView(tab).sheet)
      return { cols: Object.fromEntries(m.cols), rows: Object.fromEntries(m.rows), defW: m.defaultColWidth, defH: m.defaultRowHeight }
    })
  const fontSize = (size: number | undefined) => page.evaluate((size) => (window as any).__sc.A.formatSel((window as any).__sc.A.ctx(), { size }), size)
  try {
    await newWorkbook(page)
    await typeCell(page, 'B2', 'A fairly long piece of text that overflows column B')
    await goto(page, 'B2')
    await page.getByRole('tablist', { name: 'Ribbon' }).getByRole('tab', { name: 'Home', exact: true }).click()
    await page.getByRole('button', { name: 'Format', exact: true }).click()
    await clickMenuPath(page, 'AutoFit column width')
    const { defW, defH } = await meta()
    await expect.poll(async () => (await meta()).cols[1] ?? 0).toBeGreaterThan(defW * 2)

    await typeCell(page, 'C3', 'Big')
    await goto(page, 'C3')
    await fontSize(28)
    await page.evaluate(() => (window as any).__sc.runCommand('fmt.autofitRows'))
    await expect.poll(async () => (await meta()).rows[2] ?? 0).toBeGreaterThan(defH * 2)
    // Back to the default font: the row goes back to the default height. Undo restores the fitted height in one step.
    await fontSize(undefined)
    await page.evaluate(() => (window as any).__sc.runCommand('fmt.autofitRows'))
    await expect.poll(async () => (await meta()).rows[2]).toBeUndefined()
    await page.keyboard.press(`${MOD}+z`)
    await expect.poll(async () => (await meta()).rows[2] ?? 0).toBeGreaterThan(defH * 2)
    // Double-click the border under row 3's header fits it too: the font is the default again, so the row shrinks back.
    const box = (await page.getByTestId('grid').boundingBox())!
    await page.mouse.dblclick(box.x + 10, box.y + 22 + 2 * defH + (await meta()).rows[2])
    await expect.poll(async () => (await meta()).rows[2]).toBeUndefined()
    // Wrapped text: one line per wrapped line.
    await typeCell(page, 'C5', 'one two three four five six seven eight')
    await goto(page, 'C5')
    await page.evaluate(() => (window as any).__sc.A.formatSel((window as any).__sc.A.ctx(), { wrap: true }))
    await page.evaluate(() => (window as any).__sc.runCommand('fmt.autofitRows'))
    await expect.poll(async () => (await meta()).rows[4] ?? 0).toBeGreaterThan(defH * 3)
  } finally {
    await quit(app)
  }
})
