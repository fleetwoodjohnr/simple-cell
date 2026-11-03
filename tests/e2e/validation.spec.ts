import { test, expect, type Page } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, cell } from './helpers'

async function validationMenu(page: Page, item: string) {
  await page.getByTestId('ribbon-tab-data').click()
  await page.getByTestId('ribbon').getByRole('button', { name: 'Data validation', exact: true }).click()
  await page.getByRole('menuitem', { name: item }).click()
}

test('data validation: list dropdown, input message, stop alert, circle invalid data', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    await goto(page, 'A1:A5')
    await validationMenu(page, 'Data validation…')
    await page.getByLabel('Allow').selectOption('list')
    await page.getByLabel('Source').fill('Red, Green, Blue')
    await page.getByRole('radio', { name: 'Input message' }).click()
    await page.getByLabel('Input title').fill('Colour')
    await page.getByRole('textbox', { name: 'Input message' }).fill('Pick one from the list')
    await page.getByRole('button', { name: 'OK' }).click()
    await goto(page, 'A1')
    await expect(page.getByTestId('input-message')).toContainText('Pick one from the list')
    // Mouse: the arrow next to the cell opens the list.
    await page.getByTestId('list-arrow').click()
    await page.getByRole('option', { name: 'Green' }).click()
    await expect.poll(() => cell(page, 'A1')).toBe('Green')
    // Keyboard: Alt+Down, arrow, Enter.
    await goto(page, 'A2')
    await page.keyboard.press('Alt+ArrowDown')
    await expect(page.getByRole('listbox', { name: 'Allowed values' })).toBeVisible()
    await page.keyboard.press('ArrowDown')
    await page.keyboard.press('ArrowDown')
    await page.keyboard.press('Enter')
    await expect.poll(() => cell(page, 'A2')).toBe('Blue')
    // A value outside the list is refused (stop alert); Cancel leaves the cell as it was.
    await typeCell(page, 'A3', 'Purple')
    await expect(page.getByRole('alert')).toContainText('doesn’t match the data validation restrictions')
    await page.getByRole('button', { name: 'Cancel' }).click()
    expect(await cell(page, 'A3')).toBeNull()
    await typeCell(page, 'A3', 'blue')
    expect(await cell(page, 'A3')).toBe('blue')
    // Circle invalid data: a rule added over existing values marks the ones that break it.
    await typeCell(page, 'B1', '5')
    await typeCell(page, 'B2', '-2')
    await goto(page, 'B1:B2')
    await validationMenu(page, 'Data validation…')
    await page.getByLabel('Allow').selectOption('whole')
    await page.getByRole('combobox', { name: 'Data', exact: true }).selectOption('greaterThan')
    await page.getByLabel('Value').fill('0')
    await page.getByRole('button', { name: 'OK' }).click()
    await validationMenu(page, 'Circle invalid data')
    await expect(page.getByTestId('invalid-circle')).toHaveCount(1)
    await typeCell(page, 'B2', '3')
    await expect(page.getByTestId('invalid-circle')).toHaveCount(0)
  } finally {
    await quit(app)
  }
})
