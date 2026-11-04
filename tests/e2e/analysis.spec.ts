import { test, expect } from '@playwright/test'
import { launch, quit, newWorkbook, typeCell, goto, cell } from './helpers'

test('Solver: objective, variables, a constraint; Solve; keep the solution with an Answer Report; the model is remembered', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    // Wyndor Glass: profit 3x + 5y; x ≤ 4, 2y ≤ 12, 3x + 2y ≤ 18.
    const cells: [string, string][] = [['A1', 'Doors'], ['A2', 'Windows'], ['B1', '0'], ['B2', '0'], ['C1', '=B1'], ['C2', '=2*B2'], ['C3', '=3*B1+2*B2'], ['E1', '4'], ['E2', '12'], ['E3', '18'], ['D1', '=3*B1+5*B2']]
    for (const [ref, v] of cells) await typeCell(page, ref, v)
    await goto(page, 'D1')
    await page.getByTestId('ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Solver', exact: true }).click()
    const dlg = page.getByRole('dialog', { name: 'Solver' })
    await expect(dlg.getByLabel('Set objective')).toHaveValue('$D$1')
    await dlg.getByLabel('By changing variable cells').fill('$B$1:$B$2')
    await dlg.getByRole('button', { name: 'Add', exact: true }).click()
    await dlg.getByLabel('Cell reference').fill('$C$1:$C$3')
    await dlg.getByLabel('Constraint', { exact: true }).fill('$E$1:$E$3')
    await dlg.getByLabel('Constraint', { exact: true }).press('Enter')
    await expect(dlg.getByRole('option', { name: '$C$1:$C$3 <= $E$1:$E$3' })).toBeAttached()
    await dlg.getByRole('button', { name: 'Solve', exact: true }).click()
    const res = page.getByRole('dialog', { name: 'Solver results' })
    await expect(res.getByRole('status')).toContainText('Solver found a solution')
    await res.getByLabel('Answer report (a new sheet)').check()
    await res.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => cell(page, 'Sheet1!D1')).toBe(36)
    expect(await cell(page, 'Sheet1!B1')).toBe(2)
    expect(await cell(page, 'Sheet1!B2')).toBe(6)
    await expect(page.getByTestId('sheet-tab-Answer Report 1')).toBeVisible()
    // The model is saved on the sheet: the dialog opens with it.
    await page.getByTestId('sheet-tab-Sheet1').click()
    await page.getByRole('button', { name: 'Solver', exact: true }).click()
    await expect(dlg.getByLabel('By changing variable cells')).toHaveValue('$B$1:$B$2')
    await expect(dlg.getByRole('option', { name: '$C$1:$C$3 <= $E$1:$E$3' })).toBeAttached()
    // An infeasible model says so and keeps the values.
    await dlg.getByRole('button', { name: 'Add', exact: true }).click()
    await dlg.getByLabel('Cell reference').fill('$B$1')
    await dlg.getByLabel('Relation').selectOption('>=')
    await dlg.getByLabel('Constraint', { exact: true }).fill('5')
    await dlg.getByRole('button', { name: 'Add', exact: true }).last().click()
    await dlg.getByRole('button', { name: 'Solve', exact: true }).click()
    await expect(res.getByRole('status')).toContainText('could not find a feasible solution')
    await res.getByRole('button', { name: 'OK', exact: true }).click()
    expect(await cell(page, 'Sheet1!B1')).toBe(2)
  } finally {
    await quit(app)
  }
})
