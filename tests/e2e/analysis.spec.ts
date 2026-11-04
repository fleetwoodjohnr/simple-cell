import { test, expect } from '@playwright/test'
import AxeBuilder from '@axe-core/playwright'
import { launch, quit, newWorkbook, typeCell, goto, cell, clickMenuPath } from './helpers'

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

test('what-if: a data table through its dialog, and the Scenario Manager (add, show, summary)', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    // Payment-like model: B2 = 1000 × B1; rates down A5:A7, the formula in B4.
    for (const [ref, v] of [['B1', '0.05'], ['B2', '=1000*B1'], ['B4', '=B2'], ['A5', '0.01'], ['A6', '0.02'], ['A7', '0.03']] as [string, string][]) await typeCell(page, ref, v)
    await goto(page, 'A4:B7')
    await page.getByTestId('ribbon-tab-data').click()
    await page.getByRole('button', { name: 'What-if analysis' }).click()
    await clickMenuPath(page, 'Data table…')
    const dt = page.getByRole('dialog', { name: 'Data table' })
    await dt.getByLabel('Column input cell').fill('$B$1')
    await dt.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(() => cell(page, 'B6')).toBe(20)
    expect(await cell(page, 'B7')).toBe(30)
    // Part of a data table can't be changed.
    await typeCell(page, 'B6', '5')
    await expect(page.getByRole('dialog', { name: /Can’t change that/ })).toBeVisible()
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    expect(await cell(page, 'B6')).toBe(20)
    // The table follows its formula.
    await typeCell(page, 'B2', '=2000*B1')
    await expect.poll(() => cell(page, 'B6')).toBe(40)

    // Scenarios on B1.
    await goto(page, 'B1')
    await page.getByRole('button', { name: 'What-if analysis' }).click()
    await clickMenuPath(page, 'Scenario manager…')
    const sm = page.getByRole('dialog', { name: 'Scenario manager' })
    for (const [name, value] of [['High', '0.08'], ['Low', '0.01']]) {
      await page.getByRole('button', { name: 'Add…' }).click()
      const add = page.getByRole('dialog', { name: 'Add scenario' })
      await add.getByLabel('Scenario name').fill(name)
      await expect(add.getByLabel('Changing cells')).toHaveValue('$B$1')
      await add.getByRole('button', { name: 'Next' }).click()
      await add.getByLabel('Value for B1').fill(value)
      await add.getByRole('button', { name: 'OK', exact: true }).click()
    }
    await expect(sm.getByRole('option')).toHaveText(['High', 'Low'])
    await sm.getByRole('listbox').selectOption({ label: 'High' })
    await sm.getByRole('button', { name: 'Show' }).click()
    await expect.poll(() => cell(page, 'B2')).toBe(160)
    await sm.getByRole('button', { name: 'Summary…' }).click()
    await page.getByRole('dialog', { name: 'Scenario summary' }).getByLabel('Result cells').fill('$B$2')
    await page.getByRole('dialog', { name: 'Scenario summary' }).getByRole('button', { name: 'OK', exact: true }).click()
    await expect(page.getByTestId('sheet-tab-Scenario Summary')).toBeVisible()
    expect(await cell(page, 'Scenario Summary!E6')).toBe(160)
    expect(await cell(page, 'Scenario Summary!F6')).toBe(20)
  } finally {
    await quit(app)
  }
})

test('Data Analysis: Regression through its dialog onto a new sheet (and the form passes axe)', async () => {
  const { app, page } = await launch()
  try {
    await newWorkbook(page)
    const data: [string, string][] = [['A1', 'Y'], ['B1', 'X']]
    ;[3, 5, 7, 10, 11, 15].forEach((y, i) => data.push([`A${i + 2}`, String(y)], [`B${i + 2}`, String(i + 1)]))
    for (const [ref, v] of data) await typeCell(page, ref, v)
    await goto(page, 'A1')
    await page.getByTestId('ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Data analysis', exact: true }).click()
    await page.getByRole('listbox', { name: 'Analysis tools' }).selectOption({ label: 'Regression' })
    await page.getByRole('button', { name: 'OK', exact: true }).click()
    const form = page.getByRole('dialog', { name: 'Regression' })
    await expect(form.getByLabel('Input Y range')).toHaveValue('$A$1:$B$7')
    await form.getByLabel('Input Y range').fill('$A$1:$A$7')
    await form.getByLabel('Input X range').fill('$B$1:$B$7')
    await form.getByLabel('Residuals').check()
    const a11y = await new AxeBuilder({ page }).setLegacyMode(true).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze()
    expect(a11y.violations.map((v) => v.id)).toEqual([])
    await form.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(page.getByTestId('sheet-tab-Regression')).toBeVisible()
    expect(await cell(page, 'Regression!A1')).toBe('SUMMARY OUTPUT')
    // Slope of y on x: 2.3143… (least squares on these six points: Σ(x−x̄)(y−ȳ) / Σ(x−x̄)² = 40.5 / 17.5).
    expect(await cell(page, 'Regression!B18')).toBeCloseTo(2.3142857142857145, 12)
    expect(await cell(page, 'Regression!A18')).toBe('X')
  } finally {
    await quit(app)
  }
})
