// Accessibility: axe-core WCAG 2.2 A/AA rules on the dashboard, every ribbon tab and every dialog, light and dark.
import { test, expect, type Page } from '@playwright/test'
import AxeBuilder from '@axe-core/playwright'
import { launch, quit, newWorkbook, typeCell, goto, cell } from './helpers'

const TAGS = ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']

async function violations(page: Page, label: string) {
  const r = await new AxeBuilder({ page }).setLegacyMode(true).withTags(TAGS).analyze()
  return r.violations.map((v) => `${label}: [${v.impact}] ${v.id} — ${v.help} (${v.nodes.length}×: ${v.nodes.slice(0, 3).map((n) => n.target.join(' ')).join(' | ')})`)
}

const AREA = { r0: 0, c0: 0, r1: 2, c1: 1 }

// Every dialog, opened through the app's store. '$tab', '$tabs', '$chart' and '$fn' are filled in inside the page.
const DIALOGS: [string, Record<string, unknown>][] = [
  ...['general', 'editing', 'keyboard', 'updates'].map((section): [string, Record<string, unknown>] => [`Settings ${section}`, { kind: 'settings', section }]),
  ...['number', 'alignment', 'font', 'border', 'fill', 'protection'].map((tab): [string, Record<string, unknown>] => [`Format cells ${tab}`, { kind: 'format-cells', tab }]),
  ['Macros', { kind: 'macros' }],
  ['SQL query', { kind: 'sql' }],
  ['Close confirmation', { kind: 'confirm-close', tabIds: '$tabs' }],
  ['Compatibility report', { kind: 'compat', tabId: '$tab' }],
  ['Keyboard shortcuts', { kind: 'shortcuts' }],
  ['About', { kind: 'about' }],
  ['Command palette', { kind: 'palette' }],
  ['Go to', { kind: 'goto' }],
  ['Find', { kind: 'find', replace: false }],
  ['Replace', { kind: 'find', replace: true }],
  ['Name manager', { kind: 'names' }],
  ['Insert function', { kind: 'insert-function' }],
  ['Sort', { kind: 'sort' }],
  ...['greaterThan', 'between', 'text', 'date', 'duplicate', 'top', 'formula'].map((preset): [string, Record<string, unknown>] => [`Conditional format ${preset}`, { kind: 'cf-rule', preset }]),
  ['Conditional formatting rules', { kind: 'cf-manage' }],
  ['Data validation', { kind: 'data-validation' }],
  ['Validation alert', { kind: 'dv-alert', alert: { style: 'stop', title: 'Not allowed', message: 'Pick a value from the list.' }, resolve: '$fn' }],
  ['Create table', { kind: 'create-table', area: AREA, headers: true }],
  ['Note', { kind: 'note' }],
  ['Insert cells', { kind: 'shift-cells', mode: 'insert' }],
  ['Delete cells', { kind: 'shift-cells', mode: 'delete' }],
  ['Paste special', { kind: 'paste-special' }],
  ['Go to special', { kind: 'goto-special' }],
  ['Print', { kind: 'print' }],
  ['Chart', { kind: 'chart', id: '$chart' }],
  ['PivotTable', { kind: 'pivot-insert', area: AREA }],
  ...['dedupe', 'text-to-columns', 'goal-seek', 'descriptive', 'solver', 'data-table', 'scenarios', 'data-analysis'].map((tool): [string, Record<string, unknown>] => [`Tool ${tool}`, { kind: 'tool', tool }]),
  ['Sparklines', { kind: 'sparklines', type: 'line' }],
  ['Link', { kind: 'link' }],
  ['Row height', { kind: 'row-height' }],
  ['Column width', { kind: 'col-width' }],
  ['Export warnings', { kind: 'export-warnings', tabId: '$tab', format: 'csv', warnings: [{ level: 'loss', message: 'Formatting is not kept.' }], resolve: '$fn' }],
  ['Message', { kind: 'message', title: 'Done', body: 'Something happened.' }]
]

function openDialog(page: Page, d: Record<string, unknown>) {
  return page.evaluate((d) => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const fill: Record<string, unknown> = { $tab: tab, $tabs: [tab], $chart: sc.getView(tab).chart, $fn: () => {} }
    sc.set({ dialog: Object.fromEntries(Object.entries(d).map(([k, v]) => [k, typeof v === 'string' && v in fill ? fill[v] : v])) })
  }, d)
}

for (const theme of ['light', 'dark'] as const)
  test(`accessibility (axe, WCAG 2.2 AA): main screens and every dialog, ${theme}`, async () => {
    const { app, page } = await launch({ settings: { theme } })
    const found: string[] = []
    let current = 'dashboard'
    page.on('pageerror', (e) => found.push(`${current}: renderer error ${e.message.slice(0, 80)}`))
    try {
      found.push(...(await violations(page, 'dashboard')))
      await newWorkbook(page)
      for (const [ref, v] of [['A1', 'Name'], ['B1', 'Value'], ['A2', 'a'], ['B2', '1'], ['A3', 'b'], ['B3', '2']]) await typeCell(page, ref, v)
      for (const tab of ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'View', 'Automate', 'Help']) {
        await page.getByRole('tablist', { name: 'Ribbon' }).getByRole('tab', { name: tab, exact: true }).click()
        found.push(...(await violations(page, `ribbon ${tab}`)))
      }
      current = 'insert chart'
      await goto(page, 'A1:B3')
      await page.evaluate(() => (window as any).__sc.runCommand('insert.chart'))
      await expect.poll(() => page.evaluate(() => (window as any).__sc.getView((window as any).__sc.get().active).chart)).toBeTruthy()
      found.push(...(await violations(page, 'grid with a chart')))
      const dialog = page.locator('[role=dialog], [role=alertdialog]')
      for (const [name, d] of DIALOGS) {
        current = `dialog ${name}`
        await goto(page, 'A1:B3')
        await openDialog(page, d)
        await expect(dialog.first(), name).toBeVisible()
        found.push(...(await violations(page, `dialog ${name}`)))
        await page.keyboard.press('Escape')
        await expect(dialog, `${name} closes with Escape`).toHaveCount(0)
      }
      // Side panels: PivotTable Fields, with fields in its areas.
      current = 'PivotTable Fields'
      await goto(page, 'A1:B3')
      await page.getByTestId('ribbon-tab-insert').click()
      await page.getByRole('button', { name: 'PivotTable' }).click()
      await page.getByRole('button', { name: 'OK', exact: true }).click()
      const panel = page.getByTestId('pivot-panel')
      await panel.getByLabel('Name', { exact: true }).click()
      await panel.getByLabel('Value', { exact: true }).click()
      await expect.poll(() => cell(page, 'B4')).toBe(1)
      found.push(...(await violations(page, 'PivotTable Fields panel')))
    } finally {
      await quit(app).catch((e) => found.push(String(e).slice(0, 200)))
    }
    expect(found).toEqual([])
  })
