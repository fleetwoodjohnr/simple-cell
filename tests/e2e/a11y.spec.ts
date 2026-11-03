// Accessibility: axe-core WCAG 2.2 A/AA rules on the dashboard, every ribbon tab and the main dialogs, light and dark.
import { test, expect, type Page } from '@playwright/test'
import AxeBuilder from '@axe-core/playwright'
import { launch, quit, newWorkbook, typeCell } from './helpers'

const TAGS = ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']

async function violations(page: Page, label: string) {
  const r = await new AxeBuilder({ page }).setLegacyMode(true).withTags(TAGS).analyze()
  return r.violations.map((v) => `${label}: [${v.impact}] ${v.id} — ${v.help} (${v.nodes.length}×: ${v.nodes.slice(0, 3).map((n) => n.target.join(' ')).join(' | ')})`)
}

const command = (page: Page, id: string) => page.evaluate((id) => (window as unknown as { __sc: { runCommand: (id: string) => void } }).__sc.runCommand(id), id)

for (const theme of ['light', 'dark'] as const)
  test(`accessibility (axe, WCAG 2.2 AA): main screens, ${theme}`, async () => {
    const { app, page } = await launch({ settings: { theme } })
    const found: string[] = []
    try {
      found.push(...(await violations(page, 'dashboard')))
      await newWorkbook(page)
      await typeCell(page, 'A1', 'Hello')
      for (const tab of ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'View', 'Automate', 'Help']) {
        await page.getByRole('tablist', { name: 'Ribbon' }).getByRole('tab', { name: tab, exact: true }).click()
        found.push(...(await violations(page, `ribbon ${tab}`)))
      }
      for (const [id, name] of [['fmt.cells', 'Format cells'], ['tools.settings', 'Settings'], ['tools.macros', 'Macros']]) {
        await command(page, id)
        await expect(page.getByRole('dialog')).toBeVisible()
        found.push(...(await violations(page, `dialog ${name}`)))
        await page.keyboard.press('Escape')
        await expect(page.getByRole('dialog')).toHaveCount(0)
      }
    } finally {
      await quit(app)
    }
    expect(found).toEqual([])
  })
