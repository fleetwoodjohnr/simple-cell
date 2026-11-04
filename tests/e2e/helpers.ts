import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, mkdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { SHORTCUTS } from '../../src/renderer/app/shortcuts'

export interface Launched {
  app: ElectronApplication
  page: Page
  userData: string
}

export function tempDir(prefix = 'sc-e2e-'): string {
  return mkdtempSync(path.join(tmpdir(), prefix))
}

export async function launch(opts: { userData?: string; settings?: Record<string, unknown>; env?: Record<string, string>; args?: string[] } = {}): Promise<Launched> {
  const userData = opts.userData ?? tempDir('sc-profile-')
  mkdirSync(userData, { recursive: true })
  if (opts.settings) writeFileSync(path.join(userData, 'settings.json'), JSON.stringify(opts.settings))
  // SC_PACKAGED=path/to/binary runs the suite against an electron-builder output instead of the dev build.
  const packaged = process.env.SC_PACKAGED
  // Update checks go to a closed local port unless a test serves a feed: tests never contact GitHub.
  const app = await electron.launch({ ...(packaged ? { executablePath: packaged, args: opts.args ?? [] } : { args: ['.', ...(opts.args ?? [])] }), env: { ...process.env, SC_UPDATE_URL: 'http://127.0.0.1:9', ...opts.env, SIMPLE_CELL_USER_DATA: userData } })
  const page = await app.firstWindow()
  const violations: string[] = []
  page.on('console', (m) => {
    if (/Content.Security.Policy/i.test(m.text())) violations.push(m.text())
    if (m.type() === 'error' && process.env.SC_DEBUG) console.log('renderer:', m.text())
  })
  page.on('pageerror', (e) => errors.get(app)?.push(e.message))
  cspViolations.set(app, violations)
  errors.set(app, [])
  await page.waitForSelector('[data-testid=dashboard], [data-testid=ribbon]')
  return { app, page, userData }
}

const cspViolations = new WeakMap<ElectronApplication, string[]>()
const errors = new WeakMap<ElectronApplication, string[]>()

/** Replace native dialogs in the main process with fixed answers. A list of save paths answers successive dialogs. */
export async function stubDialogs(app: ElectronApplication, d: { save?: string | string[]; open?: string[] }) {
  await app.evaluate(({ dialog }, d) => {
    const saves = d.save === undefined ? undefined : [d.save].flat()
    if (saves) dialog.showSaveDialog = (async () => ({ canceled: false, filePath: saves.length > 1 ? saves.shift() : saves[0] })) as never
    if (d.open !== undefined) dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: d.open })) as never
  }, d)
}

/** Every test fails on a Content-Security-Policy violation or an uncaught renderer error. */
export async function quit(app: ElectronApplication) {
  await app.evaluate(({ app }) => app.exit(0)).catch(() => {})
  const v = cspViolations.get(app)
  if (v?.length) throw new Error(`CSP violations:\n${v.join('\n')}`)
  const e = errors.get(app)
  if (e?.length) throw new Error(`Renderer errors:\n${e.join('\n')}`)
}

export async function newWorkbook(page: Page) {
  await page.getByTestId('new-doc').click()
  await page.getByTestId('grid').waitFor()
  await page.waitForTimeout(150)
}

const mac = process.platform === 'darwin'
export const MOD = mac ? 'Meta' : 'Control'

/** Value of a cell in the active workbook (A1 notation, active sheet unless "Sheet!A1"). */
export async function cell(page: Page, ref: string): Promise<unknown> {
  return page.evaluate(async (ref) => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const book = sc.books.get(tab)
    const view = sc.getView(tab)
    let sheet = view.sheet
    let a = ref
    if (ref.includes('!')) {
      const [s, r] = ref.split('!')
      sheet = book.meta.sheets.find((x: { name: string }) => x.name === s).id
      a = r
    }
    const m = /^([A-Z]+)(\d+)$/.exec(a)!
    let c = 0
    for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64
    const v = await book.api.value(sheet, Number(m[2]) - 1, c - 1)
    return v && typeof v === 'object' && 'code' in v ? v.code : v
  }, ref)
}

/** Displayed text of a cell. */
export async function shown(page: Page, ref: string): Promise<string> {
  return page.evaluate(async (ref) => {
    const sc = (window as any).__sc
    const tab = sc.get().active
    const book = sc.books.get(tab)
    const view = sc.getView(tab)
    const m = /^([A-Z]+)(\d+)$/.exec(ref)!
    let c = 0
    for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64
    const cells = await book.api.viewport(view.sheet, Number(m[2]) - 1, Number(m[2]) - 1, c - 1, c - 1)
    return cells[0]?.t ?? ''
  }, ref)
}

/** Select a cell through the Name Box, like a user typing a reference and pressing Enter. */
export async function goto(page: Page, ref: string) {
  const nb = page.getByTestId('name-box')
  await nb.click()
  await nb.fill(ref)
  await nb.press('Enter')
  await page.waitForTimeout(50)
}

/** Type into the active cell and press Enter (as a user would). */
export async function typeCell(page: Page, ref: string, text: string, key = 'Enter') {
  await goto(page, ref)
  await page.keyboard.type(text)
  await page.keyboard.press(key)
  await page.waitForTimeout(40)
}

export async function settle(page: Page, ms = 300) {
  await page.waitForTimeout(ms)
}

/** Hold one kind of worker reply until released, to exercise dialogs while their initial data is still loading. */
export async function holdWorkerReply(page: Page, method: 'noteAt' | 'filterValues') {
  await page.evaluate((method) => {
    const sc = (window as any).__sc
    const book = sc.books.get(sc.get().active)
    const api = book.api
    const gate = new Promise<void>((resolve) => {
      sc.releaseWorkerReply = () => {
        book.api = api
        resolve()
      }
    })
    book.api = new Proxy(api, {
      get: (target, key) => key === method ? async (...args: unknown[]) => {
        const result = await target[key](...args)
        await gate
        return result
      } : target[key]
    })
  }, method)
  return () => page.evaluate(() => (window as any).__sc.releaseWorkerReply())
}

/**
 * The key for a command on this OS, from the app's own shortcut registry ("Mod-Shift-k" → "Meta+Shift+k" on macOS),
 * so tests press what a user of this platform would (Mac Excel's bindings differ from Windows').
 */
export function keyFor(command: string): string {
  const def = SHORTCUTS.find((s) => s.command === command)
  if (!def) throw new Error(`No shortcut for ${command}`)
  const keys = (mac ? def.mac ?? def.win : def.win)[0]
  return keys
    .split('-')
    .map((k) => (k === 'Mod' ? MOD : k === 'Ctrl' ? 'Control' : k))
    .join('+')
}

/**
 * Click through nested menu items, moving the pointer gradually as a hand would. A submenu that opens to the left
 * (no room on the right, as on a 1024 × 768 screen) closes when the pointer jumps into it in one step.
 */
export async function clickMenuPath(page: Page, ...names: string[]) {
  for (const name of names) {
    const item = page.getByRole('menuitem', { name, exact: true })
    await item.waitFor()
    const b = (await item.boundingBox())!
    await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 12 })
    await item.click()
  }
}

/** Audit the visible UI after fonts, opening animations and color transitions have settled.
 * Infinite animations (e.g. carets) never block the audit; cancelled animations are harmless.
 */
export async function waitForUiSettled(page: Page) {
  await page.evaluate(async () => {
    await document.fonts.ready
    const animations = document.getAnimations().filter((animation) => Number.isFinite(animation.effect?.getComputedTiming().endTime))
    await Promise.all(animations.map((animation) => animation.finished.catch(() => {})))
  })
}
