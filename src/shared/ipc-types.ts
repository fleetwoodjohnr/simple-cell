// Contract between the main process and the renderer. The preload script exposes
// `window.simpleCell` implementing `SimpleCellApi`; main validates every argument.

export interface OpenedFile {
  path: string
  name: string
  data: Uint8Array
  mtime: number
}

export interface RecentFile {
  path: string
  name: string
  openedAt: number
  pinned?: boolean
  favorite?: boolean
  thumbnail?: string
}

export interface RecoveryEntry {
  id: string
  savedAt: number
  /** JSON produced by the renderer (workbook snapshot + tab metadata). */
  payload: string
}

export type ThemeSetting = 'system' | 'light' | 'dark'

export interface Settings {
  theme: ThemeSetting
  accent: string
  defaultFont: string
  defaultFontSize: number
  defaultSaveFormat: 'xlsx' | 'ods' | 'csv'
  autosave: boolean
  autosaveIntervalSec: number
  recoveryIntervalSec: number
  authorName: string
  /** SHA-256 hashes of macro sets the user enabled (a changed macro asks again). */
  trustedMacros: string[]
  ribbonCollapsed: boolean
  /** Ribbon layout: the full Excel-style ribbon, or one compact row. */
  ribbonMode: 'expanded' | 'compact'
  reduceMotion: boolean
  uiScale: number
  /** Mouse pointer over cells: the system arrow, or Excel's fat white cross. */
  cellCursor: 'arrow' | 'cross'
  /** Keyboard preset the shortcuts start from; `shortcuts` overrides individual commands. */
  shortcutPreset: 'excel' | 'sheets'
  shortcuts: Record<string, string>
  quickAccess: string[]
  /** CSV export prefixes cells that start with = + - @ so other apps don't run them as formulas. */
  csvFormulaGuard: boolean
  /** Ask GitHub for a newer version at startup (at most daily). */
  updateCheck: boolean
  /** When the automatic update check last ran (ms since epoch). */
  lastUpdateCheck: number
}

export const DEFAULT_SETTINGS: Settings = {
  theme: 'system',
  accent: '#1b7f53',
  defaultFont: 'Calibri',
  defaultFontSize: 11,
  defaultSaveFormat: 'xlsx',
  autosave: true,
  autosaveIntervalSec: 30,
  recoveryIntervalSec: 5,
  authorName: '',
  trustedMacros: [],
  ribbonCollapsed: false,
  ribbonMode: 'expanded',
  reduceMotion: false,
  uiScale: 1,
  cellCursor: 'arrow',
  shortcutPreset: 'excel',
  shortcuts: {},
  quickAccess: ['file.save', 'edit.undo', 'edit.redo'],
  csvFormulaGuard: true,
  updateCheck: true,
  lastUpdateCheck: 0
}

export interface PdfOptions {
  /** Page size and orientation come from the printed HTML's @page rule. */
  title?: string
}

export interface SimpleCellApi {
  platform: 'darwin' | 'win32' | 'linux' | string
  /** Native open dialog; `filter` replaces the default workbook filter (single selection). */
  openDialog(filter?: { name: string; extensions: string[] }): Promise<OpenedFile[]>
  readFile(path: string): Promise<OpenedFile>
  saveDialog(defaultName: string, extensions: string[]): Promise<string | null>
  writeFile(path: string, data: Uint8Array): Promise<{ mtime: number }>
  /** Resolve a dropped File to its path and grant access to it. */
  grantDroppedFile(file: File): string | null

  recentList(): Promise<RecentFile[]>
  recentAdd(entry: RecentFile): Promise<RecentFile[]>
  recentUpdate(path: string, patch: Partial<RecentFile>): Promise<RecentFile[]>
  recentRemove(path: string): Promise<RecentFile[]>

  recoveryWrite(id: string, payload: string): Promise<void>
  recoveryList(): Promise<RecoveryEntry[]>
  recoveryRemove(id: string): Promise<void>

  /** Version history, keyed by workbook path. `auto` versions are kept at most every 10 minutes. */
  historyAdd(docPath: string, json: string, auto: boolean): Promise<boolean>
  historyList(docPath: string): Promise<{ id: string; savedAt: number; bytes: number }[]>
  historyRead(docPath: string, id: string): Promise<string>

  settingsGet(): Promise<Settings>
  settingsSet(patch: Partial<Settings>): Promise<Settings>

  exportPdf(options: PdfOptions): Promise<Uint8Array>
  print(): Promise<boolean>
  /** A 256 px wide PNG of that part of the window, or null when nothing could be captured. */
  capturePage(rect: { x: number; y: number; width: number; height: number }): Promise<Uint8Array | null>

  libreOfficeAvailable(): Promise<boolean>
  libreOfficeConvert(data: Uint8Array, fromExt: string, toExt: string): Promise<Uint8Array>

  windowMinimize(): void
  windowToggleMaximize(): void
  windowClose(): void
  setTitleBarColors(colors: { color: string; symbolColor: string }): void
  openExternal(url: string): Promise<void>

  onOpenFiles(cb: (files: OpenedFile[]) => void): () => void
  onCloseRequest(cb: () => void): () => void
  onMenuCommand(cb: (id: string) => void): () => void
  confirmClose(): void
  /** Check GitHub for a newer version now; the main process shows the outcome in native dialogs. */
  checkForUpdates(): Promise<void>
  appVersion(): Promise<string>
}
