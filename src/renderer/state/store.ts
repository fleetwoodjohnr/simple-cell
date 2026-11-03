import type { CfPreset } from '../app/CondFormat'
import type { DvAlert } from '@/core/validation'
import type { AlertChoice } from '../app/Validation'
import { create } from 'zustand'
import { DEFAULT_SETTINGS, type RecentFile, type RecoveryEntry, type Settings } from '@/shared/ipc-types'
import type { FormatId, FormatWarning } from '@/core/formats/types'
import type { Book } from './client'

export type RibbonTab = 'home' | 'insert' | 'page' | 'formulas' | 'data' | 'review' | 'view' | 'automate' | 'help' | 'table'
export type SaveStatus = 'saved' | 'unsaved' | 'saving' | 'error'

export interface TabMeta {
  id: string
  title: string
  path: string | null
  format: FormatId | null
  dirty: boolean
  status: SaveStatus
  warnings: FormatWarning[]
  /** Autosave to the file is paused until the user saves explicitly (e.g. after a lossy import). */
  autosaveBlocked: boolean
  savedAt: number | null
  error?: string
}

export type Dialog =
  | { kind: 'settings'; section?: string }
  | { kind: 'confirm-close'; tabIds: string[]; then?: 'quit' }
  | { kind: 'compat'; tabId: string }
  | { kind: 'shortcuts' }
  | { kind: 'about' }
  | { kind: 'palette' }
  | { kind: 'goto' }
  | { kind: 'find'; replace: boolean }
  | { kind: 'names' }
  | { kind: 'insert-function'; initial?: string }
  | { kind: 'format-cells'; tab?: string }
  | { kind: 'sort' }
  | { kind: 'cf-rule'; preset: CfPreset }
  | { kind: 'cf-manage' }
  | { kind: 'data-validation' }
  | { kind: 'note' }
  | { kind: 'paste-special' }
  | { kind: 'goto-special' }
  | { kind: 'print' }
  | { kind: 'chart'; id: number }
  | { kind: 'shift-cells'; mode: 'insert' | 'delete' }
  | { kind: 'link' }
  | { kind: 'create-table'; area: { r0: number; c0: number; r1: number; c1: number }; headers: boolean; style?: string }
  | { kind: 'dv-alert'; alert: DvAlert; resolve: (c: AlertChoice) => void }
  | { kind: 'row-height' | 'col-width' }
  | { kind: 'csv-import'; name: string; path: string; data: Uint8Array }
  | { kind: 'history' }
  | { kind: 'export-warnings'; tabId: string; format: FormatId; warnings: FormatWarning[]; resolve: (choice: 'save' | 'xlsx' | 'cancel') => void }
  | { kind: 'message'; title: string; body: string }

export interface Toast {
  id: number
  kind: 'info' | 'success' | 'error'
  message: string
}

export interface AppState {
  ready: boolean
  settings: Settings
  recents: RecentFile[]
  recovery: RecoveryEntry[]
  tabs: TabMeta[]
  active: string
  backstage: boolean
  ribbonTab: RibbonTab
  dialog: Dialog | null
  toasts: Toast[]
  libreOffice: boolean
  dark: boolean
  /** View → Show formulas (Ctrl+`). */
  showFormulas: boolean
  showHeadings: boolean
  showFormulaBar: boolean
  /** The formula bar shows several lines (Ctrl+Shift+U). */
  formulaBarTall: boolean
}

export const HOME = 'home'

export const useApp = create<AppState>(() => ({
  ready: false,
  settings: DEFAULT_SETTINGS,
  recents: [],
  recovery: [],
  tabs: [],
  active: HOME,
  backstage: false,
  ribbonTab: 'home',
  dialog: null,
  toasts: [],
  libreOffice: false,
  dark: false,
  showFormulas: false,
  showHeadings: true,
  showFormulaBar: true,
  formulaBarTall: false
}))

export const set = useApp.setState
export const get = useApp.getState

/** Open workbooks, keyed by tab id (kept out of React state). */
export const books = new Map<string, Book>()

export function activeTab(): TabMeta | null {
  return get().tabs.find((t) => t.id === get().active) ?? null
}

export function activeBook(): Book | null {
  return books.get(get().active) ?? null
}

export function updateTab(id: string, patch: Partial<TabMeta>) {
  set((s) => ({ tabs: s.tabs.map((t) => (t.id === id ? { ...t, ...patch } : t)) }))
}

let toastSeq = 0
export function toast(message: string, kind: Toast['kind'] = 'info') {
  const id = ++toastSeq
  set((s) => ({ toasts: [...s.toasts.slice(-3), { id, kind, message }] }))
  setTimeout(() => set((s) => ({ toasts: s.toasts.filter((t) => t.id !== id) })), kind === 'error' ? 7000 : 3500)
}

export async function updateSettings(patch: Partial<Settings>) {
  set((s) => ({ settings: { ...s.settings, ...patch } }))
  const next = await window.simpleCell.settingsSet(patch)
  set({ settings: next })
}
