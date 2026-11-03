import { useEffect, useState } from 'react'
import { useShallow } from 'zustand/react/shallow'
import { CheckCircle2, AlertCircle, Info, FileUp } from 'lucide-react'
import { HOME, useApp, set, get, books, toast } from '../state/store'
import * as docs from '../state/documents'
import { handleKeydown, runCommand } from './commands'
import { TitleBar } from './TitleBar'
import { Ribbon } from './Ribbon'
import { Dashboard, Backstage } from './Dashboard'
import { Dialogs } from './Dialogs'
import { KeyTipsLayer, keyTipsKeyup } from './KeyTips'
import { FormulaBar } from './FormulaBar'
import { StatusBar } from './SheetBar'
import { TooltipProvider } from '../ui'
import { Grid } from '@/grid/Grid'
import { getView, useViews } from '../state/view'
import * as A from './actions'

async function boot() {
  const api = window.simpleCell
  const [settings, recents, recovery, libreOffice] = await Promise.all([api.settingsGet(), api.recentList(), api.recoveryList(), api.libreOfficeAvailable()])
  set({ settings, recents, recovery, libreOffice, ready: true })
}

function useTheme() {
  const { theme, accent, reduceMotion } = useApp(useShallow((s) => ({ theme: s.settings.theme, accent: s.settings.accent, reduceMotion: s.settings.reduceMotion })))
  useEffect(() => {
    const mq = window.matchMedia('(prefers-color-scheme: dark)')
    const apply = () => {
      const dark = theme === 'dark' || (theme === 'system' && mq.matches)
      document.documentElement.classList.toggle('dark', dark)
      document.documentElement.style.setProperty('--accent', accent)
      set({ dark })
      window.simpleCell.setTitleBarColors({ color: dark ? '#121214' : '#f7f7f8', symbolColor: dark ? '#d4d4d8' : '#3f3f46' })
    }
    apply()
    mq.addEventListener('change', apply)
    return () => mq.removeEventListener('change', apply)
  }, [theme, accent])
  useEffect(() => {
    document.documentElement.classList.toggle('reduce-motion', reduceMotion)
  }, [reduceMotion])
}

function Toasts() {
  const toasts = useApp((s) => s.toasts)
  return (
    <div className="pointer-events-none fixed inset-x-0 bottom-12 z-[120] flex flex-col items-center gap-2" role="status" aria-live="polite">
      {toasts.map((t) => (
        <div key={t.id} className="slide-up pointer-events-auto flex max-w-[560px] items-center gap-2.5 rounded-xl bg-[#1c1c22] px-4 py-2.5 text-[13px] text-white shadow-[0_12px_32px_-8px_rgb(0_0_0/0.45)] dark:bg-[#2a2a30]">
          {t.kind === 'success' ? <CheckCircle2 className="size-4 text-emerald-400" /> : t.kind === 'error' ? <AlertCircle className="size-4 text-red-400" /> : <Info className="size-4 text-sky-300" />}
          {t.message}
        </div>
      ))}
    </div>
  )
}

function DropOverlay() {
  const [over, setOver] = useState(false)
  useEffect(() => {
    let depth = 0
    const hasFiles = (e: DragEvent) => e.dataTransfer?.types.includes('Files')
    const enter = (e: DragEvent) => {
      if (!hasFiles(e)) return
      depth++
      setOver(true)
    }
    const leave = () => {
      depth = Math.max(0, depth - 1)
      if (!depth) setOver(false)
    }
    const overH = (e: DragEvent) => {
      if (hasFiles(e)) e.preventDefault()
    }
    const drop = (e: DragEvent) => {
      depth = 0
      setOver(false)
      if (!hasFiles(e)) return
      e.preventDefault()
      const files = Array.from(e.dataTransfer?.files ?? [])
      if (files.length) void docs.openDroppedFiles(files)
    }
    window.addEventListener('dragenter', enter)
    window.addEventListener('dragleave', leave)
    window.addEventListener('dragover', overH)
    window.addEventListener('drop', drop)
    return () => {
      window.removeEventListener('dragenter', enter)
      window.removeEventListener('dragleave', leave)
      window.removeEventListener('dragover', overH)
      window.removeEventListener('drop', drop)
    }
  }, [])
  if (!over) return null
  return (
    <div className="fade-in pointer-events-none fixed inset-3 z-[110] flex items-center justify-center rounded-2xl border-2 border-dashed border-accent bg-accent-soft backdrop-blur-sm">
      <div className="flex flex-col items-center gap-2 text-accent-ink">
        <FileUp className="size-8" />
        <div className="text-[15px] font-semibold">Drop to open</div>
      </div>
    </div>
  )
}

function Workbook({ tab, hidden }: { tab: string; hidden: boolean }) {
  const showFormulaBar = useApp((s) => s.showFormulaBar)
  const view = useViews((s) => s.views[tab])
  if (!books.get(tab) || !view) return null
  return (
    <div hidden={hidden} className="absolute inset-0 flex flex-col">
      {!hidden && <Ribbon tab={tab} />}
      {showFormulaBar && !hidden && <FormulaBar tab={tab} />}
      <div className="relative min-h-0 flex-1">
        <Grid tab={tab} hidden={hidden} />
      </div>
      {!hidden && <StatusBar tab={tab} />}
    </div>
  )
}

export function App() {
  const { ready, active, tabs } = useApp(useShallow((s) => ({ ready: s.ready, active: s.active, tabs: s.tabs })))
  useTheme()
  useEffect(() => {
    void boot()
    const api = window.simpleCell
    window.addEventListener('keydown', handleKeydown, true)
    window.addEventListener('keyup', keyTipsKeyup, true)
    const offs = [api.onOpenFiles((f) => void docs.openFiles(f)), api.onCloseRequest(() => docs.requestQuit()), api.onMenuCommand((id) => void runCommand(id))]
    // Test/automation hook (renderer-only; exposes nothing the page could not already reach).
    ;(window as unknown as Record<string, unknown>).__sc = { get, set, books, docs, runCommand, getView, A, toast }
    return () => {
      window.removeEventListener('keydown', handleKeydown, true)
      window.removeEventListener('keyup', keyTipsKeyup, true)
      offs.forEach((o) => o())
    }
  }, [])

  if (!ready) return <div className="h-full bg-bg" />
  const home = active === HOME
  return (
    <TooltipProvider>
      <div className="flex h-full flex-col">
        <TitleBar />
        <div className="relative flex min-h-0 flex-1 flex-col">
          <main className="relative min-w-0 flex-1" aria-label="Workbook">
            {tabs.map((t) => (
              <Workbook key={t.id} tab={t.id} hidden={home || t.id !== active} />
            ))}
            {home && (
              <div className="absolute inset-0">
                <Dashboard />
              </div>
            )}
          </main>
          {!home && <Backstage />}
        </div>
        <Dialogs />
        <KeyTipsLayer />
        <Toasts />
        <DropOverlay />
      </div>
    </TooltipProvider>
  )
}
