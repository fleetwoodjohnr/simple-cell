// Ribbon KeyTips (Windows/Linux): tap Alt, then letters — a tab first, then a command on it. Esc steps back.
import { useEffect, useState } from 'react'
import { create } from 'zustand'
import { get, set } from '../state/store'
import type { RibbonTab } from '../state/store'

/** Excel's tab letters. */
const TAB_KEYS: Record<string, RibbonTab | 'file'> = { F: 'file', H: 'home', N: 'insert', M: 'formulas', A: 'data', W: 'view', Y: 'help', J: 'table' }

interface Tip {
  key: string
  x: number
  y: number
  el: HTMLElement
}

const useTips = create<{ stage: 'off' | 'tabs' | 'commands'; typed: string; tips: Tip[] }>(() => ({ stage: 'off', typed: '', tips: [] }))

/** Letters for the visible commands of the open tab: first letters of their names, two letters where they clash. */
function commandTips(): Tip[] {
  const bar = document.querySelector('[data-testid="ribbon"] [role="toolbar"]')
  if (!bar) return []
  const els = [...bar.querySelectorAll<HTMLElement>('button[aria-label], input[aria-label], select[aria-label]')].filter((e) => e.offsetParent !== null)
  const used = new Set<string>()
  const pick = (label: string): string => {
    const letters = label.toUpperCase().replace(/[^A-Z0-9]/g, '')
    for (const ch of letters) if (!used.has(ch) && !TAB_KEYS[ch]) return ch
    for (const a of letters) for (const b of letters + 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') if (!used.has(a + b) && ![...used].some((u) => u.length === 1 && u === a)) return a + b
    return ''
  }
  const out: Tip[] = []
  for (const el of els) {
    const key = pick(el.getAttribute('aria-label') ?? '')
    if (!key) continue
    used.add(key)
    const r = el.getBoundingClientRect()
    out.push({ key, x: r.left + r.width / 2, y: r.bottom - 6, el })
  }
  return out
}

function tabTips(): Tip[] {
  const out: Tip[] = []
  const file = document.querySelector<HTMLElement>('[data-testid="file-tab"]')
  const add = (key: string, el: HTMLElement | null) => {
    if (!el) return
    const r = el.getBoundingClientRect()
    out.push({ key, x: r.left + r.width / 2, y: r.bottom - 4, el })
  }
  add('F', file)
  for (const [key, id] of Object.entries(TAB_KEYS)) if (id !== 'file') add(key, document.querySelector<HTMLElement>(`[data-testid="ribbon-tab-${id}"]`))
  return out
}

const isMac = typeof navigator !== 'undefined' && /Mac/i.test(navigator.platform)
let altAlone = false

export function keyTipsOff() {
  useTips.setState({ stage: 'off', typed: '', tips: [] })
}

/** Called first by the global keydown handler. Returns true when KeyTips used the key. */
export function keyTipsKeydown(e: KeyboardEvent): boolean {
  if (isMac) return false
  altAlone = e.key === 'Alt' && !e.ctrlKey && !e.shiftKey && !e.metaKey
  const { stage, typed, tips } = useTips.getState()
  if (stage === 'off') return false
  e.preventDefault()
  e.stopPropagation()
  if (e.key === 'Escape') {
    if (stage === 'commands') useTips.setState({ stage: 'tabs', typed: '', tips: tabTips() })
    else keyTipsOff()
    return true
  }
  if (e.key === 'Alt') {
    keyTipsOff()
    return true
  }
  const ch = e.key.length === 1 ? e.key.toUpperCase() : ''
  if (!ch) return true
  const next = typed + ch
  const hit = tips.find((t) => t.key === next)
  if (stage === 'tabs' && hit) {
    const id = TAB_KEYS[next]
    if (id === 'file') {
      keyTipsOff()
      set({ backstage: true })
      return true
    }
    set({ ribbonTab: id })
    if (get().settings.ribbonCollapsed) void import('../state/store').then((m) => m.updateSettings({ ribbonCollapsed: false }))
    // The tab renders first; then its commands get their letters.
    requestAnimationFrame(() => requestAnimationFrame(() => useTips.setState({ stage: 'commands', typed: '', tips: commandTips() })))
    return true
  }
  if (stage === 'commands' && hit) {
    keyTipsOff()
    hit.el.focus()
    hit.el.click()
    return true
  }
  if (tips.some((t) => t.key.startsWith(next))) useTips.setState({ typed: next })
  return true
}

/** Alt released on its own toggles KeyTips. */
export function keyTipsKeyup(e: KeyboardEvent) {
  if (isMac || e.key !== 'Alt' || !altAlone) return
  altAlone = false
  e.preventDefault()
  const { stage } = useTips.getState()
  if (stage === 'off' && !document.querySelector('[role="dialog"]')) useTips.setState({ stage: 'tabs', typed: '', tips: tabTips() })
}

export function KeyTipsLayer() {
  const { stage, typed, tips } = useTips()
  const [, force] = useState(0)
  useEffect(() => {
    if (stage === 'off') return
    const off = () => keyTipsOff()
    const resize = () => force((n) => n + 1)
    window.addEventListener('mousedown', off, true)
    window.addEventListener('resize', resize)
    window.addEventListener('blur', off)
    return () => {
      window.removeEventListener('mousedown', off, true)
      window.removeEventListener('resize', resize)
      window.removeEventListener('blur', off)
    }
  }, [stage])
  if (stage === 'off') return null
  return (
    <div className="pointer-events-none fixed inset-0 z-[60]" data-testid="keytips">
      {tips
        .filter((t) => t.key.startsWith(typed))
        .map((t) => (
          <span key={t.key} data-for={t.el.getAttribute('aria-label') ?? t.el.textContent ?? ''} className="absolute -translate-x-1/2 rounded-[3px] border border-[#3b3b3b] bg-[#f5f5f5] px-1 text-[10.5px] font-semibold leading-[15px] text-[#1f1f1f] shadow" style={{ left: t.x, top: t.y }}>
            {t.key}
          </span>
        ))}
    </div>
  )
}
