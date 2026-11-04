// Returning the keyboard to the grid after a dialog. Done at once, not in a later frame or timer: on a slow or
// throttled window (a background window, a CI runner) the frame came late and keys typed meanwhile were lost.
import { flushSync } from 'react-dom'
import { set } from '../state/store'

export const gridInput = () => document.querySelector<HTMLElement>('[data-grid-focus]:not([hidden])')

export function focusGrid() {
  gridInput()?.focus()
}

/** Close the open dialog, committing the render now so its focus trap is gone before the grid takes focus. */
export function closeDialog() {
  flushSync(() => set({ dialog: null }))
}
