import { useCallback, useSyncExternalStore } from 'react'
import type { Book } from '../state/client'

/** Re-render when the workbook's cache or metadata changes. */
export function useBookVersion(book: Book | undefined): number {
  return useSyncExternalStore(
    useCallback((fn) => (book ? book.subscribe(fn) : () => undefined), [book]),
    () => book?.version ?? 0
  )
}
