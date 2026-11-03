import type { SimpleCellApi } from '../shared/ipc-types'

declare global {
  interface Window {
    simpleCell: SimpleCellApi
  }
}
export {}
