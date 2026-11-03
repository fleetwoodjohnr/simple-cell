// File format dispatch. Every adapter turns bytes into a Workbook (with notes on anything it could not keep) and back.
import type { Workbook } from '../model/workbook'
import { exportDelimited, importDelimited } from './csv'
import type { ExportResult, FormatId, ImportResult } from './types'

export type { FormatId }

export async function importWorkbook(format: FormatId, bytes: Uint8Array, opts: Record<string, unknown> = {}): Promise<ImportResult> {
  switch (format) {
    case 'csv':
      return importDelimited(bytes, { delimiter: (opts.delimiter as string) ?? undefined })
    case 'tsv':
      return importDelimited(bytes, { delimiter: '\t' })
    default:
      throw new Error(`Opening ${format.toUpperCase()} files is not supported yet.`)
  }
}

export async function exportWorkbook(format: FormatId, wb: Workbook, opts: Record<string, unknown> = {}): Promise<ExportResult> {
  switch (format) {
    case 'csv':
      return exportDelimited(wb, ',', { sheet: opts.sheet as number | undefined, guard: opts.guard !== false })
    case 'tsv':
      return exportDelimited(wb, '\t', { sheet: opts.sheet as number | undefined, guard: opts.guard !== false })
    default:
      throw new Error(`Saving as ${format.toUpperCase()} is not supported yet.`)
  }
}
