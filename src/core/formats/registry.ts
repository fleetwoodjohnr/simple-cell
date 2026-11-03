// File format dispatch. Every adapter turns bytes into a Workbook (with notes on anything it could not keep) and back.
import type { Workbook } from '../model/workbook'
import { exportDelimited, importDelimited } from './csv'
import { readXlsx } from './xlsx/read'
import { writeXlsx } from './xlsx/write'
import { readOds } from './ods/read'
import { writeOds } from './ods/write'
import { exportJson, importJson } from './json'
import { exportHtml, importHtml } from './html'
import { exportXml2003, importXml2003 } from './xml2003'
import type { ExportResult, FormatId, ImportResult } from './types'

export type { FormatId }

export async function importWorkbook(format: FormatId, bytes: Uint8Array, opts: Record<string, unknown> = {}): Promise<ImportResult> {
  switch (format) {
    case 'csv':
      return importDelimited(bytes, { delimiter: (opts.delimiter as string) ?? undefined })
    case 'tsv':
      return importDelimited(bytes, { delimiter: '\t' })
    case 'xlsx':
    case 'xlsm':
    case 'xltx':
    case 'xltm':
      return readXlsx(bytes)
    case 'ods':
      return readOds(bytes)
    case 'json':
      return importJson(bytes)
    case 'html':
      return importHtml(bytes)
    case 'xml':
      return importXml2003(bytes)
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
    case 'xlsx':
    case 'xlsm':
    case 'xltx':
    case 'xltm':
      return writeXlsx(wb, format, (opts.defaultFont as { name: string; size: number }) ?? undefined)
    case 'ods':
      return writeOds(wb)
    case 'json':
      return exportJson(wb, { sheet: opts.sheet as number | undefined })
    case 'html':
      return exportHtml(wb, { sheet: opts.sheet as number | undefined })
    case 'xml':
      return exportXml2003(wb)
    default:
      throw new Error(`Saving as ${format.toUpperCase()} is not supported yet.`)
  }
}
