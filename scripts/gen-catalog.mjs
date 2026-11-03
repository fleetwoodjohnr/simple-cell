// Regenerates src/core/engine/catalog.json from Microsoft's "Excel functions (alphabetical)" page.
// Usage: node scripts/gen-catalog.mjs   (needs curl and python3)
import { execFileSync } from 'node:child_process'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'

const URL = 'https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188'
const dir = mkdtempSync(path.join(tmpdir(), 'sc-catalog-'))
const html = path.join(dir, 'alpha.html')
execFileSync('curl', ['-sL', '-A', 'Mozilla/5.0 SimpleCell-catalog', '-o', html, URL])
const rows = JSON.parse(execFileSync('python3', ['-I', path.join(import.meta.dirname, 'parse-ms-functions.py'), html], { encoding: 'utf8' }))

const fixCategory = (r) => {
  if (r.category === 'Look and reference') return 'Lookup and reference'
  if (!r.category && r.description.startsWith('Add-in and Automation:')) return 'Add-in and Automation'
  return r.category
}
// Newer functions not (yet) on that page; listed so formulas using them are reported as unsupported, not #NAME?.
const extra = [
  { name: 'COPILOT', category: 'AI', description: 'Uses Copilot in Excel to generate a result from a prompt', since: 'Microsoft 365' },
  { name: 'PY', category: 'Python', description: 'Runs Python code in Excel (Python in Excel)', since: 'Microsoft 365' },
  { name: 'FIELDVALUE', category: 'Lookup and reference', description: 'Returns a field from a linked data type value', since: 'Microsoft 365' },
  { name: 'IMPORTCSV', category: 'Web', description: 'Imports data from a CSV file', since: 'Microsoft 365' },
  { name: 'IMPORTTEXT', category: 'Web', description: 'Imports data from a text file', since: 'Microsoft 365' }
]
const all = [...rows.map((r) => ({ ...r, category: fixCategory(r), description: r.description.replace(/^Add-in and Automation:\s*/, '') })), ...extra.filter((e) => !rows.some((r) => r.name === e.name))]
all.sort((a, b) => a.name.localeCompare(b.name))
const out = all.map((r) => ({ n: r.name, cat: r.category, d: r.description, v: r.since ?? null }))
const target = path.join(import.meta.dirname, '../src/core/engine/catalog.json')
writeFileSync(target, JSON.stringify(out, null, 0).replace(/\},\{/g, '},\n{') + '\n')
console.log(`${out.length} functions → ${path.relative(process.cwd(), target)}`)
