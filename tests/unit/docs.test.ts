// The reference documents are generated from the registries, so they can't drift from the code.
// `UPDATE_DOCS=1 npx vitest run tests/unit/docs.test.ts` rewrites them; otherwise this test fails if they are stale.
import { describe, expect, it } from 'vitest'
import { readFileSync, writeFileSync, existsSync } from 'node:fs'
import path from 'node:path'
import '@/core/doc'
import { allFunctions } from '@/core/engine/evaluate'
import { CATALOG } from '@/core/engine/catalog'
import { CASES, DEVIATIONS } from './functions/cases'
import { SHORTCUTS, BUILTIN, RESERVED } from '@/renderer/app/shortcuts'

const root = path.join(import.meta.dirname, '../..')
const update = process.env.UPDATE_DOCS === '1'

function check(file: string, content: string) {
  const p = path.join(root, file)
  if (update) writeFileSync(p, content)
  expect(existsSync(p) ? readFileSync(p, 'utf8') : '', `${file} is out of date: run UPDATE_DOCS=1 npx vitest run tests/unit/docs.test.ts`).toBe(content)
}

/** Why a function isn't calculated (or calculated with a limit). */
const NOTES: Record<string, string> = {
  'FORECAST.ETS': 'Exponential triple smoothing (AAA ETS) is not implemented yet; formulas keep the value saved in the file.',
  'FORECAST.ETS.CONFINT': 'See FORECAST.ETS.',
  'FORECAST.ETS.SEASONALITY': 'See FORECAST.ETS.',
  'FORECAST.ETS.STAT': 'See FORECAST.ETS.',
  WEBSERVICE: 'Needs network access during calculation; planned behind an explicit permission.',
  FILTERXML: 'Needs an XPath engine; planned.',
  IMPORTCSV: 'Needs file/network access during calculation; planned.',
  IMPORTTEXT: 'Needs file/network access during calculation; planned.',
  STOCKHISTORY: 'Microsoft online data service; not available.',
  TRANSLATE: 'Microsoft online translation service; not available.',
  DETECTLANGUAGE: 'Microsoft online service; not available.',
  COPILOT: 'Microsoft Copilot service; not available.',
  PY: 'Python in Excel runs in Microsoft’s cloud; not available.',
  IMAGE: 'Images in cells are planned.',
  FIELDVALUE: 'Linked data types (Stocks, Geography) are Microsoft online services; not available.',
  GETPIVOTDATA: 'Arrives with PivotTables (M6).',
  PIVOTBY: 'Planned with PivotTables (M6).',
  RTD: 'COM real-time data servers are Windows-only Excel add-ins; not available.',
  CALL: 'Calls DLL procedures; not available (security).',
  'REGISTER.ID': 'Registers DLL procedures; not available (security).',
  PHONETIC: 'Returns the text itself: furigana (phonetic guides) are not stored.',
  ASC: 'Full-width to half-width ASCII and spaces; katakana conversion is not included.',
  DBCS: 'Half-width to full-width ASCII and spaces; katakana conversion is not included.',
  JIS: 'Same as DBCS.',
  LENB: 'Counts characters (single-byte languages); double-byte byte counts are not used.',
  LEFTB: 'Same as LEFT (single-byte languages).',
  RIGHTB: 'Same as RIGHT (single-byte languages).',
  MIDB: 'Same as MID (single-byte languages).',
  FINDB: 'Same as FIND (single-byte languages).',
  SEARCHB: 'Same as SEARCH (single-byte languages).',
  REPLACEB: 'Same as REPLACE (single-byte languages).',
  CELL: '"color", "prefix", "protect", "format" and "width" follow Excel’s rules for common formats; "filename" needs a saved file.',
  INFO: '"osversion" and "system" report the platform Simple Cell runs on.',
  REGEXEXTRACT: 'Uses JavaScript regular expressions, which differ from Excel’s PCRE2 in rare constructs.',
  REGEXREPLACE: 'See REGEXEXTRACT.',
  REGEXTEST: 'See REGEXEXTRACT.',
  GROUPBY: 'Core grouping, aggregation, headers, totals and sorting; subtotal depth beyond 1 and field relationships are not implemented.',
  ACCRINT: 'Accrued interest from issue to settlement (calc_method TRUE); first_interest is not used for the FALSE method yet.',
  ODDFPRICE: 'Short first periods follow Microsoft’s formula exactly; long first periods use the documented quasi-coupon method (not verified against Excel).',
  ODDFYIELD: 'Solves ODDFPRICE for the yield.',
  'CUBEKPIMEMBER': 'OLAP cube connections are not supported.',
  CUBEMEMBER: 'OLAP cube connections are not supported.',
  CUBEMEMBERPROPERTY: 'OLAP cube connections are not supported.',
  CUBERANKEDMEMBER: 'OLAP cube connections are not supported.',
  CUBESET: 'OLAP cube connections are not supported.',
  CUBESETCOUNT: 'OLAP cube connections are not supported.',
  CUBEVALUE: 'OLAP cube connections are not supported.',
  EUROCONVERT: 'Fixed euro conversion rates for the 20 legacy currencies.',
  DAVERAGE: 'Criteria ranges with formula (computed) criteria are not supported yet; value criteria work as in Excel.'
}

describe('generated documentation', () => {
  it('docs/FORMULA_COVERAGE.md', () => {
    const impl = new Set(allFunctions().map((f) => f.name))
    const tested = new Set(Object.keys(CASES))
    const cats = [...new Set(CATALOG.map((e) => e.cat))].sort()
    const total = CATALOG.length
    const done = CATALOG.filter((e) => impl.has(e.n)).length
    const withTests = CATALOG.filter((e) => impl.has(e.n) && tested.has(e.n)).length
    const caseCount = Object.values(CASES).reduce((s, l) => s + l.length, 0)
    const out: string[] = [
      '# Formula coverage',
      '',
      '_Generated from the function catalog (`src/core/engine/catalog.json`, built from Microsoft’s “Excel functions (alphabetical)” page) and the function tests (`tests/unit/functions/cases.ts`). Do not edit by hand: run `UPDATE_DOCS=1 npx vitest run tests/unit/docs.test.ts`._',
      '',
      `**${done} of ${total}** Excel worksheet functions are calculated. **${withTests}** of them have dedicated test cases (${caseCount} formulas, most taken from the examples in Microsoft’s documentation).`,
      'The formulas are also run through LibreOffice Calc as an independent check ([LIBREOFFICE_CROSSCHECK.md](LIBREOFFICE_CROSSCHECK.md)). Simple Cell has not been checked against Excel itself: no copy of Excel was available to the project.',
      '',
      'A formula that uses a function Simple Cell does not calculate is kept unchanged and shows the result saved in the file, marked as not recalculated. Without a saved result it shows `#NAME?`. It is never shown as 0.',
      '',
      '| Status | Meaning |',
      '|---|---|',
      '| ✅ | Calculated, with test cases |',
      '| 🟡 | Calculated, no dedicated test case yet |',
      '| ⛔ | Not calculated (formula kept; saved result shown) |',
      '',
      '## By category',
      '',
      '| Category | Calculated | Tested | Total |',
      '|---|---|---|---|'
    ]
    for (const c of cats) {
      const list = CATALOG.filter((e) => e.cat === c)
      out.push(`| ${c} | ${list.filter((e) => impl.has(e.n)).length} | ${list.filter((e) => impl.has(e.n) && tested.has(e.n)).length} | ${list.length} |`)
    }
    out.push(`| **All** | **${done}** | **${withTests}** | **${total}** |`, '')
    if (Object.keys(DEVIATIONS).length) {
      out.push('## Known deviations from Excel', '', '| Function | Deviation |', '|---|---|')
      for (const [k, v] of Object.entries(DEVIATIONS)) out.push(`| ${k} | ${v.note} |`)
      out.push('')
    }
    for (const c of cats) {
      out.push(`## ${c}`, '', '| Function | Status | Since | Notes |', '|---|---|---|---|')
      for (const e of CATALOG.filter((x) => x.cat === c)) {
        const st = impl.has(e.n) ? (tested.has(e.n) ? '✅' : '🟡') : '⛔'
        out.push(`| ${e.n} | ${st} | ${e.v ?? '—'} | ${(NOTES[e.n] ?? '').replace(/\|/g, '\\|')} |`)
      }
      out.push('')
    }
    check('docs/FORMULA_COVERAGE.md', out.join('\n'))
  })

  it('docs/SHORTCUT_REFERENCE.md', () => {
    const src = readFileSync(path.join(root, 'src/renderer/app/commands.ts'), 'utf8')
    const labels = new Map([...src.matchAll(/def\(\{ id: '([^']+)', label: '([^']+)'/g)].map((m) => [m[1], m[2].replace(/\\u2019/g, '’')]))
    for (const s of SHORTCUTS) expect(labels.has(s.command), `shortcut for unknown command ${s.command}`).toBe(true)
    const fmt = (k: string, mac: boolean) =>
      k
        .split('-')
        .map((p) => (p === 'Mod' ? (mac ? '⌘' : 'Ctrl') : p === 'Alt' ? (mac ? '⌥' : 'Alt') : p === 'Shift' ? (mac ? '⇧' : 'Shift') : p === 'Ctrl' ? '⌃' : p === 'Minus' ? '-' : p === 'Plus' ? '+' : p.length === 1 ? p.toUpperCase() : p))
        .join(mac ? '' : '+')
    const out = [
      '# Keyboard shortcuts',
      '',
      '_Generated from `src/renderer/app/shortcuts.ts`. Run `UPDATE_DOCS=1 npx vitest run tests/unit/docs.test.ts` after changing it._',
      '',
      'Simple Cell follows Microsoft Excel’s shortcuts (from Microsoft’s “Keyboard shortcuts in Excel” pages for Windows and Mac). Linux uses the Windows keys. Every shortcut can be changed in **Help → Shortcuts**. The **Google Sheets** preset (Settings → Keyboard) switches the ones that differ.',
      '',
      '## Commands',
      '',
      '| Action | Windows / Linux | macOS | Google Sheets preset |',
      '|---|---|---|---|'
    ]
    for (const s of SHORTCUTS) {
      const win = s.win.map((k) => `\`${fmt(k, false)}\``).join(', ') || '—'
      const mac = (s.mac ?? s.win).map((k) => `\`${fmt(k, true)}\``).join(', ') || '—'
      const sh = s.sheets ? s.sheets.map((k) => `\`${fmt(k, false)}\``).join(', ') : 'same'
      out.push(`| ${labels.get(s.command)} | ${win} | ${mac} | ${sh} |`)
    }
    out.push('', '## Navigation and editing', '', '| Action | Windows / Linux | macOS | Status |', '|---|---|---|---|')
    for (const b of BUILTIN) out.push(`| ${b.action} | \`${fmt(b.keys, false)}\` | \`${fmt(b.mac ?? b.keys, true)}\` | ${b.status === 'done' ? '✅' : 'Planned'} |`)
    out.push('', '## Reserved by the operating system', '', 'These keys never reach Simple Cell:', '', '| Keys | System | Note |', '|---|---|---|')
    for (const r of RESERVED) out.push(`| \`${fmt(r.keys, r.os === 'macOS')}\` | ${r.os} | ${r.note} |`)
    out.push('')
    check('docs/SHORTCUT_REFERENCE.md', out.join('\n'))
  })

  it('shortcut bindings are written in the order key events produce (Mod, Ctrl, Alt, Shift, key)', () => {
    for (const s of SHORTCUTS)
      for (const k of [...s.win, ...(s.mac ?? []), ...(s.sheets ?? [])]) expect(k, `${s.command}: ${k}`).toMatch(/^(Mod-)?(Ctrl-)?(Alt-)?(Shift-)?[^-]+$/)
  })

  it('shortcut registry has no conflicting bindings', () => {
    for (const platform of ['win', 'mac'] as const)
      for (const preset of ['excel', 'sheets'] as const) {
        const seen = new Map<string, string>()
        for (const s of SHORTCUTS) {
          const keys = preset === 'sheets' && s.sheets ? s.sheets : platform === 'mac' ? (s.mac ?? s.win) : s.win
          for (const k of keys) {
            expect(seen.get(k), `${platform}/${preset}: ${k} is bound to both ${seen.get(k)} and ${s.command}`).toBeUndefined()
            seen.set(k, s.command)
          }
        }
      }
  })
})
