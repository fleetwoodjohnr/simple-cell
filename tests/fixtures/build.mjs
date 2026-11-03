// Regenerates the binary fixtures from tests/fixtures/src/*.fods with LibreOffice (a different producer than the code
// under test). Usage: node tests/fixtures/build.mjs
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readdirSync, copyFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'

const here = import.meta.dirname
const profile = pathToFileURL(mkdtempSync(path.join(tmpdir(), 'sc-lo-profile-'))).href
const out = mkdtempSync(path.join(tmpdir(), 'sc-fixtures-'))
for (const f of readdirSync(path.join(here, 'src')).filter((f) => f.endsWith('.fods'))) {
  const src = path.join(here, 'src', f)
  for (const target of ['xlsx:Calc MS Excel 2007 XML', 'ods']) execFileSync('soffice', ['--headless', '--norestore', `-env:UserInstallation=${profile}`, '--convert-to', target, '--outdir', out, src], { stdio: 'inherit' })
}
for (const f of readdirSync(out)) copyFileSync(path.join(out, f), path.join(here, f))
console.log('fixtures:', readdirSync(out).join(', '))
