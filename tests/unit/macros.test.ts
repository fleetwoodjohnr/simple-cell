import { describe, expect, it } from 'vitest'
import { unzipSync, zipSync, strFromU8 } from 'fflate'
import { Host } from '@/core/host'
import { decompress, readCfb, vbaModules } from '@/core/vba'

/** MS-OVBA compression with literal tokens only (valid, just not small). */
function compress(data: Uint8Array): Uint8Array {
  const out: number[] = [1]
  for (let p = 0; p < data.length; ) {
    const chunk = data.subarray(p, p + 3640) // 455 groups × 9 bytes = 4095 ≤ 4096 bytes per chunk
    p += chunk.length
    const body: number[] = []
    for (let i = 0; i < chunk.length; i += 8) {
      body.push(0)
      for (let j = i; j < Math.min(i + 8, chunk.length); j++) body.push(chunk[j])
    }
    const header = 0x8000 | 0x3000 | (body.length + 2 - 3)
    out.push(header & 0xff, header >> 8, ...body)
  }
  return new Uint8Array(out)
}

/** A minimal compound file: Root ▸ VBA (storage) ▸ dir + module streams, all in the mini stream. */
function cfb(streams: { name: string; data: Uint8Array }[]): Uint8Array {
  const S = 512
  const mini: Uint8Array[] = []
  const entries: { name: string; type: number; child: number; right: number; start: number; size: number }[] = []
  let miniSectors = 0
  const miniFat: number[] = []
  const placed = streams.map((s) => {
    const n = Math.ceil(s.data.length / 64)
    const start = miniSectors
    for (let i = 0; i < n; i++) miniFat.push(i === n - 1 ? 0xfffffffe : start + i + 1)
    miniSectors += n
    const padded = new Uint8Array(n * 64)
    padded.set(s.data)
    mini.push(padded)
    return { ...s, start }
  })
  const ministream = new Uint8Array(miniSectors * 64)
  let off = 0
  for (const m of mini) (ministream.set(m, off), (off += m.length))
  entries.push({ name: 'Root Entry', type: 5, child: 1, right: 0xffffffff, start: 0, size: ministream.length })
  entries.push({ name: 'VBA', type: 1, child: 2, right: 0xffffffff, start: 0, size: 0 })
  placed.forEach((p, i) => entries.push({ name: p.name, type: 2, child: 0xffffffff, right: i + 1 < placed.length ? 3 + i : 0xffffffff, start: p.start, size: p.data.length }))
  const dirSectors = Math.ceil((entries.length * 128) / S)
  const miniFatSectors = Math.ceil((miniFat.length * 4) / S)
  const msSectors = Math.ceil(ministream.length / S)
  // Sector 0: FAT; then directory, mini FAT, ministream.
  const total = 1 + dirSectors + miniFatSectors + msSectors
  const fat = new Array(S / 4).fill(0xffffffff)
  fat[0] = 0xfffffffd
  const run = (start: number, n: number) => {
    for (let i = 0; i < n; i++) fat[start + i] = i === n - 1 ? 0xfffffffe : start + i + 1
  }
  run(1, dirSectors)
  run(1 + dirSectors, miniFatSectors)
  run(1 + dirSectors + miniFatSectors, msSectors)
  const buf = new Uint8Array(S * (1 + total))
  const dv = new DataView(buf.buffer)
  dv.setUint32(0, 0xe011cfd0, true)
  dv.setUint32(4, 0xe11ab1a1, true)
  dv.setUint16(24, 0x3e, true)
  dv.setUint16(26, 3, true)
  dv.setUint16(28, 0xfffe, true)
  dv.setUint16(30, 9, true)
  dv.setUint16(32, 6, true)
  dv.setUint32(44, 1, true)
  dv.setUint32(48, 1, true)
  dv.setUint32(56, 4096, true)
  dv.setUint32(60, 1 + dirSectors, true)
  dv.setUint32(64, miniFatSectors, true)
  dv.setUint32(68, 0xfffffffe, true)
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, i === 0 ? 0 : 0xffffffff, true)
  const sec = (n: number) => S * (1 + n)
  fat.forEach((v, i) => dv.setUint32(sec(0) + i * 4, v, true))
  entries.forEach((e, i) => {
    const o = sec(1) + i * 128
    for (let k = 0; k < e.name.length; k++) dv.setUint16(o + k * 2, e.name.charCodeAt(k), true)
    dv.setUint16(o + 64, (e.name.length + 1) * 2, true)
    buf[o + 66] = e.type
    dv.setUint32(o + 68, 0xffffffff, true)
    dv.setUint32(o + 72, e.right, true)
    dv.setUint32(o + 76, e.child, true)
    dv.setUint32(o + 116, e.type === 5 ? 1 + dirSectors + miniFatSectors : e.start, true)
    dv.setUint32(o + 120, e.size, true)
  })
  miniFat.forEach((v, i) => dv.setUint32(sec(1 + dirSectors) + i * 4, v, true))
  buf.set(ministream, sec(1 + dirSectors + miniFatSectors))
  return buf
}

function vbaProject(modules: { name: string; code: string }[]): Uint8Array {
  const enc = new TextEncoder()
  const rec = (id: number, data: Uint8Array) => {
    const out = new Uint8Array(6 + data.length)
    const dv = new DataView(out.buffer)
    dv.setUint16(0, id, true)
    dv.setUint32(2, data.length, true)
    out.set(data, 6)
    return out
  }
  const u16 = (n: number) => new Uint8Array([n & 0xff, n >> 8])
  const u32 = (n: number) => new Uint8Array([n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, n >>> 24])
  const parts: Uint8Array[] = [rec(0x0001, u32(1)), rec(0x0003, u16(1252)), rec(0x0004, enc.encode('VBAProject')), rec(0x000f, u16(modules.length))]
  for (const m of modules) parts.push(rec(0x0019, enc.encode(m.name)), rec(0x001a, enc.encode(m.name)), rec(0x0031, u32(0)), rec(0x002b, new Uint8Array()))
  parts.push(rec(0x0010, new Uint8Array()))
  const dir = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let o = 0
  for (const p of parts) (dir.set(p, o), (o += p.length))
  return cfb([{ name: 'dir', data: compress(dir) }, ...modules.map((m) => ({ name: m.name, data: compress(enc.encode(`Attribute VB_Name = "${m.name}"\r\n${m.code.replace(/\n/g, '\r\n')}`)) }))])
}

describe('macros', () => {
  it('JavaScript macros are kept in XLSX (customXml) and ODS', async () => {
    const h = new Host()
    h.setMacros([{ name: 'Double', code: 'async function main(wb) {\n  // a ]]> in a comment\n  await wb.activeSheet.range("A1").setValue(2)\n}' }])
    const xlsx = await h.save('xlsx')
    const files = unzipSync(xlsx.data)
    expect(strFromU8(files['customXml/item1.xml'])).toContain('urn:simple-cell:macros:1')
    expect(strFromU8(files['[Content_Types].xml'])).toContain('customXmlProperties+xml')
    const back = await Host.open(xlsx.data, 'xlsx')
    expect(back.macros()).toEqual(h.macros())
    expect(back.warnings.some((w) => /JavaScript macro/.test(w.message))).toBe(true)
    const ods = await Host.open((await h.save('ods')).data, 'ods')
    expect(ods.macros()).toEqual(h.macros())
  })

  it('VBA source is read from vbaProject.bin (compound file, compressed streams), never run', async () => {
    const mods = [
      { name: 'Module1', code: 'Sub Hello()\n    MsgBox "Hi"\nEnd Sub' },
      { name: 'Sheet1', code: "' long\n" + 'x = 1\n'.repeat(400) } // compressed size stays under the 4096-byte mini-stream cutoff
    ]
    const bin = vbaProject(mods)
    expect([...readCfb(bin).keys()].sort()).toEqual(['VBA/Module1', 'VBA/Sheet1', 'VBA/dir'])
    expect(new TextDecoder().decode(decompress(compress(new TextEncoder().encode('abc'.repeat(3000)))))).toBe('abc'.repeat(3000))
    expect(vbaModules(bin)).toEqual(mods.map((m) => ({ name: m.name, code: m.code.trimEnd() })))
    // Through the host, from an .xlsm.
    const h = new Host()
    const files = unzipSync((await h.save('xlsm')).data)
    files['xl/vbaProject.bin'] = bin as Uint8Array<ArrayBuffer>
    const opened = await Host.open(zipSync(files), 'xlsm')
    expect(opened.vba()?.map((m) => m.name)).toEqual(['Module1', 'Sheet1'])
  })
})
