// Read-only VBA source: a Compound File Binary reader, MS-OVBA decompression and the dir stream's module list.
// Simple Cell never runs VBA; this only shows what a workbook's macros contain.

/** Streams of a compound file (vbaProject.bin), by path ("VBA/dir", "VBA/Module1"). */
export function readCfb(buf: Uint8Array): Map<string, Uint8Array> {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  if (dv.getUint32(0, true) !== 0xe011cfd0 || dv.getUint32(4, true) !== 0xe11ab1a1) throw new Error('Not a compound file')
  const sectorSize = 1 << dv.getUint16(30, true)
  const miniSize = 1 << dv.getUint16(32, true)
  const nFat = dv.getUint32(44, true)
  const dirStart = dv.getUint32(48, true)
  const miniCutoff = dv.getUint32(56, true)
  const miniFatStart = dv.getUint32(60, true)
  let difStart = dv.getUint32(68, true)
  const sector = (n: number) => 512 + n * sectorSize
  // FAT sector list: 109 in the header, then the DIFAT chain.
  const fatSectors: number[] = []
  for (let i = 0; i < 109 && fatSectors.length < nFat; i++) fatSectors.push(dv.getUint32(76 + i * 4, true))
  for (let guard = 0; fatSectors.length < nFat && difStart < 0xfffffffa && guard < 10000; guard++) {
    const off = sector(difStart)
    for (let i = 0; i < sectorSize / 4 - 1 && fatSectors.length < nFat; i++) fatSectors.push(dv.getUint32(off + i * 4, true))
    difStart = dv.getUint32(off + sectorSize - 4, true)
  }
  const fat: number[] = []
  for (const s of fatSectors) for (let i = 0; i < sectorSize / 4; i++) fat.push(dv.getUint32(sector(s) + i * 4, true))
  const chain = (start: number, table: number[]) => {
    const out: number[] = []
    for (let s = start; s < 0xfffffffa && out.length < 1_000_000; s = table[s] ?? 0xfffffffe) out.push(s)
    return out
  }
  const readChain = (start: number) => {
    const ss = chain(start, fat)
    const out = new Uint8Array(ss.length * sectorSize)
    ss.forEach((s, i) => out.set(buf.subarray(sector(s), sector(s) + sectorSize), i * sectorSize))
    return out
  }
  const dir = readChain(dirStart)
  const ddv = new DataView(dir.buffer)
  interface Entry {
    name: string
    type: number
    left: number
    right: number
    child: number
    start: number
    size: number
  }
  const entries: Entry[] = []
  for (let off = 0; off + 128 <= dir.length; off += 128) {
    const nameLen = ddv.getUint16(off + 64, true)
    let name = ''
    for (let i = 0; i < Math.max(0, nameLen / 2 - 1); i++) name += String.fromCharCode(ddv.getUint16(off + i * 2, true))
    entries.push({ name, type: dir[off + 66], left: ddv.getUint32(off + 68, true), right: ddv.getUint32(off + 72, true), child: ddv.getUint32(off + 76, true), start: ddv.getUint32(off + 116, true), size: ddv.getUint32(off + 120, true) })
  }
  const root = entries[0]
  const ministream = root ? readChain(root.start) : new Uint8Array()
  const miniFat: number[] = []
  if (miniFatStart < 0xfffffffa) {
    const mf = readChain(miniFatStart)
    const mdv = new DataView(mf.buffer)
    for (let i = 0; i < mf.length / 4; i++) miniFat.push(mdv.getUint32(i * 4, true))
  }
  const streamOf = (e: Entry): Uint8Array => {
    if (e.size < miniCutoff) {
      const ss = chain(e.start, miniFat)
      const out = new Uint8Array(ss.length * miniSize)
      ss.forEach((s, i) => out.set(ministream.subarray(s * miniSize, (s + 1) * miniSize), i * miniSize))
      return out.subarray(0, e.size)
    }
    return readChain(e.start).subarray(0, e.size)
  }
  const out = new Map<string, Uint8Array>()
  const walk = (i: number, prefix: string, depth: number) => {
    if (i >= entries.length || i === 0xffffffff || depth > 64) return
    const e = entries[i]
    walk(e.left, prefix, depth + 1)
    walk(e.right, prefix, depth + 1)
    const path = prefix ? `${prefix}/${e.name}` : e.name
    if (e.type === 2) out.set(path, streamOf(e))
    if (e.type === 1) walk(e.child, path, depth + 1)
  }
  if (root) walk(root.child, '', 0)
  return out
}

/** MS-OVBA 2.4.1 decompression of a CompressedContainer. */
export function decompress(src: Uint8Array): Uint8Array {
  if (src[0] !== 1) throw new Error('Not a compressed container')
  const out: number[] = []
  let pos = 1
  while (pos + 2 <= src.length) {
    const header = src[pos] | (src[pos + 1] << 8)
    const size = (header & 0x0fff) + 3
    const compressed = (header & 0x8000) !== 0
    const end = Math.min(src.length, pos + size)
    pos += 2
    const chunkStart = out.length
    if (!compressed) {
      for (let i = 0; i < 4096 && pos < end; i++) out.push(src[pos++])
      continue
    }
    while (pos < end) {
      const flags = src[pos++]
      for (let bit = 0; bit < 8 && pos < end; bit++) {
        if (!(flags & (1 << bit))) {
          out.push(src[pos++])
          continue
        }
        const token = src[pos] | (src[pos + 1] << 8)
        pos += 2
        const d = out.length - chunkStart
        let bitCount = 4
        while (1 << bitCount < d) bitCount++
        bitCount = Math.max(4, Math.min(12, bitCount))
        const lengthMask = 0xffff >> bitCount
        const offset = (token >> (16 - bitCount)) + 1
        const length = (token & lengthMask) + 3
        for (let i = 0; i < length; i++) out.push(out[out.length - offset])
      }
    }
  }
  return new Uint8Array(out)
}

/** The VBA modules' names and source text. */
export function vbaModules(bin: Uint8Array): { name: string; code: string }[] {
  const streams = readCfb(bin)
  const dirKey = [...streams.keys()].find((k) => /(^|\/)VBA\/dir$/i.test(k))
  if (!dirKey) return []
  const base = dirKey.slice(0, -3)
  const dir = decompress(streams.get(dirKey)!)
  const dv = new DataView(dir.buffer)
  const latin = (b: Uint8Array) => new TextDecoder('windows-1252').decode(b)
  const modules: { name: string; stream: string; offset: number }[] = []
  let cur: { name: string; stream: string; offset: number } | null = null
  let codePage = 1252
  for (let p = 0; p + 6 <= dir.length; ) {
    const id = dv.getUint16(p, true)
    let size = dv.getUint32(p + 2, true)
    // PROJECTVERSION has a fixed 6-byte payload whatever its size field says.
    if (id === 0x0009) size = 6
    const data = dir.subarray(p + 6, p + 6 + size)
    if (id === 0x0003) codePage = dv.getUint16(p + 6, true)
    if (id === 0x0019) {
      cur = { name: latin(data), stream: latin(data), offset: 0 }
      modules.push(cur)
    }
    if (id === 0x001a && cur) cur.stream = latin(data)
    if (id === 0x0031 && cur) cur.offset = dv.getUint32(p + 6, true)
    p += 6 + size
  }
  const dec = (() => {
    try {
      return new TextDecoder(codePage === 65001 ? 'utf-8' : `windows-${codePage}`)
    } catch {
      return new TextDecoder('windows-1252')
    }
  })()
  return modules.flatMap((m) => {
    const s = streams.get(base + m.stream) ?? [...streams.entries()].find(([k]) => k.toLowerCase() === (base + m.stream).toLowerCase())?.[1]
    if (!s) return []
    try {
      const code = dec.decode(decompress(s.subarray(m.offset))).replace(/\r\n/g, '\n')
      // Hidden attribute lines (Attribute VB_Name = …) aren't shown in the editor.
      return [{ name: m.name, code: code.replace(/^Attribute VB_[^\n]*\n/gm, '').trimEnd() }]
    } catch {
      return []
    }
  })
}
