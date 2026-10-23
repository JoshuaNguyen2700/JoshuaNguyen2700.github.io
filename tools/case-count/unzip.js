// Reads .zip files in the browser, so a zipped folder of exports loads like the folder itself.
// Uses the browser's own DecompressionStream; nothing is uploaded. Handles stored and deflated entries
// (what Windows "Send to > Compressed folder" and most zip tools write), including ZIP64 archives.

const u16 = (v, o) => v.getUint16(o, true);
const u32 = (v, o) => v.getUint32(o, true);
const u64 = (v, o) => u32(v, o) + u32(v, o + 4) * 2 ** 32;

// Where the central directory starts and how many entries it lists.
function directory(v) {
  const stop = Math.max(0, v.byteLength - 22 - 65535);
  for (let i = v.byteLength - 22; i >= stop; i--) {
    if (u32(v, i) !== 0x06054b50) continue;
    let count = u16(v, i + 10), start = u32(v, i + 16);
    // ZIP64: the real numbers sit in a second end record, found through its locator just before this one
    if ((count === 0xffff || start === 0xffffffff) && i >= 20 && u32(v, i - 20) === 0x07064b50) {
      const e = u64(v, i - 12);
      if (u32(v, e) === 0x06064b50) { count = u64(v, e + 32); start = u64(v, e + 48); }
    }
    return { count, start };
  }
  throw new Error('it is not a zip file, or it is damaged');
}

// DOS date and time -> milliseconds (local time, as zip tools write it)
const dosTime = (d, t) => new Date(1980 + (d >> 9), ((d >> 5) & 15) - 1, d & 31, t >> 11, (t >> 5) & 63, (t & 31) * 2).getTime();

async function inflate(bytes) {
  if (typeof DecompressionStream === 'undefined') throw new Error('this browser cannot open zip files; unzip it first, or use a current Edge or Chrome');
  const s = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}

// The files in a zip as [{ name: 'folder/file.xls', file: File }]. Folders, encrypted entries and
// macOS resource forks are left out.
export async function unzip(blob) {
  const buf = await blob.arrayBuffer(), v = new DataView(buf);
  const { count, start } = directory(v);
  const out = [];
  for (let i = 0, p = start; i < count; i++) {
    if (u32(v, p) !== 0x02014b50) throw new Error('its file list is damaged');
    const flags = u16(v, p + 8), method = u16(v, p + 10), time = u16(v, p + 12), date = u16(v, p + 14);
    let size = u32(v, p + 20), at = u32(v, p + 42);
    const nLen = u16(v, p + 28), xLen = u16(v, p + 30), cLen = u16(v, p + 32);
    // (Windows PowerShell's Compress-Archive writes "folder\file", so backslashes count as folder breaks)
    const name = new TextDecoder(flags & 0x800 ? 'utf-8' : 'windows-1252').decode(new Uint8Array(buf, p + 46, nLen)).replace(/\\/g, '/');
    // ZIP64 extra field: the 64-bit values appear in this order, only for the fields set to 0xffffffff
    for (let x = p + 46 + nLen, end = x + xLen; x + 4 <= end; x += 4 + u16(v, x + 2)) {
      if (u16(v, x) !== 1) continue;
      let o = x + 4;
      if (u32(v, p + 24) === 0xffffffff) o += 8;
      if (size === 0xffffffff) { size = u64(v, o); o += 8; }
      if (at === 0xffffffff) at = u64(v, o);
    }
    p += 46 + nLen + xLen + cLen;
    if (name.endsWith('/') || /(^|\/)__MACOSX\//.test(name) || flags & 1) continue;
    const data = at + 30 + u16(v, at + 26) + u16(v, at + 28);
    const raw = new Uint8Array(buf, data, size);
    let bytes;
    if (method === 0) bytes = raw;
    else if (method === 8) bytes = await inflate(raw);
    else continue;   // other compression methods are rare; skip rather than fail the whole zip
    const base = name.split('/').pop();
    out.push({ name, file: new File([bytes], base, { lastModified: dosTime(date, time) }) });
  }
  return out;
}
