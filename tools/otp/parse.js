// Reads the regional OTP Excel exports (.xlsx) and builds the dashboard data, entirely in the browser.
// The files are 100-200 MB each with ~1 GB of sheet XML inside, so nothing is loaded whole:
// zip entries are read straight from the File with Blob.slice, inflated with the built-in
// DecompressionStream, and the XML is scanned row by row, keeping only the columns the dashboard uses.
// Mirrors build_otp_dashboard.py; runs in a Web Worker (worker.js) and also under Node for testing.

const NEEDED_BASE = { cust: 1, custName: 2, ship: 3, transit: 6, origZone: 8, destZone: 10, pod: 28, defect: 33, exception: 34, revenue: 53 };
const EXPECTED = { 0: 'HAWB Number', 1: 'Customer', 2: "Customer's Name", 3: 'Ship Date (8A)', 6: 'Transit Time', 8: 'Origin Zone',
  10: 'Destination Zone', 28: 'Delivery Date (8A)', 33: 'Defect Code', 34: 'Exception/Comments', 53: 'Total Revenue Charges' };
const ZONES_OK = new Set(['A', 'B', 'C', 'D', 'E']);
const DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30);

// ---------- small helpers ----------
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const unescapeXml = (t) => (t.indexOf('&') < 0 ? t : t.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
  e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1)) : (ENT[e] ?? m)));
function colIndex(letters) { let n = 0; for (let i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n - 1; }
const str = (v) => (v == null ? '' : typeof v === 'number' ? (Number.isInteger(v) ? String(v) : String(v)) : String(v));
const regionName = (file) => file.replace(/\.[^.]+$/, '').replace(/\++$/, '').replace(/_/g, '/').toUpperCase();

// Excel serial day <-> parts (UTC, so no time-zone drift)
const serialDate = (n) => new Date(EPOCH + Math.floor(n) * DAY);
function toSerial(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v > 0 ? Math.floor(v) : null;
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v);
  if (m) return Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / DAY);
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(v.trim());
  if (m) return Math.round((Date.UTC(+m[3], +m[1] - 1, +m[2]) - EPOCH) / DAY);
  return null;
}
const hasTime = (v) => typeof v === 'number' && !Number.isInteger(v);
const isWorkday = (serial, hol) => { const wd = serialDate(serial).getUTCDay(); return wd !== 0 && wd !== 6 && !hol.has(serial); };
function workday(start, n, hol) { let d = start; while (n > 0) { d++; if (isWorkday(d, hol)) n--; } return d; }   // Excel WORKDAY
function parts(serial) {
  const d = serialDate(serial), y = d.getUTCFullYear();
  const jan1 = new Date(Date.UTC(y, 0, 1)), yday = Math.round((d - jan1) / DAY) + 1;
  return { year: y, month: d.getUTCMonth() + 1, week: Math.floor((yday + (jan1.getUTCDay() + 6) % 7 - 1) / 7) + 1 };   // WEEKNUM(d,2)
}

// ---------- zip access ----------
async function bytes(blob, a, b) { return new Uint8Array(await blob.slice(a, b).arrayBuffer()); }

async function listZip(blob) {
  const tailLen = Math.min(blob.size, 65557 + 20), tail = await bytes(blob, blob.size - tailLen, blob.size);
  const dv = new DataView(tail.buffer);
  let p = -1;
  for (let i = tail.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { p = i; break; }
  if (p < 0) throw new Error('is not an .xlsx file (no zip directory found)');
  let cdSize = dv.getUint32(p + 12, true), cdOff = dv.getUint32(p + 16, true);
  if ((cdOff === 0xFFFFFFFF || cdSize === 0xFFFFFFFF) && p >= 20 && dv.getUint32(p - 20, true) === 0x07064b50) {
    const at = Number(dv.getBigUint64(p - 12, true)), z = new DataView((await bytes(blob, at, at + 56)).buffer);
    cdSize = Number(z.getBigUint64(40, true)); cdOff = Number(z.getBigUint64(48, true));
  }
  const cd = await bytes(blob, cdOff, cdOff + cdSize), cv = new DataView(cd.buffer), dec = new TextDecoder();
  const entries = new Map();
  for (let q = 0; q + 46 <= cd.length && cv.getUint32(q, true) === 0x02014b50;) {
    const method = cv.getUint16(q + 10, true), nlen = cv.getUint16(q + 28, true), xlen = cv.getUint16(q + 30, true), clen = cv.getUint16(q + 32, true);
    let csize = cv.getUint32(q + 20, true), usize = cv.getUint32(q + 24, true), off = cv.getUint32(q + 42, true);
    const name = dec.decode(cd.subarray(q + 46, q + 46 + nlen));
    for (let x = q + 46 + nlen, end = x + xlen; x + 4 <= end; x += 4 + cv.getUint16(x + 2, true)) {
      if (cv.getUint16(x, true) !== 1) continue;           // zip64 sizes
      let y = x + 4;
      if (usize === 0xFFFFFFFF) { usize = Number(cv.getBigUint64(y, true)); y += 8; }
      if (csize === 0xFFFFFFFF) { csize = Number(cv.getBigUint64(y, true)); y += 8; }
      if (off === 0xFFFFFFFF) off = Number(cv.getBigUint64(y, true));
    }
    entries.set(name, { method, csize, usize, off });
    q += 46 + nlen + xlen + clen;
  }
  return entries;
}

async function entryReader(blob, e, onBytes) {
  const lh = new DataView((await bytes(blob, e.off, e.off + 30)).buffer);
  const start = e.off + 30 + lh.getUint16(26, true) + lh.getUint16(28, true);
  let s = blob.slice(start, start + e.csize).stream();
  if (onBytes) s = s.pipeThrough(new TransformStream({ transform(ch, c) { onBytes(ch.byteLength); c.enqueue(ch); } }));
  if (e.method === 8) s = s.pipeThrough(new DecompressionStream('deflate-raw'));
  else if (e.method !== 0) throw new Error('uses an unsupported zip compression method');
  return s.pipeThrough(new TextDecoderStream()).getReader();
}
async function entryText(blob, e) {
  const r = await entryReader(blob, e); let t = '';
  for (;;) { const { value, done } = await r.read(); if (value) t += value; if (done) return t; }
}
// Feed complete pieces (ending at `sep`) to fn; return false from fn to stop early.
async function scan(reader, sep, fn) {
  let buf = '';
  for (;;) {
    const { value, done } = await reader.read();
    if (value) buf += value;
    const cut = done ? buf.length : buf.lastIndexOf(sep) + (buf.lastIndexOf(sep) >= 0 ? sep.length : 0);
    if (cut > 0) { const piece = buf.slice(0, cut); buf = buf.slice(cut); if (fn(piece) === false) { await reader.cancel(); return; } }
    if (done) return;
  }
}

// ---------- sheet XML ----------
const ROW_RE = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
const V_RE = /<v>([\s\S]*?)<\/v>/, F_RE = /<f\b[^>]*>([\s\S]*?)<\/f>/, T_RE = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
// A cell value: number, string, or {s: index} for a shared string (resolved after the shared-string pass)
function cellValue(attrs, inner) {
  const t = /\st="(\w+)"/.exec(attrs)?.[1];
  if (!inner) return null;
  if (t === 'inlineStr') { let s = '', m; T_RE.lastIndex = 0; while ((m = T_RE.exec(inner))) s += m[1]; return unescapeXml(s); }
  const v = V_RE.exec(inner); if (!v) return null;
  if (t === 's') return { s: +v[1] };
  if (t === 'str' || t === 'e') return unescapeXml(v[1]);
  if (t === 'b') return v[1] === '1' ? 1 : 0;
  if (t === 'd') return toSerial(v[1]);
  const n = parseFloat(v[1]); return isNaN(n) ? null : n;
}
// Calls onCell(rowNumber, colIndex, attrs, inner) for each cell in the piece. wantCol(c) filters cheaply.
function eachCell(piece, wantCol, onCell, onRowEnd) {
  ROW_RE.lastIndex = 0; let rm;
  while ((rm = ROW_RE.exec(piece))) {
    const r = +(/\sr="(\d+)"/.exec(rm[1])?.[1] || 0);
    if (rm[2]) {
      CELL_RE.lastIndex = 0; let cm, pos = -1;
      while ((cm = CELL_RE.exec(rm[2]))) {
        const ref = /\sr="([A-Z]+)\d*"/.exec(cm[1]);
        const c = ref ? colIndex(ref[1]) : pos + 1; pos = c;
        if (wantCol(c, r)) onCell(r, c, cm[1], cm[2]);
      }
    }
    if (onRowEnd && onRowEnd(r) === false) return false;
  }
}
// Small sheets (CSR, delay codes): whole grid as rows of values
async function readGrid(blob, path) {
  const rows = [];
  await scan(await entryReader(blob, path), '</row>', (piece) => {
    eachCell(piece, () => true, (r, c, a, inner) => { (rows[r - 1] ||= [])[c] = cellValue(a, inner); });
  });
  return rows;
}

// Shared strings: keep only the indices in `need` (a Set), so memory stays small.
async function readSharedStrings(blob, e, need) {
  const out = new Map(); if (!e || !need.size) return out;
  let idx = 0;
  await scan(await entryReader(blob, e), '</si>', (piece) => {
    let from = 0;
    for (;;) {
      const end = piece.indexOf('</si>', from); if (end < 0) break;
      if (need.has(idx)) {
        const seg = piece.slice(piece.indexOf('<si', from), end).replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
        let s = '', m; T_RE.lastIndex = 0; while ((m = T_RE.exec(seg))) s += m[1];
        out.set(idx, unescapeXml(s));
      }
      idx++; from = end + 5;
    }
  });
  return out;
}

// Column store for the big data sheet: one number slot and one string code per cell.
class Col {
  constructor(n) { this.num = new Float64Array(n).fill(NaN); this.code = new Int32Array(n).fill(-1); }
  grow(n) { const num = new Float64Array(n).fill(NaN), code = new Int32Array(n).fill(-1); num.set(this.num); code.set(this.code); this.num = num; this.code = code; }
}

// ---------- one workbook ----------
async function processWorkbook(file, agg, progress) {
  const name = file.name, region = regionName(name);
  const z = await listZip(file);
  const wbXml = await entryText(file, z.get('xl/workbook.xml') || (() => { throw new Error('is not an Excel workbook'); })());
  const relXml = await entryText(file, z.get('xl/_rels/workbook.xml.rels'));
  const rels = new Map(), sheets = [];
  for (const m of relXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = /\sId="([^"]+)"/.exec(m[1])?.[1], target = /\sTarget="([^"]+)"/.exec(m[1])?.[1];
    if (id && target) rels.set(id, target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, ''));
  }
  for (const m of wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const nm = unescapeXml(/\sname="([^"]*)"/.exec(m[1])?.[1] || ''), rid = /\sr:id="([^"]+)"/.exec(m[1])?.[1];
    sheets.push({ name: nm, path: rels.get(rid) });
  }
  const sstPath = [...rels.values()].find((p) => /sharedStrings\.xml$/i.test(p)) || 'xl/sharedStrings.xml';
  const sstEntry = z.get(sstPath);
  const entryOf = (sheetName) => { const s = sheets.find((x) => x.name === sheetName); return s && z.get(s.path); };
  const data = sheets[0] && z.get(sheets[0].path);
  if (!data) throw new Error('has no worksheets');
  const csrE = entryOf('CSR'), delayE = entryOf('delay codes in WP');
  if (!csrE) throw new Error('has no "CSR" tab');
  if (!delayE) throw new Error('has no "delay codes in WP" tab');

  // pass 1: header row + row-2 formulas, and the two lookup tabs
  progress({ file: name, stage: 'Reading lookup tabs', pct: 0 });
  const head = [], f2 = [];
  await scan(await entryReader(file, data), '</row>', (piece) => eachCell(piece, (c, r) => r <= 2, (r, c, a, inner) => {
    if (r === 1) head[c] = cellValue(a, inner);
    else { const f = inner && F_RE.exec(inner); if (f) f2[c] = unescapeXml(f[1]); }
  }, (r) => r < 2));
  const csrGrid = await readGrid(file, csrE), delayGrid = await readGrid(file, delayE);
  const need1 = new Set();
  const noteS = (v) => { if (v && typeof v === 'object') need1.add(v.s); };
  head.forEach(noteS); csrGrid.forEach((r) => r && r.forEach(noteS)); delayGrid.forEach((r) => r && r.forEach(noteS));
  progress({ file: name, stage: 'Reading lookup tabs', pct: 2 });
  let sst = await readSharedStrings(file, sstEntry, need1);
  const val = (v) => (v && typeof v === 'object' ? (sst.get(v.s) ?? '') : v);
  const header = []; for (let i = 0; i < head.length; i++) header[i] = str(val(head[i])).trim();
  for (const [i, nm] of Object.entries(EXPECTED)) if (header[i] !== nm) throw new Error(`is not an OTP export (column ${+i + 1} is "${header[i] || ''}", expected "${nm}")`);
  const low = header.map((h) => h.toLowerCase()), find = (n) => low.indexOf(n);

  const lookupCol = (key, dflt) => { const m = /CSR!\$?[A-Z]+:\$?[A-Z]+\s*,\s*(\d+)/.exec(f2[find(key)] || ''); return m ? +m[1] - 1 : dflt; };
  const accCol = lookupCol('accounts', 1), csrCol = lookupCol('csr', 3);
  const hm = /CSR!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/.exec(f2[find('due date adjusted')] || '');
  const csrMap = new Map();
  for (const r of csrGrid.slice(1)) {
    if (!r) continue;
    const k = str(val(r[0])).toUpperCase(); if (!k || csrMap.has(k)) continue;
    csrMap.set(k, [str(val(r[accCol])).trim() || '(blank)', str(val(r[csrCol])).trim() || '(blank)']);
  }
  const hol = new Set();
  if (hm) { const c = colIndex(hm[1]); for (let r = +hm[2] - 1; r <= +hm[4] - 1; r++) { const s = toSerial(val(csrGrid[r]?.[c])); if (s) hol.add(s); } }
  const ctrl = new Map();
  for (const r of delayGrid.slice(1)) { if (!r) continue; const k = str(val(r[0])).toUpperCase(); if (k && !ctrl.has(k)) ctrl.set(k, str(val(r[2])).trim()); }

  // pass 2: the data rows, only the columns we use
  const cachedIdx = { accounts: find('accounts'), csr: find('csr'), onTime: find('on-time'), netLate: find('net_late') };
  const slots = { ...NEEDED_BASE }; for (const [k, i] of Object.entries(cachedIdx)) if (i >= 0) slots['c_' + k] = i;
  const slotOfCol = new Map(Object.entries(slots).map(([k, c]) => [c, k]));
  let cap = 1 << 16; const cols = {}; for (const k of Object.keys(slots)) cols[k] = new Col(cap);
  let hasA = new Uint8Array(cap), maxRow = 0;
  const inline = [], inlineId = new Map();
  const intern = (s) => { let id = inlineId.get(s); if (id === undefined) { id = inline.length; inline.push(s); inlineId.set(s, id); } return -2 - id; };
  let read = 0; const total = data.csize; let lastPct = -1;
  const reader = await entryReader(file, data, (n) => {
    read += n; const pct = 3 + Math.floor((read / total) * 92);
    if (pct !== lastPct) { lastPct = pct; progress({ file: name, stage: 'Reading shipments', pct }); }
  });
  await scan(reader, '</row>', (piece) => eachCell(piece, (c, r) => r >= 2 && (c === 0 || slotOfCol.has(c)), (r, c, a, inner) => {
    const i = r - 2;
    if (i >= cap) { let n = cap; while (n <= i) n *= 2; for (const k in cols) cols[k].grow(n); const h = new Uint8Array(n); h.set(hasA); hasA = h; cap = n; }
    if (i + 1 > maxRow) maxRow = i + 1;
    const v = cellValue(a, inner);
    if (c === 0) { if (v != null && v !== '') hasA[i] = 1; if (!slotOfCol.has(0)) return; }
    const col = cols[slotOfCol.get(c)];
    if (v == null) return;
    if (typeof v === 'number') col.num[i] = v;
    else if (typeof v === 'object') col.code[i] = v.s;
    else if (v !== '') col.code[i] = intern(v);
  }));

  progress({ file: name, stage: 'Calculating', pct: 96 });
  const need2 = new Set(); for (const k in cols) { const cd = cols[k].code; for (let i = 0; i < maxRow; i++) if (cd[i] >= 0) need2.add(cd[i]); }
  sst = await readSharedStrings(file, sstEntry, need2);
  const raw = (k, i) => { const col = cols[k]; if (!col) return null; const cd = col.code[i]; if (cd >= 0) return sst.get(cd) ?? ''; if (cd <= -2) return inline[-2 - cd]; const n = col.num[i]; return isNaN(n) ? null : n; };
  const label = (k, i) => { const v = str(raw(k, i)).trim(); return !v || v[0] === '#' || v === '0' ? '' : v; };

  let rows = 0; const mism = {}, anomalies = {};
  const bump = (o, k) => (o[k] = (o[k] || 0) + 1);
  for (let i = 0; i < maxRow; i++) {
    const custRaw = str(raw('cust', i));
    if (!hasA[i] && custRaw === '') continue;
    const cust = custRaw.toUpperCase(), map = csrMap.get(cust);
    const account = label('c_accounts', i) || (map ? map[0] : '#N/A (not on CSR tab)');
    const csr = label('c_csr', i) || (map ? map[1] : '#N/A (not on CSR tab)');
    const ship = toSerial(raw('ship', i)), podRaw = raw('pod', i), pod = toSerial(podRaw);
    const slaTxt = str(raw('transit', i)).slice(0, 1), slaErr = !/^\d$/.test(slaTxt);
    const oz = str(raw('origZone', i)).toUpperCase(), dz = str(raw('destZone', i)).toUpperCase();
    let status;
    if (pod == null) status = 'No Pod Yet';
    else if (slaErr) { status = 'Error'; bump(anomalies, 'SLA not numeric (Transit Time) with POD'); }
    else if (ship == null) { status = 'Late'; bump(anomalies, 'POD but no ship date'); }
    else {
      const slaAdj = +slaTxt + (ZONES_OK.has(dz) ? 0 : 1) + (ZONES_OK.has(oz) ? 0 : 1);
      const due = workday(ship, slaAdj, hol);
      // Excel compares the full POD timestamp: on the due date but with a time of day is "Late"
      status = pod < due || (pod === due && !hasTime(podRaw)) ? 'On-Time' : 'Late';
    }
    const defect = str(raw('defect', i));
    const carrier = defect === '' ? 'No Delay Code' : (ctrl.get(defect.toUpperCase()) ?? '#N/A');
    if (status === 'Late' && carrier === '#N/A') bump(anomalies, 'Late with defect code not in delay-code table');
    const exc = str(raw('exception', i)).trim();
    const delay = status === 'On-Time' ? 'On-Time' : (exc || 'Missing Delay Code');
    const netLate = status === 'Late' && carrier === 'Controllable';
    const rv = raw('revenue', i), rev = typeof rv === 'number' ? rv : 0;
    const cOn = str(raw('c_onTime', i));
    if ((cOn === 'On-Time' || cOn === 'Late' || cOn === 'No Pod Yet') && cOn !== status) bump(mism, 'On-Time');
    const cNet = raw('c_netLate', i);
    if (typeof cNet === 'number' && Boolean(cNet) !== netLate) bump(mism, 'Net_Late');
    agg.add(region, str(raw('custName', i)).trim() || cust, account, csr, ship, pod, status, delay, carrier === 'Controllable', netLate, rev);
    rows++;
  }
  progress({ file: name, stage: 'Done', pct: 100, rows });
  return { region, rows, mismatches: mism, anomalies, holidays: hol.size };
}

// ---------- roll-up (same shape as build_otp_dashboard.py output) ----------
class Aggregator {
  constructor() { this.dims = { region: new Map(), cust: new Map(), account: new Map(), csr: new Map(), delay: new Map() }; this.pod = new Map(); this.ship = new Map(); this.delay = new Map(); }
  di(kind, v) { const d = this.dims[kind]; let i = d.get(v); if (i === undefined) { i = d.size; d.set(v, i); } return i; }
  add(region, cust, account, csr, ship, pod, status, delay, controllable, netLate, rev) {
    const base = [this.di('region', region), this.di('cust', cust), this.di('account', account), this.di('csr', csr)];
    if (ship != null) { const p = parts(ship), k = [...base, p.year, p.week, p.month].join(','); this.ship.set(k, (this.ship.get(k) || 0) + 1); }
    if (pod != null && (status === 'On-Time' || status === 'Late')) {
      const p = parts(pod), k = [...base, p.year, p.week, p.month].join(',');
      let a = this.pod.get(k); if (!a) this.pod.set(k, (a = [0, 0, 0, 0]));
      a[0]++; if (status === 'Late') a[1]++; if (netLate) a[2]++; a[3] += rev;
      if (status === 'Late') { const dk = k + ',' + this.di('delay', delay) + ',' + (controllable ? 1 : 0); this.delay.set(dk, (this.delay.get(dk) || 0) + 1); }
    }
  }
  result(sources, checks) {
    const keys = (k) => k.split(',').map(Number), pad = (n) => String(n).padStart(2, '0'), d = new Date();
    return {
      generated: `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`,
      sources, checks,
      dims: Object.fromEntries(Object.entries(this.dims).map(([k, m]) => [k, [...m.keys()]])),
      pod: [...this.pod].map(([k, a]) => [...keys(k), a[0], a[1], a[2], Math.round(a[3] * 100) / 100]),
      ship: [...this.ship].map(([k, n]) => [...keys(k), n]),
      delay: [...this.delay].map(([k, n]) => [...keys(k), n]),
    };
  }
}

// files: File/Blob objects with .name. progress({file, stage, pct}) is called as work advances.
export async function buildFromFiles(files, progress = () => {}) {
  const agg = new Aggregator(), checks = {}, sources = [], skipped = [];
  for (const f of files) {
    if (!/\.xlsx$/i.test(f.name)) { skipped.push({ file: f.name, reason: 'is not an .xlsx file' }); progress({ file: f.name, stage: 'Skipped: not an .xlsx file', pct: 100, skipped: true }); continue; }
    try {
      const r = await processWorkbook(f, agg, progress);
      checks[r.region] = { rows: r.rows, mismatches: r.mismatches, anomalies: r.anomalies };
      sources.push(f.name);
    } catch (e) {
      skipped.push({ file: f.name, reason: e.message });
      progress({ file: f.name, stage: 'Skipped: ' + e.message, pct: 100, skipped: true });
    }
  }
  if (!sources.length) throw new Error(skipped.length ? `None of the files could be read. ${skipped[0].file} ${skipped[0].reason}.` : 'No files were chosen.');
  return Object.assign(agg.result(sources, checks), { skipped });
}
