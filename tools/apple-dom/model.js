// Core of the apple-dom tool: reads the APPLE_AMR export, computes the weekly OTD numbers and delay-code
// Paretos, writes them into the "Apple Dom WK# Charts Template" workbook, and builds the PPT tables.
// Pure module (no DOM), so it runs in the page's worker and under Node for testing.
import { openBook, listZip, entryReader } from '../otp/parse.js';

// ---------- Apple fiscal calendar ----------
// Weeks run Sunday to Saturday; the fiscal year ends on the last Saturday of September
// (FY27 = Sep 27 2026 - Sep 25 2027). Quarters are 13 weeks; in a 53-week year Q1 has 14.
const DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30);
export function fyEnd(fy) { const d = Date.UTC(fy, 8, 30); return d - ((new Date(d).getUTCDay() + 1) % 7) * DAY; }
const fyStart = (fy) => fyEnd(fy - 1) + DAY;
export const weeksIn = (fy) => Math.round((fyEnd(fy) - fyStart(fy) + DAY) / (7 * DAY));
export function fiscalOf(ms) {
  const y = new Date(ms).getUTCFullYear(), fy = ms > fyEnd(y) ? y + 1 : y;
  return { fy, week: Math.floor((ms - fyStart(fy)) / (7 * DAY)) + 1 };
}
export function quarterOf(fy, week) {
  const starts = weeksIn(fy) === 53 ? [1, 15, 28, 41] : [1, 14, 27, 40];
  let q = 0; while (q < 3 && week >= starts[q + 1]) q++;
  const end = q < 3 ? starts[q + 1] - 1 : weeksIn(fy);
  return { q: q + 1, first: starts[q], last: end, qweek: week - starts[q] + 1, len: end - starts[q] + 1 };
}
export const weekRange = (fy, week) => { const s = fyStart(fy) + (week - 1) * 7 * DAY; return [s, s + 6 * DAY]; };
// Date of the export: the YYYYMMDD stamp in its name (APPLE_AMR_INBOUND_PROCESSED_20261007_092200), else today.
export function fileDate(name) {
  const m = /(20\d\d)(\d\d)(\d\d)/.exec(name || '');
  if (m) { const t = Date.UTC(+m[1], +m[2] - 1, +m[3]); if (!isNaN(t)) return t; }
  const n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate());
}

// ---------- reading the APPLE_AMR export ----------
const FIELDS = { oem: 'Oem', tl: 'TL?', otp: 'OTP', wk: 'Due WK', cause: 'Cause', code: 'Delay Code', q: 'Piece Quantity',
  hawb: 'Intl Hawb', comments: 'Comments', origin: 'Origin City', dzip: 'Destination Zip Code', dname: 'Destination Name' };
const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
export const looksLikeAmrName = (name) => /apple[\s_-]*amr/i.test(name || '');

async function resolve(book, vals) {
  const need = new Set();
  for (const v of vals) if (v && typeof v === 'object') need.add(v.s);
  const sst = need.size ? await book.sharedStrings(need) : new Map();
  return (v) => (v && typeof v === 'object' ? sst.get(v.s) ?? '' : v);
}

// Returns { recs, name } where recs are the rows with an Oem. onBytes(n) reports compressed bytes read.
export async function readAmr(file, name, onBytes) {
  const book = await openBook(file);
  const sheet = book.sheets.find((s) => /^data$/i.test(s.name.trim())) || book.sheets[0];
  const e = sheet && book.entryOf(sheet.name);
  if (!e) throw new Error('has no worksheets');
  const head = [];
  await book.cells(e, { want: (c, r) => r === 1, onCell: (r, c, v) => { if (r === 1) head[c] = v; }, onRowEnd: (r) => r < 1 });
  const rs = await resolve(book, head), cols = {};
  head.forEach((v, c) => { const k = norm(rs(v)); if (k && !(k in cols)) cols[k] = c; });
  const at = {}, missing = [];
  for (const [f, h] of Object.entries(FIELDS)) { const c = cols[norm(h)]; if (c == null) missing.push(h); else at[f] = c; }
  if (missing.length) throw new Error(`does not look like the APPLE_AMR export (no ${missing.join(', ')} column${missing.length > 1 ? 's' : ''} on sheet "${sheet.name}")`);
  const byCol = new Map(Object.entries(at).map(([f, c]) => [c, f]));
  const raw = Object.fromEntries(Object.keys(FIELDS).map((f) => [f, []]));
  let n = 0;
  await book.cells(e, {
    want: (c, r) => r > 1 && byCol.has(c), onBytes,
    onCell: (r, c, v) => { raw[byCol.get(c)][r - 2] = v; if (r - 1 > n) n = r - 1; },
  });
  const all = []; for (const f in raw) for (const v of raw[f]) if (v && typeof v === 'object') all.push(v);
  const str = await resolve(book, all);
  const S = (f, i) => { const v = str(raw[f][i]); return v == null ? '' : String(v).trim(); };
  const recs = [];
  for (let i = 0; i < n; i++) {
    const oem = S('oem', i).toUpperCase(); if (!oem) continue;
    const wkRaw = str(raw.wk[i]), wk = typeof wkRaw === 'number' ? wkRaw : /^\d+(\.0+)?$/.test(String(wkRaw ?? '').trim()) ? parseFloat(wkRaw) : null;
    const q = str(raw.q[i]), qn = typeof q === 'number' ? q : parseFloat(q);
    recs.push({ oem, tl: S('tl', i).toUpperCase(), late: norm(S('otp', i)) === 'late', wk, carrier: norm(S('cause', i)) === 'carrier',
      code: S('code', i), q: isFinite(qn) ? qn : 0, hawb: S('hawb', i), comments: S('comments', i), origin: S('origin', i),
      dzip: S('dzip', i), dname: S('dname', i) });
  }
  // HAWB view: first row of each Intl Hawb, like Excel's Remove Duplicates on the WK#-HAWB copy
  const seen = new Set();
  for (const r of recs) { const k = r.hawb.toUpperCase(); r.firstHawb = !!k && !seen.has(k); if (k) seen.add(k); }
  return { recs, name };
}

// ---------- numbers ----------
const AC = new Set(['ACN', 'ACP']);
// A template tab's scope: "Overall(...)" tabs cover every OEM (NO AC drops ACN/ACP); other tabs name their OEMs, e.g. "LO55&5660", "ACN,ACP".
export function tabScope(tab) {
  const t = tab.trim().toUpperCase();
  if (/^OVERALL/.test(t)) return { overall: true, noAC: /NO\s*AC/.test(t), hawb: /HAWB/.test(t), oems: null };
  return { overall: false, noAC: false, hawb: false, oems: new Set(t.split(/[,&]/).map((s) => s.trim()).filter(Boolean)) };
}
const inScope = (sc) => (r) => (sc.oems ? sc.oems.has(r.oem) : !(sc.noAC && AC.has(r.oem)));
const tlIs = (tl) => (tl === 'FTL' ? (r) => r.tl === 'FTL' : tl === 'LTL' ? (r) => r.tl === 'NO' : () => true);

export function otd(recs, { wk, weeks, scope, tl, hawb }) {
  const f = inScope(scope), g = tlIs(tl), ws = weeks ? new Set(weeks) : null, o = { due: 0, late: 0, carrier: 0 };
  for (const r of recs) {
    if (ws ? !ws.has(r.wk) : r.wk !== wk) continue;
    if (!f(r) || !g(r) || (hawb && !r.firstHawb)) continue;
    const v = hawb ? 1 : r.q;
    o.due += v; if (r.late) { o.late += v; if (r.carrier) o.carrier += v; }
  }
  return o;
}
export function pareto(recs, { wk, scope, tl, hawb }) {
  const f = inScope(scope), g = tlIs(tl), m = new Map(); let blank = 0;
  for (const r of recs) {
    if (r.wk !== wk || !r.late || !f(r) || !g(r) || (hawb && !r.firstHawb)) continue;
    const v = hawb ? 1 : r.q;
    if (!r.code) { blank += v; continue; }
    m.set(r.code, (m.get(r.code) || 0) + v);
  }
  const items = [...m].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  return { items, total: items.reduce((s, x) => s + x[1], 0), blank };
}
export const pct = (o, adj) => (o.due ? 1 - (adj ? o.carrier : o.late) / o.due : null);

// Weeks present in the export and the default one: the latest fiscal week that ended before the export date.
export function weekChoices(recs, name) {
  const date = fileDate(name), today = fiscalOf(date);
  const nums = [...new Set(recs.map((r) => r.wk).filter((w) => w != null && w >= 1 && w <= 53))];
  // Due WK has no year: use the fiscal year that puts the week closest to the export date (W50-52 in early October are last year's)
  const fyOf = (w) => (Math.abs(weekRange(today.fy, w)[0] - date) <= Math.abs(weekRange(today.fy - 1, w)[0] - date) ? today.fy : today.fy - 1);
  const list = nums.map((w) => ({ wk: w, fy: fyOf(w) })).sort((a, b) => b.fy - a.fy || b.wk - a.wk);
  let def = today.week > 1 ? { wk: today.week - 1, fy: today.fy } : { wk: weeksIn(today.fy - 1), fy: today.fy - 1 };
  if (!list.some((x) => x.wk === def.wk)) def = list.find((x) => x.fy < def.fy || (x.fy === def.fy && x.wk <= def.wk)) || list[0];
  return { list, def };
}

// ---------- PowerPoint tables ----------
const OEM_TABS = ['6151', 'ACN,ACP', 'LO65', 'LO55&5660', 'LO45', 'LO22', 'RL03', 'RL06', 'RL01', 'V135'];
const shortCode = (code) => code.replace(/^[^-]*-\s*/, '').replace(/\s*\(.*$/, '').trim() || code;
// Consignee brand for the BK/HB split: the name up to its site part (" DC", " - ", ".COM", " C/O " ...), compared without
// punctuation so "X-Y DC 12" and "XY DC - 34" group together. Shown in title case; 2-3 letter names stay upper case.
const SITE_SEPS = [' C/O ', ' ATTN', '/ATT', ':', ' RDC', ' DC', ' - ', '.COM', ' #', ' (', ',', ' DRY', ' ECOMMERCE'];
export function brandOf(name) {
  let s = String(name || '').toUpperCase().replace(/\s+/g, ' ').trim();
  for (const sep of SITE_SEPS) { const i = s.indexOf(sep); if (i > 0) s = s.slice(0, i); }
  return s.replace(/[\s\-\/,:]+$/, '').trim() || '(blank)';
}
const brandKey = (b) => b.replace(/[^A-Z0-9]/g, '') || b;
const titleCase = (b) => (/^[A-Z]{1,3}$/.test(b) ? b : b.split(' ').map((w) => (/\.|\d/.test(w) || w.length <= 2 ? w
  : w.split('-').map((x) => x.charAt(0) + x.slice(1).toLowerCase()).join('-'))).join(' '));
// Apple origin cities to airport codes (from "Commonly used terms.docx")
const ORIGIN = { CLAYTON: 'IND', LEBANON: 'BNA', CARLISLE: 'MDT', SPARKS: 'RNO', RIALTO: 'ONT', PLAINFIELD: 'ONT' };
const originCode = (city) => ORIGIN[String(city || '').trim().toUpperCase()] || String(city || '').trim();
const station = (z) => String(z || '').trim().replace(/^DS-/i, '');

// qtd: quarter-to-date totals from the filled template ({ units, unitsNoAC, hawb, hawbNoAC }); without a template QTD is left blank.
export function pptTables(recs, sel, qtd = null) {
  const ALL = { oems: null }, NOAC = { oems: null, noAC: true }, wk = sel.wk;
  const row = (label, sc, opt = {}) => { const o = otd(recs, { wk, scope: sc, ...opt }); return { label, ...o }; };
  // Slide 4: units, FTL + LTL
  const s4 = [row('Overall (WITH AC)', ALL), row('Overall (NO AC)', NOAC)];
  const has5660 = recs.some((r) => r.oem === '5660' && r.wk === wk);
  for (const t of ['LO22', 'LO55&5660', 'LO45', 'LO65', 'RL03', 'RL01', 'RL06', '6151', 'V135']) {
    const r = row(t === 'LO55&5660' && !has5660 ? 'LO55' : t, tabScope(t)); if (r.due) s4.push(r);
  }
  const known = new Set(OEM_TABS.flatMap((t) => [...tabScope(t).oems]));
  const other = row('Other OEMs', { oems: new Set(recs.filter((r) => !known.has(r.oem)).map((r) => r.oem)) });
  if (other.due) s4.push(other);
  s4.push(row('AppleCare', tabScope('ACN,ACP')));
  s4.push({ label: 'QTD', ...(qtd?.units || { due: null, late: null, carrier: null }) });
  // Slide 6 / Pivot-Hawb J:O: LTL HAWBs
  const H = { tl: 'LTL', hawb: true };
  const s6 = [row('Overall(WITH AC)', ALL, H), row('Overall(NO AC)', NOAC, H)];
  for (const t of OEM_TABS) s6.push(row(t === 'ACN,ACP' ? 'Applecre(ACP,ACN)' : t, tabScope(t), H));
  s6.push({ label: 'QTD', ...(qtd?.hawb || { due: null, late: null, carrier: null }) });
  // Slide 5: carrier delays. A1/C by lane (top 10 + Others); every other carrier code grouped by comment.
  const car = recs.filter((r) => r.wk === wk && r.late && r.carrier);
  const lane = (o, d) => `${originCode(o)}-${station(d)}`;
  const sum = (rows) => ({ hawb: new Set(rows.map((r) => r.hawb)).size, units: rows.reduce((s, r) => s + r.q, 0) });
  const groupBy = (rows, key) => { const m = new Map(); for (const r of rows) { const k = key(r); (m.get(k) || m.set(k, []).get(k)).push(r); } return m; };
  const isA1 = (r) => /^A1\b/i.test(r.code);
  const s5 = [];
  const a1 = car.filter(isA1);
  if (a1.length) {
    s5.push({ head: true, label: a1[0].code, lane: '', ...sum(a1) });
    const lanes = [...groupBy(a1, (r) => lane(r.origin, r.dzip))].map(([k, rows]) => ({ label: shortCode(a1[0].code), lane: k, rows, ...sum(rows) })).sort((a, b) => b.units - a.units);
    lanes.slice(0, 10).forEach((x) => s5.push(x));
    if (lanes.length > 10) { const rest = lanes.slice(10).flatMap((x) => x.rows); s5.push({ label: 'Others', lane: `${lanes.length - 10} lanes`, ...sum(rest) }); }
  }
  const oth = car.filter((r) => !isA1(r));
  if (oth.length) {
    const codes = [...new Set(oth.map((r) => r.code))];
    const p1 = codes.find((c) => /^P1\b/i.test(c));
    s5.push({ head: true, label: p1 ? (codes.length > 1 ? `${p1} & OTHER CARRIER DELAY` : p1) : 'OTHER CARRIER DELAY', lane: '', ...sum(oth) });
    const items = [...groupBy(oth, (r) => norm(r.comments) || '#' + r.code)].map(([, rows]) => {
      const o = [...new Set(rows.map((r) => originCode(r.origin)))].join('/'), d = [...new Set(rows.map((r) => station(r.dzip)))].join('/');
      return { label: rows[0].comments.replace(/\s+/g, ' ').trim() || shortCode(rows[0].code), lane: `${o}-${d}`, ...sum(rows) };
    }).sort((a, b) => b.units - a.units);
    items.forEach((x) => s5.push(x));
  }
  // Slide 8: LTL Pareto (units) and the BK/HB consignee split by brand, with and without AppleCare
  const s8 = (sc) => {
    const keep = inScope(sc), bk = recs.filter((r) => r.wk === wk && r.late && r.tl === 'NO' && keep(r) && /^(BK|HB)\b/i.test(r.code));
    const brands = [...groupBy(bk, (r) => brandKey(brandOf(r.dname)))].map(([, rows]) => ({
      brand: titleCase(brandOf([...groupBy(rows, (r) => r.dname)].sort((x, y) => y[1].reduce((t, r) => t + r.q, 0) - x[1].reduce((t, r) => t + r.q, 0))[0][0])), units: rows.reduce((s, r) => s + r.q, 0),
      names: [...groupBy(rows, (r) => r.dname)].map(([n, rr]) => [n, rr.reduce((s, r) => s + r.q, 0)]).sort((a, b) => b[1] - a[1]),
    })).sort((a, b) => b.units - a.units);
    return { pareto: pareto(recs, { wk, scope: sc, tl: 'LTL' }), bkTotal: bk.reduce((s, r) => s + r.q, 0), brands };
  };
  return { s4, s5, s6, s8: { withAC: s8(ALL), noAC: s8(NOAC) } };
}

// ---------- xlsx editing (template) ----------
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const unx = (t) => t.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENT[e] ?? m);
const xesc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const colIdx = (L) => { let n = 0; for (const ch of L) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
export const colL = (n) => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const attr = (a, k) => new RegExp(`\\s${k}="([^"]*)"`).exec(a)?.[1];
const setAttr = (a, k, v) => (attr(a, k) != null ? a.replace(new RegExp(`\\s${k}="[^"]*"`), v == null ? '' : ` ${k}="${v}"`) : v == null ? a : a + ` ${k}="${v}"`);

async function entryBytes(blob, e) {
  const r = await entryReader(blob, e, null, true), parts = []; let n = 0;
  for (;;) { const { value, done } = await r.read(); if (value) { parts.push(value); n += value.length; } if (done) break; }
  const out = new Uint8Array(n); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out;
}
async function unzip(blob) {
  const list = await listZip(blob), files = new Map();
  for (const [name, e] of list) files.set(name, await entryBytes(blob, e));
  return files;
}
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = (b) => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
async function deflate(b) { return new Uint8Array(await new Response(new Blob([b]).stream().pipeThrough(new CompressionStream('deflate-raw'))).arrayBuffer()); }
async function zip(files) {
  const enc = new TextEncoder(), parts = [], cd = []; let off = 0;
  const d = new Date(), dt = ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF, dd = (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF;
  for (const [name, data] of files) {
    const nb = enc.encode(name), comp = await deflate(data), crc = crc32(data);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 8, true);
    h.setUint16(10, dt, true); h.setUint16(12, dd, true); h.setUint32(14, crc, true); h.setUint32(18, comp.length, true); h.setUint32(22, data.length, true);
    h.setUint16(26, nb.length, true); h.setUint16(28, 0, true);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 8, true);
    c.setUint16(12, dt, true); c.setUint16(14, dd, true); c.setUint32(16, crc, true); c.setUint32(20, comp.length, true); c.setUint32(24, data.length, true);
    c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
    parts.push(new Uint8Array(h.buffer), nb, comp); cd.push(new Uint8Array(c.buffer), nb);
    off += 30 + nb.length + comp.length;
  }
  const cdSize = cd.reduce((s, x) => s + x.length, 0), e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.size, true); e.setUint16(10, files.size, true); e.setUint32(12, cdSize, true); e.setUint32(16, off, true);
  return new Blob([...parts, ...cd, new Uint8Array(e.buffer)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// Shift the relative references in a formula (for expanding shared formulas). String literals are left alone.
function shiftFormula(f, dr, dc) {
  return f.split(/("(?:[^"]|"")*")/).map((part, i) => i % 2 ? part : part.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(!])/g, (m, ca, c, ra, r, at, s) => {
    const prev = s[at - 1]; if (prev && /[A-Za-z_.]/.test(prev)) return m;
    return ca + (ca ? c : colL(colIdx(c) + dc)) + ra + (ra ? r : +r + dr);
  })).join('');
}

// A worksheet whose <sheetData> is held as rows/cells that can be edited and written back.
class Sheet {
  constructor(xml, sst) {
    this.sst = sst;
    const a = xml.indexOf('<sheetData'), selfClose = /^<sheetData\s*\/>/.test(xml.slice(a));
    const open = xml.indexOf('>', a) + 1, close = selfClose ? open : xml.indexOf('</sheetData>', open);
    this.pre = xml.slice(0, a) + '<sheetData>'; this.post = '</sheetData>' + xml.slice(selfClose ? open : close + 12);
    this.rows = new Map();
    const body = selfClose ? '' : xml.slice(open, close);
    for (const rm of body.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const r = +attr(rm[1], 'r'), cells = new Map();
      for (const cm of (rm[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) cells.set(colIdx(/[A-Z]+/.exec(attr(cm[1], 'r'))[0]), { a: cm[1], x: cm[2] || '' });
      this.rows.set(r, { a: rm[1], cells });
    }
    this.expandShared();
  }
  cell(r, c) { return this.rows.get(r)?.cells.get(c); }
  text(r, c) {
    const k = this.cell(r, c); if (!k) return '';
    const t = attr(k.a, 't');
    if (t === 'inlineStr') return unx([...k.x.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => m[1]).join(''));
    const v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(k.x)?.[1]; if (v == null) return '';
    return t === 's' ? this.sst[+v] ?? '' : unx(v);
  }
  num(r, c) { const k = this.cell(r, c); if (!k || attr(k.a, 't')) return null; const v = /<v\b[^>]*>([^<]*)<\/v>/.exec(k.x)?.[1]; return v == null ? null : +v; }
  formula(r, c) { const f = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(this.cell(r, c)?.x || ''); return f ? unx(f[1]) : null; }
  // Every shared formula becomes a plain one, so any cell can be overwritten without breaking a group.
  expandShared() {
    const masters = new Map();
    for (const [r, row] of this.rows) for (const [c, k] of row.cells) {
      const m = /<f\b([^>]*?)>([\s\S]*?)<\/f>/.exec(k.x);
      if (m && attr(m[1], 't') === 'shared' && attr(m[1], 'ref')) masters.set(attr(m[1], 'si'), { r, c, f: unx(m[2]) });
    }
    for (const [r, row] of this.rows) for (const [c, k] of row.cells) {
      const m = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(k.x);
      if (!m || attr(m[1], 't') !== 'shared') continue;
      const ms = masters.get(attr(m[1], 'si')); if (!ms) continue;
      k.x = k.x.replace(m[0], `<f>${xesc(shiftFormula(ms.f, r - ms.r, c - ms.c))}</f>`);
    }
  }
  ensure(r, c, styleFrom) {
    let row = this.rows.get(r); if (!row) this.rows.set(r, (row = { a: ` r="${r}"`, cells: new Map() }));
    row.a = setAttr(row.a, 'spans', null);
    let k = row.cells.get(c);
    if (!k) { const s = styleFrom != null ? attr(this.cell(r, styleFrom)?.a || '', 's') : null; row.cells.set(c, (k = { a: ` r="${colL(c)}${r}"` + (s ? ` s="${s}"` : ''), x: '' })); }
    return k;
  }
  setNum(r, c, v, styleFrom) { const k = this.ensure(r, c, styleFrom); k.a = setAttr(k.a, 't', null); k.x = `<v>${v}</v>`; }
  setStr(r, c, s, styleFrom) { const k = this.ensure(r, c, styleFrom); k.a = setAttr(k.a, 't', 'inlineStr'); k.x = `<is><t xml:space="preserve">${xesc(s)}</t></is>`; }
  setFormula(r, c, f, styleFrom) { const k = this.ensure(r, c, styleFrom); k.a = setAttr(k.a, 't', null); k.x = `<f>${xesc(f)}</f>`; }
  clear(r, c) { const k = this.cell(r, c); if (k) { k.a = setAttr(k.a, 't', null); k.x = ''; } }
  // Find cells whose text matches re: [{r, c, text}]
  find(re) { const out = []; for (const [r, row] of this.rows) for (const c of row.cells.keys()) { const t = this.text(r, c).trim(); if (t && re.test(t)) out.push({ r, c, text: t }); } return out.sort((a, b) => a.r - b.r || a.c - b.c); }
  toXml() {
    const rows = [...this.rows].sort((a, b) => a[0] - b[0]).map(([, row]) => {
      const cells = [...row.cells].sort((a, b) => a[0] - b[0]).map(([, k]) => {
        // formulas recalculate on open, so drop their stale cached values
        let x = k.x, a = k.a;
        if (/<f\b/.test(x)) { x = x.replace(/<v\b[^>]*>[\s\S]*?<\/v>|<v\/>/g, ''); a = setAttr(a, 't', null); }
        return x ? `<c${a}>${x}</c>` : `<c${a}/>`;
      }).join('');
      return cells ? `<row${row.a}>${cells}</row>` : `<row${row.a}/>`;
    }).join('');
    return this.pre + rows + this.post;
  }
}

const dec = new TextDecoder(), enc = new TextEncoder();
const relsOf = (files, path) => {
  const rp = path.replace(/[^/]+$/, (f) => '_rels/' + f + '.rels'), x = files.get(rp); const out = new Map();
  if (!x) return out;
  const base = path.replace(/[^/]+$/, '');
  for (const m of dec.decode(x).matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    let t = attr(m[1], 'Target'); if (!t || attr(m[1], 'TargetMode') === 'External') continue;
    t = t.startsWith('/') ? t.slice(1) : base + t;
    const segs = []; for (const s of t.split('/')) { if (s === '..') segs.pop(); else if (s !== '.') segs.push(s); }
    out.set(attr(m[1], 'Id'), segs.join('/'));
  }
  return out;
};
const refRe = /^(?:'((?:[^']|'')+)'|([^!']+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/;
function parseRef(f) {
  const m = refRe.exec(unx(f).trim()); if (!m) return null;
  const sheet = (m[1] ?? m[2]).replace(/''/g, "'");
  return { sheet, c1: colIdx(m[3]), r1: +m[4], c2: m[5] ? colIdx(m[5]) : colIdx(m[3]), r2: m[6] ? +m[6] : +m[4] };
}
const fmtRef = (sheet, c1, r1, c2, r2) => xesc(`'${sheet.replace(/'/g, "''")}'!$${colL(c1)}$${r1}` + (c1 !== c2 || r1 !== r2 ? `:$${colL(c2)}$${r2}` : ''));
const chartTitle = (x) => unx([...(x.split('<c:plotArea>')[0]).matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('')).trim();

export async function isChartsTemplate(blob) {
  try {
    const z = await listZip(blob); if (!z.has('xl/workbook.xml')) return false;
    const book = await openBook(blob); return book.sheets.some((s) => /^overall/i.test(s.name.trim()));
  } catch (e) { return false; }
}

// Fill the template for the selected week. Returns { blob, tabs, warnings }.
// tabs: per sheet, the OTD blocks (13-week series after the update) and Paretos, for the page preview.
export async function fillTemplate(blob, recs, sel) {
  const files = await unzip(blob);
  const wbPath = 'xl/workbook.xml', wbXml = dec.decode(files.get(wbPath)), wbRels = relsOf(files, wbPath);
  const sstXml = files.has('xl/sharedStrings.xml') ? dec.decode(files.get('xl/sharedStrings.xml')) : '';
  const sst = [...sstXml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => unx([...m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join('')));
  // Only the selected week is written. Earlier weeks keep what was reported: the export covers about a month of
  // ship dates and only the current week has its delay causes filled, so recomputing older weeks would undercount them.
  const Q = quarterOf(sel.fy, sel.wk);
  const warnings = [], tabs = [];
  for (const m of wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const name = unx(attr(m[1], 'name') || ''), path = wbRels.get(attr(m[1], 'r:id'));
    if (!path || !files.has(path)) continue;
    const sh = new Sheet(dec.decode(files.get(path)), sst), scope = tabScope(name), hawb = scope.hawb;
    const tab = { name, blocks: [], paretos: [] };
    // --- OTD blocks: "Total Weekly (Units) Due" labels; left one is FTL, right one LTL
    const labels = sh.find(/^total weekly (units )?due$/i);
    const blocks = [];
    for (const L of labels) {
      if (!/late/i.test(sh.text(L.r + 1, L.c)) || !/carrier/i.test(sh.text(L.r + 2, L.c))) continue;
      let hr = null; for (let r = L.r - 1; r >= Math.max(1, L.r - 10); r--) if (/^w\s*\d+$/i.test(sh.text(r, L.c + 1))) { hr = r; break; }
      const find = (re) => { for (let r = (hr ?? L.r - 8) + 1; r < L.r; r++) if (re.test(sh.text(r, L.c).trim())) return r; return null; };
      blocks.push({ c: L.c, r: L.r, hr, unadj: find(/^unadj/i), adj: find(/^adj otd/i) });
    }
    blocks.sort((a, b) => a.c - b.c).splice(2);
    blocks.forEach((b, i) => { b.tl = i === 0 ? 'FTL' : 'LTL'; });
    if (blocks.length < 2) warnings.push(`${name}: could not find the FTL and LTL OTD tables`);
    for (const b of blocks) {
      const c0 = b.c + 1, series = [];
      if (b.hr) for (let i = 1; i <= Q.len; i++) sh.setStr(b.hr, b.c + i, 'W' + i, c0);
      for (let i = 1; i <= Q.len; i++) {
        const w = Q.first + i - 1, c = b.c + i, C = colL(c);
        if (w === sel.wk) {
          const o = otd(recs, { wk: w, scope, tl: b.tl, hawb });
          const was = { due: sh.num(b.r, c), late: sh.num(b.r + 1, c), carrier: sh.num(b.r + 2, c) };   // what the template had
          sh.setNum(b.r, c, o.due, c0); sh.setNum(b.r + 1, c, o.late, c0); sh.setNum(b.r + 2, c, o.carrier, c0);
          // OTD % cells: the template's formula; weeks with nothing due show 100% (typed as 1 in past reports)
          if (b.unadj) { if (!o.due) sh.setNum(b.unadj, c, 1, c0); else if (sh.formula(b.unadj, c) == null) sh.setFormula(b.unadj, c, `IFERROR(IF(${C}${b.r}=0,NA(),(1-(${C}${b.r + 1}/${C}${b.r}))),0)`, c0); }
          if (b.adj) { if (!o.due) sh.setNum(b.adj, c, 1, c0); else if (sh.formula(b.adj, c) == null) sh.setFormula(b.adj, c, `IFERROR(IF(${C}${b.r + 2}="",100%,IF(${C}${b.r}=0,NA(),(1-(${C}${b.r + 2}/${C}${b.r})))),0)`, c0); }
          series.push({ w, ...o, fromFile: true, was });
        } else {
          const due = sh.num(b.r, c), late = sh.num(b.r + 1, c), carrier = sh.num(b.r + 2, c);
          series.push(w <= sel.wk && due != null ? { w, due, late: late || 0, carrier: carrier || 0, fromFile: false } : { w, due: null });
        }
      }
      tab.blocks.push({ tl: b.tl, series, label: hawb ? 'HAWBs' : 'Units' });
    }
    // middle "ToTal / Late / Carrier" =SUM(FTL,LTL) cells point at the selected week
    if (blocks.length === 2) {
      const [F, L] = blocks, qc = Q.qweek;
      for (let k = 0; k < 3; k++) {
        const r = F.r + k, row = sh.rows.get(r); if (!row) continue;
        for (const c of row.cells.keys()) {
          const f = sh.formula(r, c), mm = f && /^SUM\(\$?[A-Z]+\$?(\d+),\$?[A-Z]+\$?(\d+)\)$/i.exec(f.replace(/\s/g, ''));
          if (mm && +mm[1] === r && +mm[2] === r) sh.setFormula(r, c, `SUM(${colL(F.c + qc)}${r},${colL(L.c + qc)}${r})`);
        }
      }
    }
    // --- charts on this sheet
    const sheetRels = relsOf(files, path);
    for (const dp of sheetRels.values()) {
      if (!/drawings\/drawing\d+\.xml$/.test(dp)) continue;
      for (const cp of relsOf(files, dp).values()) {
        if (!/charts\/chart\d+\.xml$/.test(cp) || !files.has(cp)) continue;
        let x = dec.decode(files.get(cp)); const title = chartTitle(x);
        const refs = [...x.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((mm) => parseRef(mm[1])).filter((r) => r && r.sheet === name);
        let changed = false;
        // Pareto series refs come as: header cell (series name), delay codes, values; e.g. C44, B45:B53, C45:C53
        if (/delay code/i.test(title) && refs.length >= 2 && refs[1].r1 === refs[0].r1 + 1) {
          const tl = /\bLTL\b/i.test(title) ? 'LTL' : /\bFTL\b/i.test(title) ? 'FTL' : null;
          const first = refs[0].r1 + 1, valC = refs[0].c1, catC = refs[1].c1;
          // Grand Total row: the one the Impact % formula divides by, else the next "Grand Total" label
          let gt = null; const f = sh.formula(first, valC + 1);
          if (f) for (const mm of f.matchAll(/\$?([A-Z]+)\$?(\d+)/g)) if (colIdx(mm[1]) === valC && +mm[2] > first) { gt = +mm[2]; break; }
          if (!gt) for (let r = first; r < first + 40; r++) if (/^grand total$/i.test(sh.text(r, catC).trim())) { gt = r; break; }
          if (!gt) { warnings.push(`${name}: "${title}" has no Grand Total row`); continue; }
          const p = pareto(recs, { wk: sel.wk, scope, tl, hawb }), slots = gt - first, was = [];
          for (let r = first; r < gt; r++) { const k = sh.text(r, catC).trim(), v = sh.num(r, valC); if (k && v != null && !/^grand total$/i.test(k)) was.push([k, v]); }
          let items = p.items;
          if (items.length > slots) { const keep = items.slice(0, slots - 1), rest = items.slice(slots - 1); items = [...keep, [`Other (${rest.length} codes)`, rest.reduce((s, z) => s + z[1], 0)]]; }
          for (let i = 0; i < slots; i++) {
            const r = first + i;
            if (i < items.length) { sh.setStr(r, catC, items[i][0], catC); sh.setNum(r, valC, items[i][1], valC); } else { sh.clear(r, catC); sh.clear(r, valC); }
          }
          sh.setStr(gt, catC, 'Grand Total', catC); sh.setNum(gt, valC, p.total, valC);
          const n = Math.max(items.length, 1);
          x = x.replace(/<c:f>([^<]*)<\/c:f>/g, (all, ff) => { const r = parseRef(ff); return r && r.sheet === name && r.r1 === first ? `<c:f>${fmtRef(name, r.c1, first, r.c2, first + n - 1)}</c:f>` : all; });
          changed = true;
          tab.paretos.push({ title, tl, items: p.items, total: p.total, blank: p.blank, unit: hawb ? 'HAWBs' : 'units', was });
        } else if (/otd/i.test(title)) {
          x = x.replace(/<c:f>([^<]*)<\/c:f>/g, (all, ff) => {
            const r = parseRef(ff); if (!r || r.sheet !== name) return all;
            const b = blocks.find((bb) => r.c1 === bb.c + 1 && r.r1 === r.r2 && r.r1 >= (bb.hr ?? bb.r - 8) && r.r1 <= bb.r + 2);
            if (!b) return all; changed = true;
            return `<c:f>${fmtRef(name, r.c1, r.r1, b.c + Q.qweek, r.r1)}</c:f>`;
          });
        }
        if (changed) {
          x = x.replace(/<c:(numCache|strCache)>[\s\S]*?<\/c:\1>/g, '');
          files.set(cp, enc.encode(x));
        }
      }
    }
    files.set(path, enc.encode(sh.toXml()));
    tabs.push(tab);
  }
  // Excel rebuilds the calculation chain and recalculates every formula when the file opens
  if (files.has('xl/calcChain.xml')) {
    files.delete('xl/calcChain.xml');
    const rp = 'xl/_rels/workbook.xml.rels';
    files.set(rp, enc.encode(dec.decode(files.get(rp)).replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, '')));
    files.set('[Content_Types].xml', enc.encode(dec.decode(files.get('[Content_Types].xml')).replace(/<Override\b[^>]*calcChain[^>]*\/>/g, '')));
  }
  let wb = wbXml;
  wb = /<calcPr\b/.test(wb) ? wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (mm, a, s) => `<calcPr${setAttr(a, 'fullCalcOnLoad', '1')}${s}>`) : wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  files.set(wbPath, enc.encode(wb));
  // quarter to date, from the template's weekly columns (Q34 / Q35 / Q36 on the Overall tabs)
  const qsum = (tab, tl) => {
    const t = tabs.find((x) => x.name.replace(/\s+/g, '').toUpperCase() === tab); if (!t) return null;
    const o = { due: 0, late: 0, carrier: 0 };
    for (const b of t.blocks) if (!tl || b.tl === tl) for (const s of b.series) if (s.due != null) { o.due += s.due; o.late += s.late || 0; o.carrier += s.carrier || 0; }
    return o;
  };
  const qtd = { units: qsum('OVERALL(WITHAC)'), unitsNoAC: qsum('OVERALL(NOAC)'), hawb: qsum('OVERALLHAWB(WITHAC)', 'LTL'), hawbAll: qsum('OVERALLHAWB(WITHAC)') };
  return { blob: await zip(files), tabs, warnings, quarter: Q, qtd };
}

export const outName = (sel) => `Apple Dom WK${sel.wk} Charts Template FY${String(sel.fy).slice(2)}.xlsx`;
