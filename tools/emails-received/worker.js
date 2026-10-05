// Background worker for the emails-received tool: reads Salesforce "Emails Received" report exports off
// the main thread and returns one compact row per email. Nothing leaves the browser.
// Accepts the formatted export (.xlsx, grouped by Email Message Date and Company Name, with Subtotal rows)
// and the details-only export (.xls, really an HTML table). Report columns are found by name.
'use strict';
const DAY = 864e5;

self.onmessage = async (e) => {
  try { self.postMessage({ type: 'done', data: await build(e.data.files) }); }
  catch (err) { self.postMessage({ type: 'error', message: err && err.message ? err.message : String(err) }); }
};
const progress = (p) => self.postMessage({ type: 'progress', ...p });

// ---------- reading a file into rows of cells ----------
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
const decodeEnt = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, k) => {
  if (k[0] === '#') return String.fromCodePoint(k[1] === 'x' || k[1] === 'X' ? parseInt(k.slice(2), 16) : +k.slice(1));
  return ENT[k.toLowerCase()] ?? m;
});
function htmlRows(text) {
  const rows = [], tr = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi, cell = /<t[hd]\b[^>]*>([\s\S]*?)<\/t[hd]>/gi;
  let m;
  while ((m = tr.exec(text))) {
    const out = []; let c;
    cell.lastIndex = 0;
    while ((c = cell.exec(m[1]))) out.push(decodeEnt(c[1].replace(/<[^>]*>/g, '')).replace(/\s+/g, ' ').trim());
    rows.push(out);
  }
  return rows;
}
let sheetjs = false;
async function readRows(file) {
  const buf = await file.arrayBuffer();
  const head = new TextDecoder('latin1').decode(new Uint8Array(buf, 0, Math.min(2048, buf.byteLength)));
  if (/^\s*</.test(head.replace(/^﻿/, '')) && /<t(able|r)\b/i.test(head)) {
    const cs = (/charset=["']?([\w-]+)/i.exec(head) || [])[1] || 'windows-1252';
    let text;
    try { text = new TextDecoder(cs).decode(buf); } catch (err) { text = new TextDecoder('windows-1252').decode(buf); }
    return htmlRows(text);
  }
  if (!sheetjs) { importScripts('/assets/vendor/xlsx.full.min.js'); sheetjs = true; }
  const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: false });
    if (findHeader(rows) >= 0) return rows;
  }
  return [];
}
// "Email Message Date  ↑" -> "email message date"
const norm = (s) => String(s ?? '').replace(/[↑↓]/g, '').replace(/\s+/g, ' ').trim();
const key = (s) => norm(s).toLowerCase();
function findHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 80); i++) {
    const r = rows[i].map(key);
    if (r.includes('case number') && r.includes('email message date')) return i;
  }
  return -1;
}

// ---------- values ----------
// "9/21/2026, 9:13 AM", "2026-09-21 09:13" or an Excel serial -> days since 1970-01-01
function dayOf(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) && v > 20000 ? Math.floor(v) - 25569 : null;
  const s = String(v).trim();
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(s);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; return Date.UTC(y, m[1] - 1, +m[2]) / DAY; }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return Date.UTC(+m[1], m[2] - 1, +m[3]) / DAY;
  return /^\d+(\.\d+)?$/.test(s) ? dayOf(+s) : null;
}
const num = (v) => { const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, '')); return isFinite(n) ? n : null; };
const truthy = (v) => v === true || /^(1|true|yes)$/i.test(String(v).trim());
const domainOf = (s) => { const m = /@([^@\s>]+)\s*$/.exec(String(s || '')); return m ? m[1].toLowerCase() : ''; };
// "CSG-USCORPORATE-LEGACY" -> "US Corporate Legacy"
function teamLabel(s) {
  let t = norm(String(s).replace(/^CSG[-_ ]*/i, '').replace(/[_-]+/g, ' '));
  t = t.replace(/^us(?=[a-z])/i, 'US ');
  return t.split(' ').map((w) => (w.length <= 2 ? w.toUpperCase() : w[0].toUpperCase() + w.slice(1).toLowerCase())).join(' ');
}

// ---------- build ----------
async function build(items) {
  const files = [];
  // An email can't be identified exactly (the formatted export has no time per email), so overlapping
  // files are handled by count: for each case/day/sender/subject, the file with the most such emails wins.
  const merged = new Map();
  for (const { file, path } of items) {
    const name = file.name, parts = (path || '').split('/').filter(Boolean);
    const info = { name: path || name, team: null, rows: 0, from: null, to: null, asOf: '', note: '' };
    files.push(info);
    progress({ file: path || name, pct: 10, stage: 'Reading' });
    let rows;
    try { rows = await readRows(file); }
    catch (err) { info.note = /\.xls[xm]$/i.test(name) ? 'Could not be read (if it is open in Excel, close it and try again)' : 'Could not be read'; progress({ file: path || name, pct: 100, stage: info.note, skipped: true }); continue; }
    const hi = findHeader(rows);
    if (hi < 0) { info.note = 'Not an emails report (needs Case Number and Email Message Date)'; progress({ file: path || name, pct: 100, stage: info.note, skipped: true }); continue; }
    // Report title block: "Branch Code equals CSG-USCORPORATE-LEGACY", "As of 2026-10-05 14:17:34 …"
    for (const r of rows.slice(0, hi)) for (const c of r) {
      const s = norm(c);
      const b = /^Branch Code equals (.+)$/i.exec(s); if (b && !info.team) info.team = b[1].split(/\s*,\s*/).map(teamLabel).join(' + ');
      const a = /^As of (\d{4}-\d{2}-\d{2} \d{2}:\d{2})/i.exec(s); if (a) info.asOf = a[1];
    }
    if (!info.team && parts.length >= 2) info.team = teamLabel(parts[parts.length - 2]);
    if (!info.team) info.team = 'All';
    const H = new Map(); rows[hi].forEach((c, i) => { const k = key(c); if (k && !H.has(k)) H.set(k, i); });
    const col = (r, c) => (H.has(c) ? r[H.get(c)] : '');
    // Formatted exports print a group's value only on its first row and close each group with Subtotal.
    const grouped = rows.slice(hi + 1).some((r) => r.some((c) => /^(Subtotal|Total)$/.test(norm(c))));
    const fill = ['email message date', 'company name'].filter((c) => H.has(c));
    const last = {};
    const local = new Map();
    for (const r of rows.slice(hi + 1)) {
      if (grouped) {
        // a Subtotal in a group column ends that group (and every group inside it)
        fill.forEach((c, i) => { if (/^(Subtotal|Total)$/.test(norm(col(r, c)))) for (const x of fill.slice(i)) last[x] = ''; });
        for (const c of fill) { const v = norm(col(r, c)); if (v && !/^(Subtotal|Total)$/.test(v)) last[c] = v; }
      }
      const cn = norm(col(r, 'case number'));
      if (!cn) continue;
      const get = (c) => (grouped && fill.includes(c) ? last[c] || '' : col(r, c));
      const day = dayOf(get('email message date'));
      if (day == null) continue;
      const web = norm(col(r, 'web email')), status = norm(col(r, 'email status')) || '(none)';
      const rec = {
        day, company: norm(get('company name')) || '(no company)', from: norm(col(r, 'from name')) || '(no name)',
        subject: norm(col(r, 'email subject')) || norm(col(r, 'subject')), status, cn, domain: domainOf(web),
        opened: dayOf(col(r, 'date/time opened')), age: num(col(r, 'age')) ?? num(col(r, 'age (hours)')),
        open: H.has('open') ? truthy(col(r, 'open')) : H.has('closed') ? !truthy(col(r, 'closed')) : null, team: info.team,
      };
      const k = [cn, day, rec.from, rec.subject].join('\u0001');
      const o = local.get(k); if (o) o.list.push(rec); else local.set(k, { list: [rec] });
      info.rows++;
      if (info.from == null || day < info.from) info.from = day;
      if (info.to == null || day > info.to) info.to = day;
    }
    for (const [k, v] of local) { const o = merged.get(k); if (!o || v.list.length > o.list.length) merged.set(k, v); }
    progress({ file: path || name, pct: 100, rows: info.rows });
  }
  if (!files.some((f) => f.rows)) throw new Error('None of these files look like an emails report. Export the "Emails Received" report from Salesforce (it needs the Case Number and Email Message Date columns).');

  // compact rows with lookup lists for repeated text
  const dict = () => { const list = [], idx = new Map(); return { list, id: (s) => { if (!idx.has(s)) { idx.set(s, list.length); list.push(s); } return idx.get(s); } }; };
  const teams = dict(), comps = dict(), froms = dict(), subjects = dict(), statuses = dict(), domains = dict();
  const rows = [];
  for (const v of merged.values()) for (const r of v.list) {
    rows.push([teams.id(r.team), r.day, comps.id(r.company), froms.id(r.from), subjects.id(r.subject), statuses.id(r.status), r.cn,
      domains.id(r.domain), /(^|\.)cevalogistics\.com$/.test(r.domain) ? 1 : 0, r.opened, r.age, r.open == null ? -1 : r.open ? 1 : 0]);
  }
  return { teams: teams.list, companies: comps.list, froms: froms.list, subjects: subjects.list, statuses: statuses.list, domains: domains.list, rows,
    files: files.map((f) => ({ name: f.name, team: f.team, rows: f.rows, from: f.from, to: f.to, asOf: f.asOf, note: f.note })) };
}
