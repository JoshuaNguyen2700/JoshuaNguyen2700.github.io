// Background worker for the case-count tool: reads the Salesforce report exports off the main thread
// and boils them down to small count tables. Nothing leaves the browser.
//
// Report types are recognised by their columns, so the file names don't matter:
//   sent     Email Message Date + Case Owner, statuses Sent / Draft / Replied   -> cases actioned
//   received Email Message Date, statuses Read / New / Replied                 -> customer emails
//            (the "from CEVA" variant is a subset of the full one; overlapping emails count once)
//   closed   Date/Time Closed                                                  -> cases closed
// Salesforce ".xls" exports are HTML tables; real .xlsx/.xls/.csv files are read with SheetJS.
// The ACTIONED & CLOSED CASE COUNT workbook, if it is among the files, supplies the layout to mirror.
'use strict';
const DAY = 864e5;

self.onmessage = async (e) => {
  try {
    const data = await build(e.data.files);
    self.postMessage({ type: 'done', data });
  } catch (err) {
    self.postMessage({ type: 'error', message: err && err.message ? err.message : String(err) });
  }
};
const progress = (p) => self.postMessage({ type: 'progress', ...p });

// ---------- reading one file into rows of text ----------
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
function sheetRows(buf) {
  if (!sheetjs) { importScripts('/assets/vendor/xlsx.full.min.js'); sheetjs = true; }
  const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
  if (isTemplate(wb)) return { template: readTemplate(wb) };
  // Otherwise the first sheet that holds a report table wins (formatted exports put a title block above it).
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: false });
    if (findHeader(rows) >= 0) return rows;
  }
  return [];
}

// ---------- the workbook layout (sections, names, customer blocks) ----------
// Only the layout is taken from the workbook: section and row names, and which customer rows belong to
// which rep. Its numbers are kept only to recognise who is who (a nickname that isn't in Salesforce),
// on days the exports also cover. Ship counts (WK nn SHIP CT columns) are ignored.
const MONTH_SHEET = /^(JAN|FEB|MAR|APR|MAY|JUNE?|JULY?|AUG|SEPT?|OCT|NOV|DEC)[A-Z]*\.?\s+(\d{4})$/i;
const monthNo = (w) => ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'].indexOf(w.slice(0, 3).toUpperCase());
const isTemplate = (wb) => wb.SheetNames.some((n) => MONTH_SHEET.test(n.trim())) && wb.SheetNames.some((n) => /^CUST\s*E-?MAILS/i.test(n.trim()));
function grid(ws) {
  if (!ws || !ws['!ref']) return [];
  const R = XLSX.utils.decode_range(ws['!ref']), out = [];
  for (let r = 0; r <= R.e.r; r++) {
    const row = [];
    for (let c = 0; c <= R.e.c; c++) { const x = ws[XLSX.utils.encode_cell({ r, c })]; row.push(x ? { v: x.v, f: x.f } : null); }
    out.push(row);
  }
  return out;
}
const txt = (c) => (c && typeof c.v === 'string' ? c.v.replace(/\s+/g, ' ').trim() : '');
const isNum = (c) => c && typeof c.v === 'number';
const serialDay = (v) => Math.round(v) - 25569;
function readTemplate(wb) {
  // Month tabs ("SEPT 2026"): a section name (the supervisor), then CASES ACTIONED and CASES CLOSED
  // blocks with a date header row and one row per person until TOTALS. Empty month tabs are skipped.
  const months = [];
  for (const name of wb.SheetNames) {
    const m = MONTH_SHEET.exec(name.trim()); if (!m) continue;
    const g = grid(wb.Sheets[name]), sections = [], rowMap = {};
    let section = null;
    const sec = (n) => { let s = sections.find((x) => x.name === n); if (!s) sections.push((s = { name: n, labels: [], hist: {} })); return s; };
    for (let r = 0; r < g.length; r++) {
      const a = txt(g[r][0]);
      const kind = /^CASES ACTIONED/i.test(a) ? 'act' : /^CASES CLOSED/i.test(a) ? 'clo' : null;
      if (kind) {
        const cols = [];
        g[r].forEach((c, ci) => { if (ci && isNum(c) && c.v > 40000) cols.push([ci, serialDay(c.v)]); });
        const s = sec(section || 'TEAM');
        let k = r + 1;
        for (; k < g.length; k++) {
          const l = txt(g[k][0]);
          if (/^TOTALS?$/i.test(l) || /^CASES /i.test(l)) break;
          if (!l) { if (g[k].some(isNum)) break; continue; }
          if (!s.labels.includes(l)) s.labels.push(l);
          const h = s.hist[l] || (s.hist[l] = { act: {}, clo: {} });
          for (const [ci, d] of cols) if (isNum(g[k][ci])) h[kind][d] = g[k][ci].v;
          rowMap[k + 1] = { section: s.name, label: l };
        }
        r = k;
        continue;
      }
      if (a && !g[r].slice(1).some((c) => c && c.v !== '' && c.v != null) && !/^TOTAL/i.test(a)) section = a;
    }
    if (sections.length) months.push({ name: name.trim(), key: +m[2] * 12 + monthNo(m[1]), sections, rowMap });
  }
  months.sort((a, b) => a.key - b.key);
  const latest = months[months.length - 1];
  const byName = new Map(months.map((x) => [x.name.toUpperCase(), x]));
  const hist = {};   // section|name -> { act: {day: n}, clo: {day: n} } from the last three month tabs
  for (const mo of months.slice(-3)) for (const s of mo.sections) for (const [l, h] of Object.entries(s.hist)) {
    const o = hist[s.name + '|' + l] || (hist[s.name + '|' + l] = { act: {}, clo: {} });
    Object.assign(o.act, h.act); Object.assign(o.clo, h.clo);
  }

  // CUST E-MAILS (latest filled tab): a rep name, their customer rows, E-MAILS FROM CEVA STATIONS, TOTAL,
  // E-MAILS ACTIONED (formulas pointing at the rep's rows on the month tab) and % ACTIONED. A row with
  // dates beside its name starts a mailbox block.
  const custNames = wb.SheetNames.filter((n) => /^CUST\s*E-?MAILS/i.test(n.trim()) && (wb.Sheets[n]['!ref'] || 'A1') !== 'A1');
  const cust = [];
  if (custNames.length) {
    const g = grid(wb.Sheets[custNames[custNames.length - 1]]);
    let cur = null;
    // The rep is whichever month-tab row most of the E-MAILS ACTIONED formulas point at (a few cells
    // in the workbook point at the wrong column or row; the majority is right).
    const refOf = (row) => {
      const votes = new Map();
      for (const c of row) {
        if (!c || !c.f) continue;
        const re = /(?:'([^']+)'|([A-Za-z0-9_]+(?: [0-9]{4})?))!\$?[A-Z]{1,3}\$?(\d+)/g; let m;
        while ((m = re.exec(c.f))) {
          const mo = byName.get((m[1] || m[2]).trim().toUpperCase()), hit = mo && mo.rowMap[+m[3]];
          if (hit) { const k = hit.section + '|' + hit.label, v = votes.get(k); if (v) v.n++; else votes.set(k, { hit, n: 1 }); }
        }
      }
      let best = null; for (const v of votes.values()) if (!best || v.n > best.n) best = v;
      return best ? best.hit : null;
    };
    for (let r = 1; r < g.length; r++) {
      const row = g[r], a = txt(row[0]), rest = row.slice(1);
      if (/^CUSTOMER\s*\/\s*REP/i.test(a)) continue;
      if (a && rest.some((c) => isNum(c) && c.v > 40000)) { cust.push((cur = { label: a, mailbox: true, customers: [a], ref: null, ceva: false, total: false, open: true })); continue; }
      if (/^E-?MAILS FROM CEVA/i.test(a)) { if (cur) cur.ceva = true; continue; }
      if (/^TOTALS?$/i.test(a)) { if (cur) cur.total = true; continue; }
      if (/^E-?MAILS ACTIONED/i.test(a)) { if (cur) cur.ref = refOf(rest); continue; }
      if (/^%\s*ACTIONED/i.test(a)) { if (cur) cur.open = false; continue; }
      if (!a) continue;
      if (!cur || !cur.open) cust.push((cur = { label: a, mailbox: false, customers: [], ref: null, ceva: false, total: false, open: true }));
      else cur.customers.push(a);
    }
  }

  // Optional ROSTER tab: A = name in the workbook, B = Salesforce name, C = section (optional)
  const roster = [];
  const rosterName = wb.SheetNames.find((n) => /^roster$/i.test(n.trim()));
  if (rosterName) for (const row of grid(wb.Sheets[rosterName])) {
    const a = txt(row[0]), b = txt(row[1]);
    if (a && b && !/salesforce/i.test(b)) roster.push({ label: a, name: b, section: txt(row[2]) });
  }

  return { month: latest ? latest.name : null, sections: latest ? latest.sections.map((s) => ({ name: s.name, labels: s.labels })) : [],
    hist, cust: cust.map(({ open: _, ...b }) => b), roster };
}
async function readRows(file) {
  const buf = await file.arrayBuffer();
  const head = new TextDecoder('latin1').decode(new Uint8Array(buf, 0, Math.min(2048, buf.byteLength)));
  if (/^\s*</.test(head.replace(/^﻿/, '')) && /<t(able|r)\b/i.test(head)) {
    const cs = (/charset=["']?([\w-]+)/i.exec(head) || [])[1] || 'windows-1252';
    let text;
    try { text = new TextDecoder(cs).decode(buf); } catch (err) { text = new TextDecoder('windows-1252').decode(buf); }
    return htmlRows(text);
  }
  return sheetRows(buf);
}
const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
function findHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 80); i++) {
    const r = rows[i].map((x) => norm(x).toLowerCase());
    if (r.includes('case number') && (r.includes('case owner') || r.includes('email message date') || r.includes('date/time closed'))) return i;
  }
  return -1;
}

// ---------- values ----------
// "9/21/2026, 9:50 AM", "9/21/2026 9:50", "2026-09-21 09:50:00" or an Excel serial -> days since 1970-01-01
function dayOf(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) && v > 20000 ? Math.floor(v) - 25569 : null;
  const s = String(v).trim();
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(s);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; return Date.UTC(y, m[1] - 1, +m[2]) / DAY; }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return Date.UTC(+m[1], m[2] - 1, +m[3]) / DAY;
  if (/^\d+(\.\d+)?$/.test(s)) return dayOf(+s);
  return null;
}
// The same, to the minute (minutes since 1970-01-01), for the Case History report's Edit Date.
function minuteOf(v) {
  const d = dayOf(v); if (d == null) return null;
  if (typeof v === 'number') return Math.round(v * 1440) - 25569 * 1440;
  const m = /(\d{1,2}):(\d{2})(?::\d{2})?\s*([AP]M)?/i.exec(String(v).replace(/^\S+\s*/, ''));
  if (!m) return d * 1440;
  let h = +m[1] % (m[3] ? 12 : 24); if (m[3] && /p/i.test(m[3])) h += 12;
  return d * 1440 + h * 60 + +m[2];
}
const isCevaMail = (s) => /@([\w-]+\.)*cevalogistics\.com\s*$/i.test(String(s || ''));

// "UScorporate-cs" -> "US Corporate CS", "CSG-USEAST" -> "US East", "Legacy" -> "Legacy"
function teamLabel(s) {
  let t = norm(String(s).replace(/^CSG[-_ ]*/i, '').replace(/[_-]+/g, ' '));
  t = t.replace(/^us(?=[a-z])/i, 'US ');
  return t.split(' ').map((w) => (w.length <= 2 ? w.toUpperCase() : w[0].toUpperCase() + w.slice(1).toLowerCase())).join(' ');
}
// Salesforce sometimes stores names in lower case ("dweana collins"); show them in title case.
function niceName(s) {
  const n = norm(s);
  return n && n === n.toLowerCase() ? n.replace(/(^|[\s.'-])([a-z])/g, (m, a, b) => a + b.toUpperCase()) : n;
}

function kindOf(h, rows) {
  if (h.has('edited by') && h.has('edit date') && h.has('new value')) return 'history';
  if (h.has('date/time closed') && h.has('case owner')) return 'closed';
  if (h.has('email message date')) {
    const st = h.get('email status');
    if (st != null) {
      let out = 0, n = 0;
      for (const r of rows) { const v = norm(r[st]).toLowerCase(); if (!v) continue; n++; if (v === 'sent' || v === 'draft') out++; }
      if (n) return out / n > 0.3 ? 'sent' : 'received';
    }
    return h.has('web email') ? 'received' : 'sent';
  }
  return null;
}

// Files picked on their own (no team folder in the path) take their team from their contents: first from
// cases they share with files whose team is known, then from the sender on a sent-emails report
// ("CEVA Ground USEAST" -> East, "CEVA Ground USCORPORATE-LEGACY" -> Legacy), then shared cases again.
function teamsFromContents(files) {
  const loose = () => files.filter((f) => !f.team && f.kind && f.kind !== 'workbook' && f.kind !== 'history');   // case history covers every team
  const byCases = () => {
    const caseTeam = new Map();
    for (const f of files) if (f.team && f.cases) for (const cn of f.cases) if (!caseTeam.has(cn)) caseTeam.set(cn, f.team);
    for (const f of loose()) {
      const votes = new Map();
      for (const cn of f.cases) { const t = caseTeam.get(cn); if (t) votes.set(t, (votes.get(t) || 0) + 1); }
      const best = [...votes].sort((a, b) => b[1] - a[1])[0];
      if (best && best[1] >= Math.max(3, f.cases.size * 0.05)) f.team = best[0];
    }
  };
  byCases();
  for (const f of loose()) {
    const top = [...f.senders].sort((a, b) => b[1] - a[1])[0];
    if (!top) continue;
    const s = top[0].replace(/^CEVA\s+Ground\s*/i, '').split(/[-_]/).pop().replace(/^US(?=[A-Z]{3,})/i, '');
    if (s) f.team = teamLabel(s);
  }
  byCases();
}

// ---------- build ----------
async function build(items) {
  const files = [];
  let template = null;
  // Phase 1: read every file and keep only the fields we need, keyed so overlapping exports count once.
  // Within one file a key can repeat (two emails in the same minute); across files the larger count wins,
  // and a case owner known from either copy is kept.
  const K = { sent: new Map(), received: new Map(), closed: new Map(), history: new Map() };
  const merge = (kind, local) => {
    const g = K[kind];
    for (const [k, v] of local) {
      const o = g.get(k);
      if (!o) { g.set(k, v); continue; }
      if (v.n > o.n) { if (!v.owner) v.owner = o.owner; g.set(k, v); } else if (!o.owner && v.owner) o.owner = v.owner;
    }
  };
  // A file at the top of the chosen folder, next to team folders (the workbook, a Case History export), isn't a team.
  const depth = Math.max(...items.map((it) => (it.path || '').split('/').filter(Boolean).length));
  for (const { file, path } of items) {
    const name = file.name, parts = (path || '').split('/').filter(Boolean);
    const folder = parts.length >= 2 && (parts.length > 2 || depth <= 2) ? parts[parts.length - 2] : null;
    const info = { name, path: path || name, team: folder ? teamLabel(folder) : null, kind: null, rows: 0, from: null, to: null, note: '' };
    files.push(info);
    progress({ file: path || name, pct: 10, stage: 'Reading' });
    let rows;
    try { rows = await readRows(file); }
    catch (err) {
      info.note = /\.xls[xm]$/i.test(name) ? 'Could not be read (if it is open in Excel, close it and load again)' : 'Could not be read';
      progress({ file: path || name, pct: 100, stage: info.note, skipped: true }); continue;
    }
    if (rows && rows.template) {
      info.kind = 'workbook'; info.team = null; template = rows.template;
      info.note = template.month ? `Layout from ${template.month}` : 'No filled month tab found';
      progress({ file: path || name, pct: 100, stage: 'Workbook layout' });
      continue;
    }
    const hi = findHeader(rows);
    if (hi < 0) { info.note = 'Not a Salesforce case or email report'; progress({ file: path || name, pct: 100, stage: info.note, skipped: true }); continue; }
    const head = rows[hi].map((x) => norm(x).toLowerCase()), H = new Map();
    head.forEach((c, i) => { if (!H.has(c)) H.set(c, i); });
    const body = rows.slice(hi + 1).filter((r) => r.some((x) => x !== '' && x != null) && norm(r[H.get('case number')]));
    const kind = kindOf(H, body);
    if (!kind) { info.note = 'Not a recognised report type'; progress({ file: path || name, pct: 100, stage: info.note, skipped: true }); continue; }
    info.kind = kind;
    if (kind === 'history') info.team = null;   // one export usually covers every team
    const col = (r, c) => (H.has(c) ? r[H.get(c)] : '');
    const local = new Map(), add = (key, rec) => { const o = local.get(key); if (o) o.n++; else local.set(key, { ...rec, n: 1 }); };
    const span = (d) => { if (d == null) return; if (info.from == null || d < info.from) info.from = d; if (info.to == null || d > info.to) info.to = d; };
    info.cases = new Set(); info.senders = new Map();   // for finding the team of a file picked on its own
    for (const r of body) {
      const cn = norm(col(r, 'case number')), owner = norm(col(r, 'case owner'));
      info.cases.add(cn);
      if (kind === 'sent') { const s = norm(col(r, 'from name')); if (s) info.senders.set(s, (info.senders.get(s) || 0) + 1); }
      if (kind === 'sent' || kind === 'received') {
        const day = dayOf(col(r, 'email message date')); if (day == null) continue;
        span(day);
        // (status is left out of the key: a "Read" email becomes "Replied" in a later export)
        const key = [cn, norm(col(r, 'email message date')), norm(col(r, 'from name')), norm(col(r, 'email subject'))].join('\u0001');
        if (kind === 'sent') add(key, { owner, cn, day, folder: info.team, fi: files.length - 1 });
        else add(key, { owner, cn, day, folder: info.team, fi: files.length - 1, company: norm(col(r, 'company name')), ceva: isCevaMail(col(r, 'web email')) });
      } else if (kind === 'history') {
        // a status change into a Closed status from an open one (Closed-Resolved -> Closed-No Action Needed is not a new close)
        if (!/^closed/i.test(norm(col(r, 'new value'))) || /^closed/i.test(norm(col(r, 'old value')))) continue;
        const at = minuteOf(col(r, 'edit date')); if (at == null) continue;
        span(Math.floor(at / 1440));
        add(cn + '\u0001' + at, { owner, cn, at, day: Math.floor(at / 1440), folder: null, fi: files.length - 1 });
      } else {
        const day = dayOf(col(r, 'date/time closed')); if (day == null) continue;
        span(day);
        add(cn + '\u0001' + norm(col(r, 'date/time closed')), { owner, cn, day, folder: info.team, fi: files.length - 1 });
      }
    }
    info.rows = body.length;
    merge(kind, local);
    progress({ file: path || name, pct: 100, rows: info.rows, kind });
  }
  if (!files.some((f) => f.kind && f.kind !== 'workbook')) throw new Error('None of these files look like Salesforce case or email reports. Choose the Cview Report folder, or the .xls exports in its team folders.');
  // With a Case History export, CASES CLOSED is counted the way the workbook was filled in, from a report
  // pulled the next day: a case counts on the day it was closed, for its owner, unless it was closed again
  // within 24 hours (it had reopened), in which case only that later close counts. The Closed Cases per
  // Agent report is the fallback; it only lists cases that are still closed now, so it runs low.
  if (K.history.size) {
    const byCase = new Map();
    for (const v of K.history.values()) { const a = byCase.get(v.cn); if (a) a.push(v); else byCase.set(v.cn, [v]); }
    const kept = new Map();
    for (const evs of byCase.values()) {
      evs.sort((a, b) => a.at - b.at);
      evs.forEach((v, i) => {
        const next = evs[i + 1];
        if (next && next.day === v.day) return;          // closed again later the same day
        if (next && next.at - v.at < 24 * 60) return;     // reopened and closed again within 24 hours
        kept.set(v.cn + '\u0001' + v.at, { ...v, n: 1 });
      });
    }
    K.closed = kept;
  }
  teamsFromContents(files);
  for (const kind of ['sent', 'closed', 'received']) for (const v of K[kind].values()) if (!v.folder && v.fi != null) v.folder = files[v.fi].team;
  // Closures from Case History take the team of the case in the other reports.
  if (K.history.size) {
    const caseTeam = new Map();
    for (const f of files) if (f.team && f.cases && f.kind !== 'history') for (const cn of f.cases) if (!caseTeam.has(cn)) caseTeam.set(cn, f.team);
    // A history export run for all cases also holds other countries' teams: keep a closure only when its case is in
    // the team reports or its owner is someone in them.
    const known = new Set();
    for (const kind of ['sent', 'received']) for (const v of K[kind].values()) if (v.owner) known.add(v.owner.toLowerCase());
    for (const [k, v] of K.closed) {
      if (!v.folder) v.folder = caseTeam.get(v.cn) || null;
      if (!v.folder && !known.has((v.owner || '').toLowerCase())) K.closed.delete(k);
    }
  }

  // Phase 2: who owns each case, and which team folder each person belongs to.
  const caseOwner = new Map();
  for (const kind of ['closed', 'sent', 'received']) for (const v of K[kind].values()) if (v.cn && v.owner && !caseOwner.has(v.cn)) caseOwner.set(v.cn, v.owner);
  const pKey = (s) => norm(s).toLowerCase();
  const people = new Map();   // key -> { name, votes: Map(team -> n) }
  const person = (raw, team, n) => {
    const k = pKey(raw); if (!k) return null;
    let p = people.get(k);
    if (!p) people.set(k, (p = { name: niceName(raw), votes: new Map() }));
    else if (p.name === p.name.toLowerCase() && raw !== raw.toLowerCase()) p.name = niceName(raw);
    if (team) p.votes.set(team, (p.votes.get(team) || 0) + n);
    return k;
  };
  const ownerOf = (v) => v.owner || (v.cn ? caseOwner.get(v.cn) || '' : '');
  for (const kind of ['sent', 'closed', 'received']) for (const v of K[kind].values()) v.pk = person(ownerOf(v), v.folder, v.n);
  const home = new Map();
  for (const [k, p] of people) {
    let best = null, bn = -1;
    for (const [t, n] of p.votes) if (n > bn || (n === bn && t < best)) { best = t; bn = n; }
    home.set(k, best || 'Other');
  }

  // Phase 3: count tables. A person's activity always counts for their home team; emails on cases
  // with no known owner stay with the team folder they came from.
  const teams = [], tIdx = new Map(), T = (t) => { t = t || 'Other'; if (!tIdx.has(t)) { tIdx.set(t, teams.length); teams.push(t); } return tIdx.get(t); };
  const names = [], pIdx = new Map(), P = (k) => { if (!k) return -1; if (!pIdx.has(k)) { pIdx.set(k, names.length); names.push(k); } return pIdx.get(k); };
  const comps = [], cIdx = new Map(), C = (c) => { c = c || '(no company)'; if (!cIdx.has(c)) { cIdx.set(c, comps.length); comps.push(c); } return cIdx.get(c); };
  // A file picked on its own (no team folder in its path) takes the team most of its case owners belong to.
  const fileVotes = files.map(() => new Map());
  for (const kind of ['sent', 'closed', 'received']) for (const v of K[kind].values()) if (v.pk && v.fi != null) { const m = fileVotes[v.fi], h = home.get(v.pk); m.set(h, (m.get(h) || 0) + v.n); }
  files.forEach((f, i) => { if (!f.team && f.kind && f.kind !== 'workbook' && f.kind !== 'history') { const best = [...fileVotes[i]].sort((a, b) => b[1] - a[1])[0]; if (best) f.team = best[0]; } });
  const teamFor = (v) => T(v.pk ? home.get(v.pk) : v.folder || (v.fi != null && files[v.fi].team) || null);
  const tally = (map, key, n) => map.set(key, (map.get(key) || 0) + n);
  const act = new Map(), clo = new Map(), inb = new Map();
  for (const v of K.sent.values()) tally(act, [teamFor(v), P(v.pk), v.day].join(), v.n);
  for (const v of K.closed.values()) tally(clo, [teamFor(v), P(v.pk), v.day].join(), v.n);
  for (const v of K.received.values()) tally(inb, [teamFor(v), P(v.pk), C(v.company), v.day, v.ceva ? 1 : 0].join(), v.n);
  const flat = (m) => [...m].map(([k, n]) => [...k.split(',').map(Number), n]);
  return {
    teams, people: names.map((k) => people.get(k).name), companies: comps,
    act: flat(act), clo: flat(clo), inb: flat(inb),
    files: files.map((f) => ({ name: f.path, team: f.team, kind: f.kind, rows: f.rows, from: f.from, to: f.to, note: f.note })),
    template,
  };
}
