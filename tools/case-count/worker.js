// Background worker for the case-count tool: reads the Salesforce report exports off the main thread
// and boils them down to small count tables. Nothing leaves the browser.
//
// Report types are recognised by their columns, so the file names don't matter:
//   sent     Email Message Date + Case Owner, statuses Sent / Draft / Replied   -> emails actioned
//   received Email Message Date, statuses Read / New / Replied                 -> emails received
//            (the "from CEVA" variant is a subset of the full one; overlapping emails count once)
//   closed   Date/Time Closed                                                  -> cases closed
//   sla      Elapsed Time (Mins) / SLA Breached?                               -> first response, open cases
// Salesforce ".xls" exports are HTML tables; real .xlsx/.xls/.csv files are read with SheetJS.
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
  // The ACTIONED & CLOSED CASE COUNT workbook itself supplies the layout to mirror.
  if (isTemplate(wb)) return { template: readTemplate(wb) };
  // Otherwise the first sheet that holds a report table wins (formatted exports put a title block above it).
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: false });
    if (findHeader(rows) >= 0) return rows;
  }
  return [];
}

// ---------- the workbook layout (sections, names, customer blocks, SLA lists) ----------
// Only the layout is taken from the workbook: section and row names, which customer rows belong to
// which rep, the WEEKLY SLA lists and the Open cases names. Its numbers are kept only to recognise
// who is who (e.g. a nickname that isn't in Salesforce), on days the exports also cover.
const MONTH_SHEET = /^(JAN|FEB|MAR|APR|MAY|JUNE?|JULY?|AUG|SEPT?|OCT|NOV|DEC)[A-Z]*\.?\s+(\d{4})$/i;
const monthNo = (w) => ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'].indexOf(w.slice(0, 3).toUpperCase());
const isTemplate = (wb) => wb.SheetNames.some((n) => MONTH_SHEET.test(n.trim())) && wb.SheetNames.some((n) => /^(WEEKLY\s*SLA|CUST\s*E-?MAILS)/i.test(n.trim()));
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
  // blocks with a date header row and one row per person until TOTALS.
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

  // CUST E-MAILS (latest tab): a rep name, their customer rows, E-MAILS FROM CEVA STATIONS, TOTAL,
  // E-MAILS ACTIONED (a formula pointing at the rep's rows on the month tab) and % ACTIONED. A row
  // with dates beside its name starts a mailbox block.
  const custNames = wb.SheetNames.filter((n) => /^CUST\s*E-?MAILS/i.test(n.trim()));
  const cust = [];
  if (custNames.length) {
    const g = grid(wb.Sheets[custNames[custNames.length - 1]]);
    let cur = null;
    const refOf = (row) => {
      for (const c of row) {
        if (!c || !c.f) continue;
        const re = /(?:'([^']+)'|([A-Za-z0-9_]+(?: [0-9]{4})?))!\$?[A-Z]{1,3}\$?(\d+)/g; let m;
        while ((m = re.exec(c.f))) {
          const mo = byName.get((m[1] || m[2]).trim().toUpperCase()), hit = mo && mo.rowMap[+m[3]];
          if (hit) return hit;
        }
      }
      return null;
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

  // WEEKLY SLA: the name lists under the latest RUN DATE, with the title at the top of each column
  const sla = [];
  const slaName = wb.SheetNames.find((n) => /^WEEKLY\s*SLA/i.test(n.trim()));
  if (slaName) {
    const g = grid(wb.Sheets[slaName]), runRows = [];
    g.forEach((row, r) => { if (row.some((c) => /^RUN DATE/i.test(txt(c)))) runRows.push(r); });
    if (runRows.length) {
      const first = runRows[0], last = runRows[runRows.length - 1];
      g[last].forEach((c, ci) => {
        if (!/^RUN DATE/i.test(txt(c))) return;
        let title = '';
        for (let r = 0; r < first && !title; r++) title = txt(g[r][ci]);
        const labels = [];
        for (let r = last + 1; r < g.length; r++) {
          const l = txt(g[r][ci]);
          if (/^CSR$/i.test(l)) continue;
          if (!l || /^TEAM AVERAGE/i.test(l)) break;
          labels.push(l);
        }
        sla.push({ title, labels });
      });
    }
  }

  // Open cases: names down column A, statuses across row 1
  let open = null;
  const openName = wb.SheetNames.find((n) => /^open cases$/i.test(n.trim()));
  if (openName) {
    const g = grid(wb.Sheets[openName]), statuses = [], names = [];
    for (let c = 1; c < (g[0] || []).length; c++) { const h = txt(g[0][c]); if (!h || /^total/i.test(h)) break; statuses.push(h); }
    for (let r = 1; r < g.length; r++) { const l = txt(g[r][0]); if (!l) continue; if (/^totals?/i.test(l)) break; names.push(l); }
    open = { statuses, names };
  }

  // Optional ROSTER tab: A = name in the workbook, B = Salesforce name, C = section (optional)
  const roster = [];
  const rosterName = wb.SheetNames.find((n) => /^roster$/i.test(n.trim()));
  if (rosterName) for (const row of grid(wb.Sheets[rosterName])) {
    const a = txt(row[0]), b = txt(row[1]);
    if (a && b && !/salesforce/i.test(b)) roster.push({ label: a, name: b, section: txt(row[2]) });
  }

  return { month: latest ? latest.name : null, sections: latest ? latest.sections.map((s) => ({ name: s.name, labels: s.labels })) : [],
    hist, cust: cust.map(({ open: _, ...b }) => b), sla, open, roster };
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
const num = (v) => { const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, '')); return isFinite(n) ? n : null; };
const monday = (d) => d - ((d + 3) % 7);   // day 0 (1970-01-01) was a Thursday
const isCevaMail = (s) => /@([\w-]+\.)*cevalogistics\.com\s*$/i.test(String(s || ''));

// "UScorporate-cs" -> "US Corporate CS", "CSG-USEAST" -> "US East"
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
  if (h.has('elapsed time (mins)') || h.has('sla breached?')) return 'sla';
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

// ---------- build ----------
async function build(items) {
  const files = [];
  let template = null;
  // Phase 1: read every file and keep only the fields we need, keyed so overlapping exports count once.
  // Within one file a key can repeat (two emails in the same minute); across files the larger count wins.
  const K = { sent: new Map(), received: new Map(), closed: new Map(), sla: new Map() };
  const merge = (kind, local) => {
    const g = K[kind];
    for (const [k, v] of local) { const o = g.get(k); if (!o || v.n > o.n) g.set(k, v); }
  };
  for (const { file, path } of items) {
    const name = file.name, parts = (path || '').split('/').filter(Boolean);
    const folder = parts.length >= 2 ? parts[parts.length - 2] : null;
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
      info.kind = 'workbook'; template = rows.template;
      info.note = template.month ? `Layout from ${template.month}` : 'No month tab found';
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
    const col = (r, c) => (H.has(c) ? r[H.get(c)] : '');
    const local = new Map(), add = (key, rec) => { const o = local.get(key); if (o) o.n++; else local.set(key, { ...rec, n: 1 }); };
    const span = (d) => { if (d == null) return; if (info.from == null || d < info.from) info.from = d; if (info.to == null || d > info.to) info.to = d; };
    for (const r of body) {
      const cn = norm(col(r, 'case number')), owner = norm(col(r, 'case owner'));
      if (kind === 'sent' || kind === 'received') {
        const day = dayOf(col(r, 'email message date')); if (day == null) continue;
        span(day);
        // (status is left out of the key: a "Read" email becomes "Replied" in a later export)
        const key = [cn, norm(col(r, 'email message date')), norm(col(r, 'from name')), norm(col(r, 'email subject'))].join('\u0001');
        if (kind === 'sent') add(key, { owner, day, folder: info.team, fi: files.length - 1 });
        else add(key, { owner, cn, day, folder: info.team, fi: files.length - 1, company: norm(col(r, 'company name')), ceva: isCevaMail(col(r, 'web email')) });
      } else if (kind === 'closed') {
        const day = dayOf(col(r, 'date/time closed')); if (day == null) continue;
        span(day);
        add(cn + '\u0001' + norm(col(r, 'date/time closed')), { owner, cn, day, folder: info.team, fi: files.length - 1 });
      } else {
        const opened = dayOf(col(r, 'date/time opened'));
        span(opened);
        const br = norm(col(r, 'sla breached?')).toLowerCase(), vio = norm(col(r, 'violation'));
        const met = /met/.test(br) ? 1 : /breach/.test(br) ? 0 : vio === '0' ? 1 : vio === '1' ? 0 : null;
        local.set(cn, { owner, cn, opened, folder: info.team, fi: files.length - 1, branch: norm(col(r, 'branch code')),
          elapsed: num(col(r, 'elapsed time (mins)')), met, status: norm(col(r, 'status')), n: 1 });
      }
    }
    info.rows = body.length;
    merge(kind, local);
    progress({ file: path || name, pct: 100, rows: info.rows, kind });
  }
  if (!files.some((f) => f.kind && f.kind !== 'workbook')) throw new Error('None of these files look like Salesforce case or email reports. Choose the .xls exports from the Cview Report folders.');

  // Phase 2: who owns each case, and which team each person belongs to.
  const caseOwner = new Map();
  for (const kind of ['sla', 'closed', 'received']) for (const v of K[kind].values()) if (v.cn && v.owner && !caseOwner.has(v.cn)) caseOwner.set(v.cn, v.owner);
  const branchTeam = new Map();   // branch code -> folder name, learned from SLA files that sit in a team folder
  for (const v of K.sla.values()) if (v.branch && v.folder && !branchTeam.has(v.branch)) branchTeam.set(v.branch, v.folder);
  const teamOfRec = (v) => v.folder || (v.branch ? branchTeam.get(v.branch) || teamLabel(v.branch) : null);

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
  for (const kind of ['sent', 'closed', 'received', 'sla']) for (const v of K[kind].values()) { v.pk = person(ownerOf(v), teamOfRec(v), v.n); v.team = teamOfRec(v); }
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
  const stats = [], sIdx = new Map(), ST = (s) => { if (!sIdx.has(s)) { sIdx.set(s, stats.length); stats.push(s); } return sIdx.get(s); };
  // A file picked on its own (no team folder in its path) takes the team most of its case owners belong to,
  // so emails on cases with no known owner still land in the right team.
  const fileVotes = files.map(() => new Map());
  for (const kind of ['sent', 'closed', 'received', 'sla']) for (const v of K[kind].values()) if (v.pk && v.fi != null) { const m = fileVotes[v.fi], h = home.get(v.pk); m.set(h, (m.get(h) || 0) + v.n); }
  files.forEach((f, i) => { if (!f.team && f.kind && f.kind !== 'workbook') { const best = [...fileVotes[i]].sort((a, b) => b[1] - a[1])[0]; if (best) f.team = best[0]; } });
  const teamFor = (v) => T(v.pk ? home.get(v.pk) : v.team || (v.fi != null && files[v.fi].team) || null);
  const tally = (map, key, n) => map.set(key, (map.get(key) || 0) + n);
  const act = new Map(), clo = new Map(), inb = new Map(), sla = new Map(), open = new Map();
  for (const v of K.sent.values()) tally(act, [teamFor(v), P(v.pk), v.day].join(), v.n);
  for (const v of K.closed.values()) tally(clo, [teamFor(v), P(v.pk), v.day].join(), v.n);
  for (const v of K.received.values()) tally(inb, [teamFor(v), P(v.pk), C(v.company), v.day, v.ceva ? 1 : 0].join(), v.n);
  for (const v of K.sla.values()) {
    const t = teamFor(v), p = P(v.pk);
    if (v.opened != null && v.elapsed != null) {
      const k = [t, p, monday(v.opened)].join(), o = sla.get(k) || [0, 0, 0, 0];
      o[0]++; o[1] += v.elapsed; if (v.met != null) { o[2]++; o[3] += v.met; }
      sla.set(k, o);
    }
    if (v.status && !/^closed/i.test(v.status)) tally(open, [t, p, ST(v.status)].join(), 1);
  }
  const flat = (m) => [...m].map(([k, n]) => [...k.split(',').map(Number), ...(Array.isArray(n) ? n : [n])]);
  return {
    teams, people: names.map((k) => people.get(k).name), companies: comps, statuses: stats,
    act: flat(act), clo: flat(clo), inb: flat(inb), sla: flat(sla), open: flat(open),
    files: files.map((f) => ({ name: f.path, team: f.team, kind: f.kind, rows: f.rows, from: f.from, to: f.to, note: f.note })),
    template,
  };
}
