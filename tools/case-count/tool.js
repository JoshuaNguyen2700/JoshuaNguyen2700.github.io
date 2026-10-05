// Logic for the case-count tool only. It is an ES module, so nothing here leaks into other pages.
// The page ships with no case data. The visitor picks the Salesforce report exports (or the folder that
// holds them); worker.js reads them in the browser and returns small count tables (nothing is uploaded).
// The page mirrors the "ACTIONED & CLOSED CASE COUNT" workbook tab for tab: a month tab per month
// (CASES ACTIONED / CASES CLOSED by section), CUST E-MAILS, WEEKLY SLA and Open cases. When that
// workbook is loaded with the reports, its layout (sections, names, customer rows) is used; the names
// are matched to Salesforce owners here, and the Names tab shows and corrects each match.
import { esc } from '/assets/core/util.js';

const $ = (s) => document.querySelector(s);
const STATE_KEY = 'cc.state';   // the open tab only; never data
const GOAL = 0.85;

// ---------- dates (whole days since 1970-01-01, as the worker sends them) ----------
const DAY = 864e5;
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const SHEET_MON = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE', 'JULY', 'AUG', 'SEPT', 'OCT', 'NOV', 'DEC'];   // as the workbook names its tabs
const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const dt = (d) => new Date(d * DAY);
const dowOf = (d) => dt(d).getUTCDay();
const monday = (d) => d - ((d + 3) % 7);
const md = (d) => `${dt(d).getUTCMonth() + 1}/${dt(d).getUTCDate()}`;
const dLong = (d, year = true) => `${MON[dt(d).getUTCMonth()]} ${dt(d).getUTCDate()}${year ? ', ' + dt(d).getUTCFullYear() : ''}`;
const span = (a, b) => (a === b ? dLong(a) : dt(a).getUTCFullYear() === dt(b).getUTCFullYear() ? `${dLong(a, false)} – ${dLong(b)}` : `${dLong(a)} – ${dLong(b)}`);
function isoWeek(d) { const th = monday(d) + 3, y = dt(th).getUTCFullYear(); return 1 + Math.floor((th - Date.UTC(y, 0, 1) / DAY) / 7); }
const toInput = (d) => dt(d).toISOString().slice(0, 10);
const fromInput = (s) => (/^\d{4}-\d{2}-\d{2}$/.test(s) ? Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) / DAY : null);
const monthKey = (d) => dt(d).getUTCFullYear() * 12 + dt(d).getUTCMonth();
const monthStart = (k) => Date.UTC(Math.floor(k / 12), k % 12, 1) / DAY;
const monthEnd = (k) => monthStart(k + 1) - 1;
const sheetName = (k) => `${SHEET_MON[k % 12]} ${Math.floor(k / 12)}`;

// ---------- formatting ----------
const fmtN = (n) => Math.round(n).toLocaleString('en-US');
const fmtP = (p, d = 0) => (p == null || !isFinite(p) ? '–' : (p * 100).toFixed(d) + '%');
const fmtHM = (mins) => { if (mins == null || !isFinite(mins)) return '–'; const t = Math.round(mins); return Math.floor(t / 60) + ':' + String(t % 60).padStart(2, '0'); };
const plural = (n, w) => `${fmtN(n)} ${w}${n === 1 ? '' : 's'}`;
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const byText = (a, b) => String(a).localeCompare(String(b), 'en', { sensitivity: 'base', numeric: true });

// ---------- data + state ----------
// D (from worker.js): teams[], people[], companies[], statuses[], files[], template, and count tables
//   act  [team, person, day, n]            emails actioned (sent)        person -1 = no known case owner
//   clo  [team, person, day, n]            cases closed
//   inb  [team, person, company, day, ceva, n]   emails received (ceva 1 = from a CEVA address)
//   sla  [team, person, monday, cases, elapsedMins, slaKnown, slaMet]
//   open [team, person, status, n]
let D = null, S = null, V = null, X = null, I = null, L = null, worker = null;
const OVR = new Map();   // name fixes made on the Names tab: "section|name" -> person (this page only)
const saved = (() => { try { return JSON.parse(localStorage.getItem(STATE_KEY)) || {}; } catch (e) { return {}; } })();
const saveView = () => { try { localStorage.setItem(STATE_KEY, JSON.stringify({ tab: S.tab })); } catch (e) {} };
const pName = (p) => (p < 0 ? '(no case owner)' : D.people[p]);
const teamOrder = (a, b) => byText(D.teams[a], D.teams[b]);
const personOrder = (a, b) => (a.p < 0) - (b.p < 0) || byText(pName(a.p), pName(b.p));

function extent() {
  let lo = Infinity, hi = -Infinity, slaLo = Infinity, slaHi = -Infinity;
  for (const r of D.act) { lo = Math.min(lo, r[2]); hi = Math.max(hi, r[2]); }
  for (const r of D.clo) { lo = Math.min(lo, r[2]); hi = Math.max(hi, r[2]); }
  for (const r of D.inb) { lo = Math.min(lo, r[3]); hi = Math.max(hi, r[3]); }
  for (const r of D.sla) { slaLo = Math.min(slaLo, r[2]); slaHi = Math.max(slaHi, r[2] + 6); }
  return { lo, hi, slaLo, slaHi, from: Math.min(lo, slaLo), to: isFinite(hi) ? hi : slaHi };
}
// Unfiltered lookups used by the workbook tabs.
function index() {
  const pday = new Map(), comp = new Map(), compOwn = new Map(), slaw = new Map(), open = new Map();
  const slot = (p, d) => { let m = pday.get(p); if (!m) pday.set(p, (m = new Map())); let v = m.get(d); if (!v) m.set(d, (v = [0, 0, 0, 0])); return v; };
  for (const r of D.act) if (r[1] >= 0) slot(r[1], r[2])[0] += r[3];
  for (const r of D.clo) if (r[1] >= 0) slot(r[1], r[2])[1] += r[3];
  for (const r of D.inb) {
    if (r[1] >= 0) { const v = slot(r[1], r[3]); v[3] += r[5]; if (r[4]) v[2] += r[5]; }
    let m = comp.get(r[2]); if (!m) comp.set(r[2], (m = new Map())); m.set(r[3], (m.get(r[3]) || 0) + r[5]);
    if (r[1] >= 0 && !r[4]) { const k = r[1] + ',' + r[2]; let o = compOwn.get(k); if (!o) compOwn.set(k, (o = new Map())); o.set(r[3], (o.get(r[3]) || 0) + r[5]); }
  }
  for (const r of D.sla) { if (r[1] < 0) continue; let m = slaw.get(r[1]); if (!m) slaw.set(r[1], (m = new Map())); const v = m.get(r[2]) || [0, 0, 0, 0]; for (let i = 0; i < 4; i++) v[i] += r[3 + i]; m.set(r[2], v); }
  for (const r of D.open) { if (r[1] < 0) continue; let m = open.get(r[1]); if (!m) open.set(r[1], (m = new Map())); m.set(r[2], (m.get(r[2]) || 0) + r[3]); }
  const weeks = [...new Set(D.sla.map((r) => r[2]))].sort((a, b) => a - b);
  // months the exports cover, oldest first
  const months = [];
  if (isFinite(X.lo)) for (let k = monthKey(X.lo); k <= monthKey(X.hi); k++) months.push(k);
  return { pday, comp, compOwn, slaw, open, weeks, months };
}
const dflt = () => ({ teams: D.teams.map((_, i) => i), person: -1, from: X.from, to: X.to });
function load(data) {
  D = data; X = extent(); I = index(); OVR.clear();
  S = { ...dflt(), tab: saved.tab };
  $('#notice').hidden = true;
  const usable = D.files.filter((f) => f.kind && f.kind !== 'workbook'), skipped = D.files.filter((f) => !f.kind);
  $('#meta').textContent = `${plural(usable.length, 'report')} · ${D.people.length === 1 ? '1 person' : fmtN(D.people.length) + ' people'}`;
  const KN = { sent: 'sent emails (actioned)', received: 'received emails', closed: 'closed cases', sla: 'SLA' };
  const miss = Object.keys(KN).filter((k) => !usable.some((f) => f.kind === k)), notes = [];
  if (!D.template) notes.push('The ACTIONED & CLOSED CASE COUNT workbook wasn\'t in the files, so the tabs are laid out by team folder with full names. Load the folder that holds it (closed in Excel) to get your sections, names and customer rows.');
  if (miss.length) notes.push(`No ${miss.map((k) => KN[k]).join(', ')} report was found, so those figures are blank.`);
  if (skipped.length) notes.push(`Skipped: ${skipped.map((f) => `${f.name.split('/').pop()} (${f.note})`).join(', ')}.`);
  if (notes.length) { $('#notice').hidden = false; $('#notice').textContent = notes.join(' '); }
  buildLayout(); buildFilters(); buildTabs(); show('app'); setTab(S.tab);
}

// ---------- screens + loading ----------
function show(view) { for (const v of ['start', 'progress', 'app']) $('#' + v).hidden = v !== view; if (view !== 'app') hideTip(); }
function showStartError(msg) { const e = $('#startError'); e.textContent = msg; e.hidden = !msg; if (msg) show(D ? 'app' : 'start'); }
const REPORT = /\.(xls|xlsx|xlsm|csv)$/i;
function readFiles(items) {
  items = items.filter((it) => REPORT.test(it.file.name) && !/^~\$/.test(it.file.name));
  if (!items.length) { showStartError('No report files there. Choose the .xls exports from Salesforce, or the folder that holds them.'); return; }
  if (typeof Worker === 'undefined') { showStartError('This browser cannot read the files. Use a current version of Edge, Chrome, Firefox or Safari.'); return; }
  showStartError('');
  const list = $('#progressList'), rows = new Map(), done = new Set();
  list.innerHTML = '';
  for (const it of items) {
    const li = document.createElement('li');
    li.innerHTML = `<div class="pf"><b>${esc(it.path)}</b><span>Waiting</span></div><div class="bar"><i></i></div>`;
    list.append(li); rows.set(it.path, li);
  }
  const count = () => { $('#progressCount').innerHTML = `<span class="cc-pcount">${done.size} of ${plural(items.length, 'file')} done.</span>`; };
  count(); show('progress');
  if (worker) worker.terminate();
  worker = new Worker(new URL('worker.js', import.meta.url));
  const KIND = { sent: 'emails sent', received: 'emails received', closed: 'cases closed', sla: 'SLA cases' };
  worker.onmessage = (e) => {
    const m = e.data;
    if (m.type === 'progress') {
      const li = rows.get(m.file); if (!li) return;
      li.querySelector('span').textContent = m.rows != null ? `${fmtN(m.rows)} ${KIND[m.kind] || 'rows'}` : m.stage;
      li.querySelector('i').style.width = m.pct + '%';
      li.classList.toggle('skip', !!m.skipped);
      if (m.pct >= 100) { done.add(m.file); count(); }
    } else {
      worker.terminate(); worker = null;
      if (m.type === 'done') { try { load(m.data); } catch (err) { showStartError(err.message); } }
      else showStartError(m.message);
    }
  };
  worker.onerror = (e) => { worker && worker.terminate(); worker = null; showStartError('Reading the files failed: ' + (e.message || 'unknown error') + '.'); };
  worker.postMessage({ files: items });
}
const fromInputEl = (fl) => [...(fl || [])].map((f) => ({ file: f, path: f.webkitRelativePath || f.name }));
$('#folderInput').addEventListener('change', (e) => { readFiles(fromInputEl(e.target.files)); e.target.value = ''; });
$('#filesInput').addEventListener('change', (e) => { readFiles(fromInputEl(e.target.files)); e.target.value = ''; });
$('#cancelBtn').addEventListener('click', () => { if (worker) { worker.terminate(); worker = null; } show(D ? 'app' : 'start'); });
$('#closeBtn').addEventListener('click', () => { D = null; V = null; L = null; show('start'); });

// Dropped folders are walked so each report keeps its team folder in its path.
async function dropItems(dtf) {
  const entries = [...(dtf.items || [])].map((i) => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null)).filter(Boolean);
  if (!entries.length) return fromInputEl(dtf.files);
  const out = [];
  const walk = async (entry, prefix) => {
    if (entry.isFile) { const f = await new Promise((ok, no) => entry.file(ok, no)); out.push({ file: f, path: prefix + f.name }); return; }
    const reader = entry.createReader(); let batch;
    do { batch = await new Promise((ok, no) => reader.readEntries(ok, no)); for (const c of batch) await walk(c, prefix + entry.name + '/'); } while (batch.length);
  };
  for (const e of entries) await walk(e, '');
  return out;
}
let dragDepth = 0;
document.addEventListener('dragenter', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) { e.preventDefault(); dragDepth++; document.body.classList.add('dragging'); } });
document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dragging'); } });
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault(); dragDepth = 0; document.body.classList.remove('dragging');
  if (worker || !e.dataTransfer) return;
  dropItems(e.dataTransfer).then(readFiles, (err) => showStartError('Those files could not be opened: ' + err.message));
});

// ---------- matching workbook names to Salesforce ----------
// Common nicknames, so BECKY finds Rebecca, PAT finds Patricia and BOB finds Robert.
const NICK = [['REBECCA', 'BECKY', 'BECCA'], ['ROBERT', 'BOB', 'ROB', 'BOBBY'], ['PATRICIA', 'PAT', 'PATTY', 'TRISH'], ['RAYMOND', 'RAY'],
  ['WILLIAM', 'BILL', 'WILL', 'BILLY'], ['ELIZABETH', 'LIZ', 'BETH', 'LIZZY'], ['KATHERINE', 'KATHRINE', 'KATHRYN', 'KATE', 'KATIE', 'KATHY'],
  ['SAMANTHA', 'SAM'], ['SAMUEL', 'SAM'], ['MICHAEL', 'MIKE'], ['JAMES', 'JIM', 'JIMMY'], ['THOMAS', 'TOM'], ['DANIEL', 'DAN', 'DANNY'],
  ['DAVID', 'DAVE'], ['CHRISTOPHER', 'CHRIS'], ['CHRISTINA', 'CHRIS', 'TINA'], ['JENNIFER', 'JEN', 'JENNY'], ['MEGAN', 'MEG'], ['ANTHONY', 'TONY'],
  ['NICHOLAS', 'NICK'], ['STEVEN', 'STEVE'], ['STEPHEN', 'STEVE'], ['JOSEPH', 'JOE'], ['JOSHUA', 'JOSH'], ['ALEXANDER', 'ALEX'], ['ALEXANDRA', 'ALEX'],
  ['DEBORAH', 'DEB', 'DEBBIE'], ['VICTORIA', 'VICKY', 'TORI'], ['ANDREW', 'ANDY', 'DREW'], ['MATTHEW', 'MATT'], ['RICHARD', 'RICK', 'RICH'],
  ['EDWARD', 'ED', 'EDDIE'], ['TIMOTHY', 'TIM'], ['SUSAN', 'SUE'], ['MARGARET', 'MAGGIE', 'PEGGY'], ['JACQUELINE', 'JACKIE'], ['ABIGAIL', 'ABBY'],
  ['BENJAMIN', 'BEN'], ['JONATHAN', 'JON'], ['JOHNATHON', 'JON'], ['KIMBERLY', 'KIM'], ['CYNTHIA', 'CINDY'], ['DONALD', 'DON'], ['RONALD', 'RON'],
  ['KENNETH', 'KEN'], ['GREGORY', 'GREG'], ['JEFFREY', 'JEFF'], ['JEFFERY', 'JEFF'], ['TERESA', 'TERRY'], ['FREDERICK', 'FRED']];
const toks = (s) => String(s).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
const nickOf = (w) => { const out = new Set([w]); for (const g of NICK) if (g.includes(w)) g.forEach((x) => out.add(x)); return out; };
const wordEq = (a, b) => a === b || (a.length >= 3 && b.startsWith(a)) || nickOf(a).has(b);
// every word of the workbook name matches a word of the Salesforce name ("ADRIANA" -> "Beatriz Adriana Ochoa")
const nameHit = (label, name) => { const l = toks(label), n = toks(name); return l.length > 0 && l.every((w) => n.some((t) => wordEq(w, t))); };
// workbook names that mean the same person ("Pat" on WEEKLY SLA = "PATRICIA" on the month tab)
const labelEq = (a, b) => { const x = toks(a).join(' '), y = toks(b).join(' '); return x === y || (x.length >= 3 && y.startsWith(x)) || (y.length >= 3 && x.startsWith(y)) || nickOf(x).has(y); };
const isWeekday = (d) => { const w = dowOf(d); return w > 0 && w < 6; };
// How well a person's daily counts fit the workbook's own numbers for that name, on the days both cover
// (0 = identical). Used when the name alone doesn't decide, e.g. BRENT or SHAY.
function fit(key, p) {
  const h = D.template && D.template.hist[key]; if (!h) return null;
  const pd = I.pday.get(p); let err = 0, base = 0, n = 0;
  [['act', 0], ['clo', 1]].forEach(([k, i]) => {
    for (const [ds, v] of Object.entries(h[k])) {
      const d = +ds; if (d < X.lo || d >= X.hi || !isWeekday(d)) continue;
      const a = pd && pd.get(d) ? pd.get(d)[i] : 0; err += Math.abs(v - a); base += v; n++;
    }
  });
  return n >= 4 && base > 0 ? err / base : null;
}
const activity = (p) => { let n = 0; const m = I.pday.get(p); if (m) for (const v of m.values()) n += v[0] + v[1]; return n; };
const personByName = (name) => D.people.findIndex((n) => n.toLowerCase() === String(name).trim().toLowerCase());

// L: the layout every workbook tab is drawn from.
//   sections [{ name, rows: [{ key, label, p, how }] }]   cust [{ label, p, mailbox, ceva, total, customers: [{ label, comps: [] }] }]
//   sla [{ title, rows: [{ label, p }] }]   open { statuses: [{ label, ids: [] }], rows: [{ label, p }] }
function buildLayout() {
  const T = D.template;
  L = { fromBook: !!(T && T.sections.length), sections: [], cust: [], sla: [], open: null };
  if (L.fromBook) {
    const all = D.people.map((_, i) => i);
    for (const s of T.sections) {
      const rows = s.labels.map((label) => ({ key: s.name + '|' + label, label, p: -1, how: 'Not found' }));
      const used = new Set();
      // 1. fixes from the Names tab or a ROSTER tab, 2. the name, 3. the numbers
      for (const r of rows) {
        if (OVR.has(r.key)) { r.p = OVR.get(r.key); r.how = 'Picked on this page'; }
        else {
          const ro = (T.roster || []).find((x) => labelEq(x.label, r.label) && (!x.section || labelEq(x.section, s.name)));
          const p = ro ? personByName(ro.name) : -1;
          if (p >= 0) { r.p = p; r.how = 'ROSTER tab'; }
          else {
            const cands = all.filter((i) => nameHit(r.label, D.people[i]));
            if (cands.length === 1) { r.p = cands[0]; r.how = 'Name'; }
            else if (cands.length > 1) {
              const scored = cands.map((i) => [i, fit(r.key, i)]).filter((x) => x[1] != null).sort((a, b) => a[1] - b[1]);
              r.p = scored.length ? scored[0][0] : cands.sort((a, b) => activity(b) - activity(a))[0];
              r.how = scored.length ? 'Name and numbers' : 'Name (several match: check)';
            }
          }
        }
        if (r.p >= 0) used.add(r.p);
      }
      for (const r of rows) {
        if (r.p >= 0 || OVR.has(r.key)) continue;
        const best = all.filter((i) => !used.has(i)).map((i) => [i, fit(r.key, i)]).filter((x) => x[1] != null && x[1] <= 0.35).sort((a, b) => a[1] - b[1])[0];
        if (best) { r.p = best[0]; r.how = 'Numbers match'; used.add(best[0]); }
      }
      L.sections.push({ name: s.name, rows });
    }
    const findRow = (sec, label) => { const s = L.sections.find((x) => !sec || x.name === sec); const pool = s ? [s] : L.sections; for (const x of pool) { const r = x.rows.find((y) => labelEq(y.label, label)); if (r) return r; } return null; };
    const anyRow = (label) => { for (const s of L.sections) { const r = s.rows.find((y) => labelEq(y.label, label)); if (r) return r; } return null; };
    // CUST E-MAILS blocks: the rep comes from the E-MAILS ACTIONED formula, else from the name
    for (const b of T.cust) {
      const r = (b.ref && findRow(b.ref.section, b.ref.label)) || anyRow(b.label), p = r ? r.p : -1, allowed = teamCompanies(p);
      L.cust.push({ label: b.label, p, mailbox: b.mailbox, ceva: b.ceva, total: b.total || !b.mailbox,
        customers: b.customers.map((c) => ({ label: c, comps: matchCompanies(c, allowed) })) });
    }
    // WEEKLY SLA lists: each list belongs to the section where most of its names are found
    for (const list of T.sla) {
      const sec = L.sections.map((s) => [s, list.labels.filter((l) => s.rows.some((r) => labelEq(r.label, l))).length]).sort((a, b) => b[1] - a[1])[0];
      L.sla.push({ title: list.title || (sec ? sec[0].name : ''), rows: list.labels.map((l) => {
        const r = sec && sec[1] ? sec[0].rows.find((x) => labelEq(x.label, l)) : null;
        let p = r ? r.p : -1;
        if (!r) { const c = D.people.map((_, i) => i).filter((i) => nameHit(l, D.people[i])); if (c.length === 1) p = c[0]; }
        return { label: l, p };
      }) });
    }
    if (T.open) L.open = { statuses: T.open.statuses.map((s) => ({ label: s, ids: statusIds(s) })), rows: T.open.names.map((n) => {
      let p = personByName(n); if (p < 0) { const c = D.people.map((_, i) => i).filter((i) => nameHit(n, D.people[i])); if (c.length === 1) p = c[0]; }
      return { label: n, p };
    }) };
  } else {
    // Without the workbook: one section per team folder, full names A to Z.
    const by = new Map();
    for (const tbl of [D.act, D.clo, D.inb, D.sla]) for (const r of tbl) if (r[1] >= 0) { if (!by.has(r[0])) by.set(r[0], new Set()); by.get(r[0]).add(r[1]); }
    for (const t of [...by.keys()].sort(teamOrder)) {
      const ps = [...by.get(t)].sort((a, b) => byText(D.people[a], D.people[b]));
      L.sections.push({ name: D.teams[t].toUpperCase(), rows: ps.map((p) => ({ key: D.teams[t] + '|' + D.people[p], label: D.people[p], p, how: 'Salesforce name' })) });
      for (const p of ps) {
        const comps = [...new Set(D.inb.filter((r) => r[1] === p && !r[4]).map((r) => r[2]))].sort((a, b) => byText(D.companies[a], D.companies[b]));
        L.cust.push({ label: D.people[p], p, mailbox: false, ceva: true, total: true, ownerOnly: true, customers: comps.map((c) => ({ label: D.companies[c], comps: [c] })) });
      }
      L.sla.push({ title: D.teams[t], rows: ps.filter((p) => I.slaw.has(p)).map((p) => ({ label: D.people[p], p })) });
    }
  }
  if (!L.open) {
    const ps = [...I.open.keys()].sort((a, b) => byText(D.people[a], D.people[b]));
    L.open = { statuses: D.statuses.map((s, i) => ({ label: s, ids: [i] })).sort((a, b) => statusRank(a.label) - statusRank(b.label)), rows: ps.map((p) => ({ label: D.people[p], p })) };
  }
}
const STATUS_ORDER = ['new', 're opened', 'answer received', 'in progress', 'on hold', 'escalated'];
const sNorm = (s) => String(s).toLowerCase().replace(/[^a-z]+/g, ' ').trim();
const statusRank = (s) => { const i = STATUS_ORDER.findIndex((x) => sNorm(s).startsWith(x)); return i < 0 ? 99 : i; };
// "In progress - Follow up needed" on the sheet = "In Progress" in Salesforce
const statusIds = (label) => D.statuses.map((s, i) => [s, i]).filter(([s]) => { const a = sNorm(s), b = sNorm(label); return a && b && (a.startsWith(b) || b.startsWith(a)); }).map((x) => x[1]);

// Customer rows on CUST E-MAILS use the workbook's own names ("BEACH CAMERA"); they are matched to
// Salesforce company names by their words, rare words counting more. A company is counted when most of
// the row's words and most of the company's own words agree ("FREIGHT CLUB" is Freight Club, not BJ's
// Wholesale Club). If nothing agrees that well, companies containing every known word of the row, starting
// with its first word, are used ("LKQ" -> LKQ Corporate Headquarters). The Names tab lists every match.
const STOP = new Set(['INC', 'LLC', 'LTD', 'CORP', 'CORPORATION', 'CO', 'COMPANY', 'THE', 'OF', 'AND', 'DBA', 'PARENT', 'HOLDINGS', 'GROUP', 'US', 'USA']);
let CT = null;
const ctoks = (s) => toks(s).filter((w) => w.length >= 2 && !STOP.has(w));
// same word, an abbreviation of it ("HSN" / "HSNI"), a plural ("HOMES" / "HOME") or a one-letter typo ("WALTZ" / "WALTS")
const typo = (a, b) => a.length >= 5 && a.length === b.length && [...a].filter((x, i) => x !== b[i]).length === 1;
const tokEq = (l, c) => l === c || (l.length >= 3 && c.startsWith(l) && c.length - l.length <= 2) || (c.length >= 4 && l.startsWith(c) && l.length - c.length <= 1) || typo(l, c);
// Only companies with emails in the rep's own team folder are considered.
function teamCompanies(p) {
  if (p < 0) return null;
  const row = D.act.find((r) => r[1] === p) || D.clo.find((r) => r[1] === p) || D.inb.find((r) => r[1] === p);
  return row ? new Set(D.inb.filter((r) => r[0] === row[0]).map((r) => r[2])) : null;
}
function matchCompanies(label, allowed) {
  if (!CT || CT.src !== D) {
    const t = D.companies.map(ctoks), df = new Map();
    for (const ts of t) for (const w of new Set(ts)) df.set(w, (df.get(w) || 0) + 1);
    CT = { src: D, t, joined: D.companies.map((c) => toks(c).join('')), idfC: (w) => Math.log(1 + t.length / (df.get(w) || 1)) };
  }
  const N = CT.t.length, words = ctoks(label), joined = toks(label).join('');
  const seen = words.map((w) => [w, CT.t.reduce((n, ts) => n + (ts.some((c) => tokEq(w, c)) ? 1 : 0), 0)]).filter((x) => x[1] > 0).map(([w, n]) => [w, Math.log(1 + N / n)]);
  const strong = [], loose = [];
  CT.t.forEach((ts, c) => {
    if (allowed && !allowed.has(c)) return;
    if (joined.length >= 5 && CT.joined[c].includes(joined)) { strong.push(c); return; }   // "SHIP DADDY" = ShipDaddy
    if (!seen.length || !ts.length) return;
    let lm = 0, lt = 0, all = true;
    for (const [w, idf] of seen) { lt += idf; if (ts.some((x) => tokEq(w, x))) lm += idf; else all = false; }
    let cm = 0, ctot = 0;
    for (const x of ts) { const i = CT.idfC(x); ctot += i; if (seen.some(([w]) => tokEq(w, x))) cm += i; }
    if (lm / lt >= 0.5 && cm / ctot >= 0.5 && seen.some(([w]) => w === words[0])) strong.push(c);
    else if (all && words.length && seen[0][0] === words[0]) loose.push(c);
  });
  return strong.length ? strong : loose;
}

// ---------- tabs ----------
function buildTabs() {
  const tabs = [];
  for (const k of I.months) tabs.push([`m${k}`, sheetName(k)], [`c${k}`, `CUST E-MAILS ${SHEET_MON[k % 12]}`]);
  tabs.push(['wsla', 'WEEKLY SLA'], ['open', 'Open cases'], ['over', 'Overview'], ['names', 'Names & sources']);
  $('#tabs').innerHTML = tabs.map(([id, label]) => `<button class="tab" role="tab" type="button" data-tab="${id}">${esc(label)}</button>`).join('');
  TABS = tabs.map((t) => t[0]);
}
let TABS = [];
function setTab(t) {
  // opens on the newest month tab, like the workbook
  S.tab = TABS.includes(t) ? t : (I.months.length ? `m${I.months[I.months.length - 1]}` : 'over');
  document.querySelectorAll('#tabs .tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === S.tab)));
  const panel = S.tab[0] === 'm' ? 'month' : S.tab[0] === 'c' ? 'cust' : S.tab;
  document.querySelectorAll('.panel').forEach((p) => (p.hidden = p.id !== 'p-' + panel));
  $('.cc-filters').hidden = S.tab !== 'over';
  saveView(); render();
}
$('#tabs').addEventListener('click', (e) => { const b = e.target.closest('.tab'); if (b) setTab(b.dataset.tab); });

// ---------- tables: one model renders to HTML, copies as text and exports to Excel ----------
// model: { cls, cols: [{ h (html), x (plain), cls, d (day) }], rows: [{ cls, c: [cell] }] }; cell: { v, f, cls, title }
// f: 's' text, 'n' count (blank when 0), 'z' count (0 shown), 'p' percent, 'hm' minutes shown as h:mm
const cS = (v, cls, title) => ({ v, f: 's', cls, title }), cN = (v, cls) => ({ v, f: 'n', cls }), cZ = (v, cls) => ({ v, f: 'z', cls }), cP = (v, cls) => ({ v, f: 'p', cls }), cHM = (v, cls) => ({ v, f: 'hm', cls });
const goalCls = (p) => (p == null ? '' : p < GOAL ? 'low' : 'ok');
function cellHtml(c) {
  if (c.f === 's') return esc(c.v ?? '');
  if (c.v == null || (c.f === 'n' && !c.v)) return '';
  return c.f === 'n' || c.f === 'z' ? fmtN(c.v) : c.f === 'p' ? fmtP(c.v) : fmtHM(c.v);
}
const emptyMsg = (t) => `<div class="empty-s">${t}</div>`;
function renderTable(el, m, scrollEnd) {
  if (!el) return;
  if (!m.rows.length) { el.innerHTML = emptyMsg(m.empty || 'Nothing for these filters.'); return; }
  const th = m.cols.map((c) => `<th class="${c.cls || ''}">${c.h}</th>`).join('');
  const body = m.rows.map((r) => `<tr class="${r.cls || ''}">${r.c.map((c, i) => `<td class="${[m.cols[i] && m.cols[i].cls, c.cls].filter(Boolean).join(' ')}"${c.title ? ` title="${esc(c.title)}"` : ''}>${cellHtml(c)}</td>`).join('')}</tr>`).join('');
  el.innerHTML = `<table class="${m.cls || ''}"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
  if (scrollEnd) el.scrollLeft = el.scrollWidth;
}
const col = (h, cls, x) => ({ h, cls, x: x ?? h.replace(/<small>/g, ' ').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&') });
const dayCol = (d) => ({ ...col(`${DOW[dowOf(d)].toUpperCase()}<small>${md(d)}</small>`, '', md(d)), d });
const sumBy = (days, f) => { let s = 0; for (const d of days) s += f(d) || 0; return s; };

// ---------- month tab (CASES ACTIONED / CASES CLOSED) ----------
// The weekdays of the month the exports cover, a TOTAL after each week; days outside the exports are blank.
function monthWeeks(k) {
  const a = Math.max(monthStart(k), monday(X.lo)), b = Math.min(monthEnd(k), monday(X.hi) + 4), weeks = [];
  for (let w = monday(a); w <= b; w += 7) {
    const days = []; for (let d = w; d < w + 5; d++) if (d >= monthStart(k) && d <= monthEnd(k)) days.push(d);
    if (days.length) weeks.push(days);
  }
  return weeks;
}
const covered = (d) => d >= X.lo && d <= X.hi;
const pv = (p, d, i) => { const m = I.pday.get(p), v = m && m.get(d); return v ? v[i] : 0; };
function monthModel(k, sec, i) {
  const weeks = monthWeeks(k), cols = [col(i ? 'CASES CLOSED' : 'CASES ACTIONED')];
  for (const w of weeks) { for (const d of w) cols.push(dayCol(d)); cols.push({ ...col('TOTAL', 'wk'), wk: w }); }
  const cell = (p, d) => (p < 0 || !covered(d) ? cS('') : cZ(pv(p, d, i)));
  const line = (r) => [cS(r.label, r.p < 0 ? 'muted' : '', r.p >= 0 ? D.people[r.p] : 'No Salesforce match: fix it on the Names & sources tab'),
    ...cols.slice(1).map((c) => (c.d != null ? cell(r.p, c.d) : r.p < 0 || !c.wk.some(covered) ? cS('') : cZ(sumBy(c.wk.filter(covered), (d) => pv(r.p, d, i)))))];
  const rows = sec.rows.map((r) => ({ c: line(r) }));
  const tot = cols.slice(1).map((c) => { const ds = (c.d != null ? [c.d] : c.wk).filter(covered); return ds.length ? cZ(sec.rows.reduce((a, r) => a + (r.p < 0 ? 0 : sumBy(ds, (d) => pv(r.p, d, i))), 0)) : cS(''); });
  rows.push({ cls: 'total', c: [cS('TOTALS'), ...tot] });
  return { cls: 'mx mirror', cols, rows };
}
function renderMonth(k) {
  const weeks = monthWeeks(k), first = weeks.length ? weeks[0][0] : null;
  const gap = first != null && first > monthStart(k) && monday(first) > monthStart(k) ? ` Days before ${md(first)} aren't in these exports.` : '';
  $('#p-month').innerHTML = `<p class="note">${esc(sheetName(k))} from the exports.${esc(gap)} Hover a name to see who it is in Salesforce.</p>` +
    L.sections.map((s, j) => `<div class="card"><div class="card-h"><h2>${esc(s.name)}</h2></div>
      ${[0, 1].map((i) => `<div class="card-h"><span class="sub">${i ? 'CASES CLOSED' : 'CASES ACTIONED'}</span><button class="cc-linkbtn copybtn" type="button" data-copy="#mt-${j}-${i}" title="Copy this table, then paste into Excel">Copy</button></div><div class="tscroll" id="mt-${j}-${i}"></div>`).join('')}
    </div>`).join('');
  L.sections.forEach((s, j) => [0, 1].forEach((i) => renderTable($(`#mt-${j}-${i}`), monthModel(k, s, i), true)));
}

// ---------- CUST E-MAILS tab ----------
// Per rep: each customer row = every email received for those companies; E-MAILS FROM CEVA STATIONS =
// emails from CEVA addresses on the rep's cases; E-MAILS ACTIONED = the rep's actioned + closed.
function custLines(k) {
  const days = monthWeeks(k).flat(), out = [];
  const compDay = (cs, d) => cs.reduce((a, c) => { const m = I.comp.get(c); return a + ((m && m.get(d)) || 0); }, 0);
  const ownDay = (p, cs, d) => cs.reduce((a, c) => { const m = I.compOwn.get(p + ',' + c); return a + ((m && m.get(d)) || 0); }, 0);
  for (const b of L.cust) {
    const custF = b.customers.map((c) => ({ label: c.label, comps: c.comps, f: (d) => (b.ownerOnly ? ownDay(b.p, c.comps, d) : compDay(c.comps, d)) }));
    const ceva = (d) => (b.ceva && b.p >= 0 ? pv(b.p, d, 2) : 0);
    const total = (d) => custF.reduce((a, c) => a + c.f(d), 0) + ceva(d);
    const act = (d) => (b.p >= 0 ? pv(b.p, d, 0) + pv(b.p, d, 1) : 0);
    if (!b.mailbox) out.push({ kind: 'rep', label: b.label, p: b.p });
    for (const c of custF) out.push({ kind: b.mailbox ? 'mbx' : 'cust', label: b.mailbox ? '' : c.label, title: c.comps.length ? c.comps.map((x) => D.companies[x]).join(', ') : 'No Salesforce company matched', f: c.f, head: b.mailbox ? b.label : null });
    if (b.ceva) out.push({ kind: 'ceva', label: 'E-MAILS FROM CEVA STATIONS', f: ceva });
    if (b.total) out.push({ kind: 'total', label: 'TOTAL', f: total });
    out.push({ kind: 'act', label: 'E-MAILS ACTIONED', f: act, title: b.p >= 0 ? `${D.people[b.p]}: emails actioned + cases closed` : 'Rep not matched' });
    out.push({ kind: 'pct', label: '% ACTIONED (GOAL 85%)', f: (d) => [act(d), total(d)] });
  }
  return { days, lines: out };
}
function custModel(k) {
  const { days, lines } = custLines(k), cdays = days.filter(covered);
  const cols = [col('CUSTOMER / REP'), ...days.map(dayCol), col('TOTALS', 'wk')];
  const rows = [];
  for (const ln of lines) {
    if (ln.kind === 'rep') { rows.push({ cls: 'grp', c: [cS(ln.label, '', ln.p >= 0 ? D.people[ln.p] : 'Rep not matched'), ...cols.slice(1).map(() => cS(''))] }); continue; }
    if (ln.kind === 'mbx') rows.push({ cls: 'grp', c: [cS(ln.head), ...cols.slice(1).map(() => cS(''))] });
    if (ln.kind === 'pct') {
      const p = (v) => (v && v[1] ? v[0] / v[1] : null), all = cdays.reduce((a, d) => { const v = ln.f(d); return [a[0] + v[0], a[1] + v[1]]; }, [0, 0]);
      rows.push({ cls: 'key', c: [cS(ln.label), ...days.map((d) => { if (!covered(d)) return cS(''); const x = p(ln.f(d)); return cP(x, goalCls(x)); }), cP(p(all), goalCls(p(all)))] });
      continue;
    }
    rows.push({ cls: ln.kind === 'total' ? 'sub' : ln.kind === 'act' || ln.kind === 'ceva' ? 'key' : '',
      c: [cS(ln.label || 'E-mails received', ln.kind === 'mbx' ? 'muted' : '', ln.title), ...days.map((d) => (covered(d) ? cZ(ln.f(d)) : cS(''))), cZ(sumBy(cdays, ln.f))] });
  }
  return { cls: 'mx mirror', cols, rows, empty: 'No CUST E-MAILS layout.' };
}
function renderCust(k) {
  $('#p-cust').innerHTML = `<div class="card"><div class="card-h"><h2>CUST E-MAILS ${esc(SHEET_MON[k % 12])}</h2><span class="sub">goal 85%; hover a customer to see the Salesforce companies counted</span><button class="cc-linkbtn copybtn" type="button" data-copy="#mCust" title="Copy this table, then paste into Excel">Copy</button></div><div class="tscroll tall" id="mCust"></div></div>`;
  renderTable($('#mCust'), custModel(k), true);
}

// ---------- WEEKLY SLA ----------
// RUN DATE = the Monday after the week the cases were opened; the figure is the average first
// response (Elapsed Time) of that person's cases, h:mm. TEAM AVERAGE averages the people with cases.
const slaOf = (p, w) => { const m = I.slaw.get(p), v = m && m.get(w); return v && v[0] ? v : null; };
function slaModel(list) {
  const weeks = I.weeks, cols = [col('CSR'), ...weeks.map((w) => ({ ...col(`RUN DATE<small>${md(w + 7)}</small>`, '', `RUN DATE ${md(w + 7)}`), d: w + 7 }))];
  const rows = list.rows.map((r) => ({ c: [cS(r.label, r.p < 0 ? 'muted' : '', r.p >= 0 ? D.people[r.p] : 'No Salesforce match'), ...weeks.map((w) => { if (r.p < 0) return cS(''); const v = slaOf(r.p, w); return cHM(v ? v[1] / v[0] : 0); })] }));
  rows.push({ cls: 'total', c: [cS('TEAM AVERAGE'), ...weeks.map((w) => { const xs = list.rows.map((r) => (r.p >= 0 ? slaOf(r.p, w) : null)).filter(Boolean).map((v) => v[1] / v[0]); return xs.length ? cHM(xs.reduce((a, b) => a + b, 0) / xs.length) : cS(''); })] });
  return { cls: 'mx mirror', cols, rows, empty: 'No SLA report loaded.' };
}
function renderSla() {
  const lists = L.sla.filter((l) => l.rows.length);
  $('#p-wsla').innerHTML = !D.sla.length ? `<div class="card">${emptyMsg('Load the SLA report (the export with Elapsed Time and SLA Breached?) to fill WEEKLY SLA.')}</div>` :
    `<p class="note">SLA 1ST RESPONSE = average time to first response (h:mm) of the cases opened in the week before each run date. Newest run date on the right.</p>` +
    lists.map((l, j) => `<div class="card"><div class="card-h"><h2>${esc(l.title || 'SLA')}</h2><span class="sub">SLA 1ST RESPONSE</span><button class="cc-linkbtn copybtn" type="button" data-copy="#ws-${j}" title="Copy this table, then paste into Excel">Copy</button></div><div class="tscroll" id="ws-${j}"></div></div>`).join('');
  lists.forEach((l, j) => renderTable($(`#ws-${j}`), slaModel(l), true));
}

// ---------- Open cases ----------
function openModel() {
  const cols = [col('CSR'), ...L.open.statuses.map((s) => col(esc(s.label))), col('Total Per person', 'wk')];
  const n = (p, s) => (p < 0 ? 0 : s.ids.reduce((a, i) => a + ((I.open.get(p) && I.open.get(p).get(i)) || 0), 0));
  const rows = L.open.rows.map((r) => { const v = L.open.statuses.map((s) => n(r.p, s)); return { c: [cS(r.label, r.p < 0 ? 'muted' : '', r.p >= 0 ? D.people[r.p] : 'No Salesforce match'), ...v.map((x) => cN(x)), cZ(v.reduce((a, b) => a + b, 0))] }; });
  const tot = L.open.statuses.map((s) => L.open.rows.reduce((a, r) => a + n(r.p, s), 0));
  rows.push({ cls: 'total', c: [cS('Totals per Status'), ...tot.map((x) => cZ(x)), cZ(tot.reduce((a, b) => a + b, 0))] });
  return { cols, rows, empty: 'No open cases.' };
}
function renderOpen() {
  $('#p-open').innerHTML = `<div class="card"><div class="card-h"><h2>Open cases</h2><span class="sub">cases in the SLA export that aren't closed yet${D.files.some((f) => f.kind === 'sla') ? ` (opened ${span(X.slaLo, X.slaHi - 6 > X.slaLo ? X.slaHi - 1 : X.slaHi)})` : ''}</span><button class="cc-linkbtn copybtn" type="button" data-copy="#tOpen" title="Copy this table, then paste into Excel">Copy</button></div><div class="tscroll" id="tOpen"></div></div>`;
  renderTable($('#tOpen'), openModel());
}

// ---------- Names & sources ----------
function renderNames() {
  const opts = (p) => `<option value="-1">(none)</option>` + D.people.map((n, i) => [n, i]).sort((a, b) => byText(a[0], b[0])).map(([n, i]) => `<option value="${i}"${i === p ? ' selected' : ''}>${esc(n)}</option>`).join('');
  $('#namesNote').innerHTML = L.fromBook
    ? `Names come from the <b>${esc(D.template.month)}</b> tab. Each is matched to a Salesforce case owner by name, nickname, or by comparing the workbook's own numbers with the exports. Pick a different person to fix a match; fixes last until the page is closed. To keep them, add a <b>ROSTER</b> tab to the workbook with the workbook name in column A and the Salesforce name in column B.`
    : 'The workbook wasn\'t loaded, so each team folder is a section and full Salesforce names are used.';
  $('#tNames').innerHTML = `<table><thead><tr><th>Section</th><th>Name in workbook</th><th>Salesforce case owner</th><th>Matched by</th></tr></thead><tbody>` +
    L.sections.flatMap((s) => s.rows.map((r) => `<tr><td>${esc(s.name)}</td><td>${esc(r.label)}</td><td>${L.fromBook ? `<select data-key="${esc(r.key)}" aria-label="Salesforce name for ${esc(r.label)}">${opts(r.p)}</select>` : esc(pName(r.p))}</td><td class="${r.p < 0 || /check/.test(r.how) ? 'low' : ''}">${esc(r.how)}</td></tr>`)).join('') + '</tbody></table>';
  $('#tCustMap').innerHTML = L.cust.length ? `<table><thead><tr><th>Rep</th><th>Customer row</th><th>Salesforce companies counted</th></tr></thead><tbody>` +
    L.cust.flatMap((b) => b.customers.map((c) => `<tr><td>${esc(b.label)}</td><td>${esc(c.label)}</td><td class="${c.comps.length ? '' : 'low'}" style="white-space:normal;text-align:left">${c.comps.length ? esc(c.comps.map((x) => D.companies[x]).join(' · ')) : 'none matched'}</td></tr>`)).join('') + '</tbody></table>' : emptyMsg('No customer rows.');
  about();
}
$('#p-names').addEventListener('change', (e) => {
  const s = e.target.closest('select[data-key]'); if (!s) return;
  OVR.set(s.dataset.key, +s.value); buildLayout(); renderNames();
});
const KIND = { sent: 'Emails sent (actioned)', received: 'Emails received', closed: 'Cases closed', sla: 'SLA', workbook: 'Workbook layout' };
function about() {
  $('#about').innerHTML = `<p>How each tab is filled from the Salesforce exports:</p>
    <ul>
      <li><b>CASES ACTIONED</b>: emails sent on the cases a person owns, on the email's date (every row of the sent-emails report).</li>
      <li><b>CASES CLOSED</b>: cases whose Date/Time Closed falls on that day, credited to the case's current owner. Salesforce keeps only a case's latest close, so closures that were later reopened, or cases now owned by a queue, aren't counted; this is why closed can run lower than numbers typed in during the day.</li>
      <li><b>CUST E-MAILS</b>: each customer row counts every email received for the Salesforce companies matched to it (listed below). <b>E-MAILS FROM CEVA STATIONS</b> are emails from @cevalogistics.com addresses on the rep's cases. <b>E-MAILS ACTIONED</b> = the rep's actioned + closed; <b>% ACTIONED</b> = that ÷ TOTAL, goal 85%.</li>
      <li><b>WEEKLY SLA</b>: the average Elapsed Time to first response of the cases opened in the week before each run date (h:mm); TEAM AVERAGE averages the people who had cases.</li>
      <li><b>Open cases</b>: cases in the SLA export whose status isn't Closed.</li>
      <li>Only weekdays are shown, like the workbook. PTO isn't in Salesforce, so a day off shows 0. Days the exports don't cover are blank. The KASEY ship counts and the EMAIL VS SHIP COUNT and New Cases tabs aren't in these exports.</li>
    </ul>`;
  renderTable($('#tFiles'), { cols: [col('File'), col('Team'), col('Report'), col('Rows'), col('Dates')],
    rows: D.files.map((f) => ({ cls: f.kind ? '' : 'muted', c: [cS(f.name), cS(f.team || '–'), cS(f.kind ? (f.kind === 'workbook' ? `${KIND.workbook}: ${f.note}` : KIND[f.kind]) : f.note || 'Skipped'), cN(f.rows), cS(f.from != null ? span(f.from, f.to) + (f.kind === 'sla' ? ' (opened)' : '') : '')] })) });
}

// ---------- Overview (filters apply here only) ----------
function buildFilters() {
  $('#teamChips').innerHTML = D.teams.map((t, i) => [t, i]).sort((a, b) => byText(a[0], b[0]))
    .map(([t, i]) => `<button class="chip" type="button" data-t="${i}" aria-pressed="${S.teams.includes(i)}">${esc(t)}</button>`).join('');
  fillPeople();
  for (const id of ['#fFrom', '#fTo']) { $(id).min = toInput(X.from); $(id).max = toInput(X.to); }
  $('#fFrom').value = toInput(S.from); $('#fTo').value = toInput(S.to);
}
function fillPeople() {
  const ts = new Set(S.teams), by = new Map();
  for (const tbl of [D.act, D.clo, D.inb, D.sla, D.open]) for (const r of tbl) if (r[1] >= 0 && ts.has(r[0])) { if (!by.has(r[0])) by.set(r[0], new Set()); by.get(r[0]).add(r[1]); }
  if (S.person >= 0 && ![...by.values()].some((s) => s.has(S.person))) S.person = -1;
  $('#fPerson').innerHTML = '<option value="-1">All people</option>' + [...by.keys()].sort(teamOrder).map((t) =>
    `<optgroup label="${esc(D.teams[t])}">${[...by.get(t)].sort((a, b) => byText(D.people[a], D.people[b])).map((p) => `<option value="${p}">${esc(D.people[p])}</option>`).join('')}</optgroup>`).join('');
  $('#fPerson').value = String(S.person);
}
$('#teamChips').addEventListener('click', (e) => {
  const b = e.target.closest('.chip'); if (!b) return;
  const i = +b.dataset.t, on = new Set(S.teams);
  if (on.has(i)) on.delete(i); else on.add(i);
  S.teams = on.size ? [...on] : D.teams.map((_, k) => k);
  document.querySelectorAll('#teamChips .chip').forEach((c) => c.setAttribute('aria-pressed', String(S.teams.includes(+c.dataset.t))));
  fillPeople(); render();
});
$('#fPerson').addEventListener('change', (e) => { S.person = +e.target.value; render(); });
$('#fFrom').addEventListener('change', (e) => { const d = fromInput(e.target.value); S.from = d == null ? X.from : d; if (S.from > S.to) { S.to = S.from; $('#fTo').value = toInput(S.to); } render(); });
$('#fTo').addEventListener('change', (e) => { const d = fromInput(e.target.value); S.to = d == null ? X.to : d; if (S.to < S.from) { S.from = S.to; $('#fFrom').value = toInput(S.from); } render(); });
$('#resetBtn').addEventListener('click', () => { Object.assign(S, dflt()); buildFilters(); render(); });

// V.pd: per team+person, a Map of day -> [actioned, closed, received, receivedFromCeva]
function aggregate() {
  const ts = new Set(S.teams), ok = (r) => ts.has(r[0]) && (S.person < 0 || r[1] === S.person), inR = (d) => d >= S.from && d <= S.to;
  const pd = new Map(), days = new Set();
  const slot = (t, p, d) => {
    const k = t + ',' + p; let o = pd.get(k);
    if (!o) pd.set(k, (o = { t, p, days: new Map() }));
    let v = o.days.get(d); if (!v) o.days.set(d, (v = [0, 0, 0, 0]));
    days.add(d); return v;
  };
  for (const r of D.act) if (ok(r) && inR(r[2])) slot(r[0], r[1], r[2])[0] += r[3];
  for (const r of D.clo) if (ok(r) && inR(r[2])) slot(r[0], r[1], r[2])[1] += r[3];
  for (const r of D.inb) if (ok(r) && inR(r[3])) { const v = slot(r[0], r[1], r[3]); v[2] += r[5]; if (r[4]) v[3] += r[5]; }
  const sla = new Map();
  for (const r of D.sla) {
    if (!ok(r) || r[2] + 6 < S.from || r[2] > S.to) continue;
    const k = r[0] + ',' + r[1]; let o = sla.get(k);
    if (!o) sla.set(k, (o = { t: r[0], p: r[1], wk: new Map() }));
    const v = o.wk.get(r[2]) || [0, 0, 0, 0]; for (let i = 0; i < 4; i++) v[i] += r[3 + i]; o.wk.set(r[2], v);
  }
  const open = new Map();
  for (const r of D.open) {
    if (!ok(r)) continue;
    const k = r[0] + ',' + r[1]; let o = open.get(k);
    if (!o) open.set(k, (o = { t: r[0], p: r[1], st: new Map() }));
    o.st.set(r[2], (o.st.get(r[2]) || 0) + r[3]);
  }
  V = { pd, days: [...days].sort((a, b) => a - b), sla, open };
}
const addVec = (a, b) => { if (!b) return a; if (!a) return b.slice(); for (let i = 0; i < b.length; i++) a[i] += b[i]; return a; };
const mergeMaps = (maps) => { const out = new Map(); for (const m of maps) for (const [k, v] of m) out.set(k, addVec(out.get(k), v)); return out; };
const pctOf = (v) => (v && v[2] ? (v[0] + v[1]) / v[2] : null);
function grouped(keys, rowOf, subOf, totalOf) {
  const byT = new Map();
  for (const k of keys) { if (!byT.has(k.t)) byT.set(k.t, []); byT.get(k.t).push(k); }
  const ts = [...byT.keys()].sort(teamOrder), multi = ts.length > 1, out = [];
  for (const t of ts) {
    const ks = byT.get(t).sort(personOrder);
    if (multi) out.push({ group: D.teams[t] });
    for (const k of ks) out.push(rowOf(k));
    if (multi) out.push({ ...subOf(t, ks), cls: 'sub' });
  }
  if (keys.length > 1) out.push({ ...totalOf(keys, multi), cls: 'total' });
  return out;
}
const groupRow = (label, n) => ({ cls: 'grp', c: [cS(label), ...Array.from({ length: n - 1 }, () => cS(''))] });
function totals(keysPd, keysSla, keysOpen) {
  const t = { act: 0, clo: 0, recv: 0, ceva: 0, cases: 0, mins: 0, known: 0, met: 0, open: 0 };
  for (const k of keysPd) for (const v of k.days.values()) { t.act += v[0]; t.clo += v[1]; t.recv += v[2]; t.ceva += v[3]; }
  for (const k of keysSla) for (const v of k.wk.values()) { t.cases += v[0]; t.mins += v[1]; t.known += v[2]; t.met += v[3]; }
  for (const k of keysOpen) for (const n of k.st.values()) t.open += n;
  return t;
}
function tile(label, value, sub) { return `<div class="ktile"><div class="eyebrow">${label}</div><div class="v">${value}</div><div class="d">${sub}</div></div>`; }
function renderTiles() {
  const t = totals(V.pd.values(), V.sla.values(), V.open.values()), p = t.recv ? (t.act + t.clo) / t.recv : null;
  const pill = p == null ? '' : `<span class="pill ${p >= GOAL ? 'ok' : 'bad'}">${p >= GOAL ? 'meets' : 'below'} the 85% goal</span>`;
  $('#tiles').innerHTML =
    tile('Emails received', fmtN(t.recv), `${fmtN(t.recv - t.ceva)} from customers · ${fmtN(t.ceva)} from CEVA stations`) +
    tile('Emails actioned', fmtN(t.act), 'emails sent on the cases') +
    tile('Cases closed', fmtN(t.clo), 'by the day they closed') +
    tile('% actioned', fmtP(p), pill || 'no emails received') +
    tile('SLA first response', t.cases ? fmtHM(t.mins / t.cases) : '–', t.cases ? `average · ${fmtP(t.known ? t.met / t.known : null)} within SLA` : 'no SLA report loaded') +
    tile('Open cases', fmtN(t.open), 'not closed, in the SLA export');
}
function summaryModel(byTeam) {
  const cols = [col(byTeam ? 'Team' : 'Person'), ...(byTeam ? [col('People')] : []), col('Emails received'), col('From customers'), col('From CEVA stations'),
    col('Emails actioned'), col('Cases closed'), col('Actioned + closed'), col('% actioned'), col('SLA first response'), col('Within SLA'), col('Open cases')];
  const line = (label, pd, sla, open, lcls) => {
    const t = totals(pd, sla, open), p = t.recv ? (t.act + t.clo) / t.recv : null;
    return [cS(label, lcls), cN(t.recv), cN(t.recv - t.ceva), cN(t.ceva), cN(t.act), cN(t.clo), cN(t.act + t.clo), cP(p, goalCls(p)),
      cHM(t.cases ? t.mins / t.cases : null), cP(t.known ? t.met / t.known : null), cN(t.open)];
  };
  const all = new Map(), keyOf = (o) => o.t + ',' + o.p;
  for (const m of [V.pd, V.sla, V.open]) for (const o of m.values()) if (!all.has(keyOf(o))) all.set(keyOf(o), { t: o.t, p: o.p });
  const parts = (ks) => [ks.map((k) => V.pd.get(keyOf(k))).filter(Boolean), ks.map((k) => V.sla.get(keyOf(k))).filter(Boolean), ks.map((k) => V.open.get(keyOf(k))).filter(Boolean)];
  if (byTeam) {
    const byT = new Map(); for (const k of all.values()) { if (!byT.has(k.t)) byT.set(k.t, []); byT.get(k.t).push(k); }
    const rows = [...byT.keys()].sort(teamOrder).map((t) => { const ks = byT.get(t), c = line(D.teams[t], ...parts(ks)); c.splice(1, 0, cN(ks.filter((k) => k.p >= 0).length)); return { c }; });
    if (byT.size > 1) { const ks = [...all.values()], c = line('All teams', ...parts(ks)); c.splice(1, 0, cN(ks.filter((k) => k.p >= 0).length)); rows.push({ cls: 'total', c }); }
    return { cols, rows };
  }
  const rows = grouped([...all.values()], (k) => ({ c: line(pName(k.p), ...parts([k]), k.p < 0 ? 'muted' : '') }),
    (t, ks) => ({ c: line(`${D.teams[t]} total`, ...parts(ks)) }), (ks, multi) => ({ c: line(multi ? 'All teams' : 'Total', ...parts(ks)) }));
  return { cols, rows: rows.map((r) => (r.group ? groupRow(r.group, cols.length) : r)) };
}
const tip = $('#tip');
function showTip(html, ev) {
  tip.innerHTML = html; tip.hidden = false;
  const pad = 14, w = tip.offsetWidth, h = tip.offsetHeight;
  let x = ev.clientX + pad, y = ev.clientY + pad;
  if (x + w > innerWidth - 8) x = ev.clientX - w - pad;
  if (y + h > innerHeight - 8) y = ev.clientY - h - pad;
  tip.style.left = Math.max(8, x) + 'px'; tip.style.top = Math.max(8, y) + 'px';
}
function hideTip() { tip.hidden = true; }
const tipRow = (k, v, color) => `<div class="r"><span>${color ? `<i class="k" style="background:${color}"></i>` : ''}${k}</span><span>${v}</span></div>`;
const NS = 'http://www.w3.org/2000/svg';
function svgEl(tag, attrs, parent) { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; }
function niceMax(v) { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))); for (const m of [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) if (m * p >= v) return m * p; return 10 * p; }
function dailyChart(el) {
  el.innerHTML = '';
  const days = V.days, all = mergeMaps([...V.pd.values()].map((k) => k.days));
  if (!days.length) { el.innerHTML = emptyMsg('No emails or closed cases for these filters.'); return; }
  const c = { muted: css('--muted'), grid: css('--o-grid'), axis: css('--o-axis'), s1: css('--s1'), s3: css('--s3') };
  const recv = days.map((d) => (all.get(d) || [0, 0, 0])[2]), work = days.map((d) => { const v = all.get(d) || [0, 0]; return v[0] + v[1]; });
  const W = el.clientWidth || 600, H = 260, m = { l: 46, r: 16, t: 14, b: 28 }, iw = W - m.l - m.r, ih = H - m.t - m.b, n = days.length, bw = iw / n;
  const max = niceMax(Math.max(...recv, ...work));
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'Emails received and worked per day' }, el);
  const Xp = (i) => m.l + (i + 0.5) * bw, Y = (v) => m.t + ih * (1 - v / max);
  for (let t = 0; t <= 4; t++) {
    const v = (max / 4) * t;
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: t === 0 ? c.axis : c.grid, 'stroke-width': 1 }, svg);
    svgEl('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = fmtN(v);
  }
  const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(iw / 44))));
  days.forEach((d, i) => { if (i % every === 0) svgEl('text', { x: Xp(i), y: H - 8, 'text-anchor': 'middle', 'font-size': 12, fill: c.muted }, svg).textContent = md(d); });
  for (const [vals, color] of [[recv, c.s1], [work, c.s3]]) {
    svgEl('path', { d: vals.map((v, i) => (i ? 'L' : 'M') + Xp(i).toFixed(1) + ',' + Y(v).toFixed(1)).join(''), fill: 'none', stroke: color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }, svg);
    if (n <= 40) vals.forEach((v, i) => svgEl('circle', { cx: Xp(i), cy: Y(v), r: 3, fill: color }, svg));
  }
  const guide = svgEl('line', { y1: m.t, y2: m.t + ih, stroke: c.axis, 'stroke-width': 1, visibility: 'hidden' }, svg);
  const hit = svgEl('rect', { x: m.l, y: m.t, width: iw, height: ih, fill: 'transparent' }, svg);
  hit.addEventListener('mousemove', (ev) => {
    const r = svg.getBoundingClientRect(), i = Math.max(0, Math.min(n - 1, Math.floor(((ev.clientX - r.left) * (W / r.width) - m.l) / bw)));
    const v = all.get(days[i]) || [0, 0, 0, 0];
    guide.setAttribute('x1', Xp(i)); guide.setAttribute('x2', Xp(i)); guide.setAttribute('visibility', 'visible');
    showTip(`<div class="tt">${DOW[dowOf(days[i])]} ${dLong(days[i])}</div>` + tipRow('Emails received', fmtN(v[2]), c.s1) + tipRow('Actioned + closed', fmtN(v[0] + v[1]), c.s3) +
      tipRow('Emails actioned', fmtN(v[0])) + tipRow('Cases closed', fmtN(v[1])) + tipRow('% actioned', fmtP(pctOf(v))), ev);
  });
  hit.addEventListener('mouseleave', () => { guide.setAttribute('visibility', 'hidden'); hideTip(); });
}

// ---------- render ----------
function renderHeader() {
  $('#eyebrow').textContent = L.fromBook ? `Actioned & closed case count · ${L.sections.map((s) => s.name).join(' / ')}` : 'Actioned & closed case count · by team';
  const parts = [];
  if (isFinite(X.lo)) parts.push(`Emails and closed cases <b>${span(X.lo, X.hi)}</b>`);
  if (I.weeks.length) parts.push(`SLA run dates <b>${md(I.weeks[0] + 7)}${I.weeks.length > 1 ? '–' + md(I.weeks[I.weeks.length - 1] + 7) : ''}</b>`);
  $('#range').innerHTML = parts.join(' · ');
}
function render() {
  if (!D) return;
  renderHeader(); hideTip();
  const t = S.tab;
  if (t[0] === 'm') renderMonth(+t.slice(1));
  else if (t[0] === 'c') renderCust(+t.slice(1));
  else if (t === 'wsla') renderSla();
  else if (t === 'open') renderOpen();
  else if (t === 'names') renderNames();
  else if (t === 'over') { aggregate(); renderTiles(); dailyChart($('#cDaily')); renderTable($('#tTeam'), summaryModel(true)); renderTable($('#tPerson'), summaryModel(false)); }
}
let rz = 0;
addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { if (D && S.tab === 'over') dailyChart($('#cDaily')); }, 150); });

// ---------- copy + export ----------
function tableText(table) {
  return [...table.rows].map((tr) => [...tr.cells].map((td) => {
    const c = td.cloneNode(true);
    c.querySelectorAll('small').forEach((e) => e.replaceWith(' ' + e.textContent));
    return c.textContent.replace(/\s+/g, ' ').trim();
  }).join('\t')).join('\n');
}
document.addEventListener('click', async (e) => {
  const b = e.target.closest('.copybtn'); if (!b) return;
  const table = document.querySelector(b.dataset.copy + ' table');
  if (!table) return;
  const text = tableText(table), label = b.dataset.label || (b.dataset.label = b.textContent);
  try { await navigator.clipboard.writeText(text); b.textContent = 'Copied'; }
  catch (err) {
    const ta = document.createElement('textarea'); ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.append(ta); ta.select();
    b.textContent = document.execCommand('copy') ? 'Copied' : 'Copy failed'; ta.remove();
  }
  setTimeout(() => { b.textContent = label; }, 1500);
});

let xlsxLoading = null;
function loadXlsx() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (!xlsxLoading) {
    xlsxLoading = new Promise((ok, fail) => {
      const s = document.createElement('script');
      s.src = '/assets/vendor/xlsx.full.min.js';
      s.onload = () => ok(window.XLSX);
      s.onerror = () => { xlsxLoading = null; fail(new Error('The Excel library could not be loaded. Check your connection and try again.')); };
      document.head.append(s);
    });
  }
  return xlsxLoading;
}
function xCell(c) {
  if (c.f === 's') return c.v ?? '';
  if (c.v == null || (c.f === 'n' && !c.v)) return '';
  return c.f === 'n' || c.f === 'z' ? { t: 'n', v: c.v, z: '#,##0' } : c.f === 'p' ? { t: 'n', v: c.v, z: '0%' } : { t: 'n', v: c.v / 1440, z: '[h]:mm' };
}
const xDate = (d) => ({ t: 'n', v: d + 25569, z: 'm/d' });
const dowRow = (m) => ['', ...m.cols.slice(1).map((c) => (c.d != null ? DOW[dowOf(c.d)].toUpperCase() : ''))];
const headRow = (m, first) => [first, ...m.cols.slice(1).map((c) => (c.d != null ? xDate(c.d) : c.x))];
const bodyRows = (m) => m.rows.map((r) => r.c.map(xCell));
// The export copies the workbook's layout, so blocks can be pasted straight into it.
async function exportExcel() {
  const btn = $('#exportBtn'), label = btn.textContent;
  btn.disabled = true; btn.textContent = 'Preparing…';
  try {
    const XLSX = await loadXlsx(), wb = XLSX.utils.book_new();
    const add = (name, aoa, widths) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = widths; XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31)); };
    for (const k of I.months) {
      const aoa = [];
      for (const s of L.sections) {
        aoa.push([s.name]);
        for (const i of [0, 1]) { const m = monthModel(k, s, i); aoa.push(dowRow(m), headRow(m, i ? 'CASES CLOSED' : 'CASES ACTIONED'), ...bodyRows(m), []); }
        aoa.push([]);
      }
      add(sheetName(k), aoa, [{ wch: 18 }]);
      const m = custModel(k), cAoa = [headRow(m, 'CUSTOMER / REP')];
      for (const r of m.rows) cAoa.push(r.c.map(xCell));
      add(`CUST E-MAILS ${SHEET_MON[k % 12]}`, cAoa, [{ wch: 30 }]);
    }
    if (D.sla.length) {
      const lists = L.sla.filter((l) => l.rows.length), aoa = [lists.flatMap((l) => [l.title, '', ''])];
      I.weeks.forEach((w, wi) => {
        const ms = lists.map(slaModel), height = Math.max(...ms.map((m) => m.rows.length));
        aoa.push(lists.flatMap(() => [`RUN DATE ${md(w + 7)}`, '', '']), lists.flatMap(() => ['CSR', 'SLA 1ST RESPONSE', '']));
        for (let r = 0; r < height; r++) aoa.push(ms.flatMap((m) => { const row = m.rows[r]; return row ? [row.c[0].v, xCell(row.c[wi + 1]), ''] : ['', '', '']; }));
        aoa.push([]);
      });
      add('WEEKLY SLA', aoa, lists.flatMap(() => [{ wch: 16 }, { wch: 18 }, { wch: 3 }]));
    }
    const om = openModel(); add('Open cases', [om.cols.map((c) => c.x), ...bodyRows(om)], [{ wch: 24 }]);
    aggregate(); const sm = summaryModel(false); add('Summary by person', [sm.cols.map((c) => c.x), ...bodyRows(sm)], [{ wch: 28 }]);
    const d = new Date(), pad = (n) => String(n).padStart(2, '0');
    XLSX.writeFile(wb, `ACTIONED & CLOSED CASE COUNT ${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}.xlsx`);
  } catch (err) {
    $('#notice').hidden = false; $('#notice').textContent = err.message;
  } finally { btn.disabled = false; btn.textContent = label; }
}
$('#exportBtn').addEventListener('click', () => { if (D) exportExcel(); });
