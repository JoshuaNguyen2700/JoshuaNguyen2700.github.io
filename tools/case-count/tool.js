// Logic for the case-count tool only. It is an ES module, so nothing here leaks into other pages.
// The page ships with no case data. The visitor picks the Salesforce report exports (or the folder that
// holds them); worker.js reads them in the browser and returns small count tables (nothing is uploaded).
// The reports rebuild the "ACTIONED & CLOSED CASE COUNT" workbook: the monthly CASES ACTIONED / CASES
// CLOSED tabs, the CUST E-MAILS tabs (% actioned, goal 85%), WEEKLY SLA and Open cases.
import { esc } from '/assets/core/util.js';

const $ = (s) => document.querySelector(s);
const STATE_KEY = 'cc.state';   // tab and view choices only; never data
const GOAL = 0.85;

// ---------- dates (whole days since 1970-01-01, as the worker sends them) ----------
const DAY = 864e5;
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
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
const today = () => { const n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / DAY; };

// ---------- formatting ----------
const fmtN = (n) => Math.round(n).toLocaleString('en-US');
const fmtP = (p, d = 0) => (p == null || !isFinite(p) ? '–' : (p * 100).toFixed(d) + '%');
const fmtHM = (mins) => { if (mins == null || !isFinite(mins)) return '–'; const t = Math.round(mins); return Math.floor(t / 60) + ':' + String(t % 60).padStart(2, '0'); };
const plural = (n, w) => `${fmtN(n)} ${w}${n === 1 ? '' : 's'}`;
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

// ---------- data + view state ----------
// D (from worker.js): teams[], people[], companies[], statuses[], files[], and count tables
//   act  [team, person, day, n]            emails actioned (sent)        person -1 = no known case owner
//   clo  [team, person, day, n]            cases closed
//   inb  [team, person, company, day, ceva, n]   emails received (ceva 1 = from a CEVA address)
//   sla  [team, person, monday, cases, elapsedMins, slaKnown, slaMet]
//   open [team, person, status, n]
let D = null, S = null, V = null, worker = null;
const saved = (() => { try { return JSON.parse(localStorage.getItem(STATE_KEY)) || {}; } catch (e) { return {}; } })();
const saveView = () => { try { localStorage.setItem(STATE_KEY, JSON.stringify({ tab: S.tab, pct: S.pct, sla: S.sla })); } catch (e) {} };
const pName = (p) => (p < 0 ? '(no case owner)' : D.people[p]);
const byText = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base', numeric: true });
const teamOrder = (a, b) => byText(D.teams[a], D.teams[b]);
const personOrder = (a, b) => (a.p < 0) - (b.p < 0) || byText(pName(a.p), pName(b.p));

function extent() {
  let lo = Infinity, hi = -Infinity, slaLo = Infinity, slaHi = -Infinity;
  for (const r of D.act) { lo = Math.min(lo, r[2]); hi = Math.max(hi, r[2]); }
  for (const r of D.clo) { lo = Math.min(lo, r[2]); hi = Math.max(hi, r[2]); }
  for (const r of D.inb) { lo = Math.min(lo, r[3]); hi = Math.max(hi, r[3]); }
  for (const r of D.sla) { slaLo = Math.min(slaLo, r[2]); slaHi = Math.max(slaHi, r[2] + 6); }
  const from = Math.min(lo, slaLo), to = isFinite(hi) ? hi : slaHi;
  return { lo, hi, slaLo, slaHi, from, to };
}
let X = null;   // data extent
const defaults = () => ({ teams: D.teams.map((_, i) => i), person: -1, from: X.from, to: X.to,
  tab: ['over', 'ac', 'cust', 'sla', 'about'].includes(saved.tab) ? saved.tab : 'over',
  pct: ['pct', 'recv', 'worked'].includes(saved.pct) ? saved.pct : 'pct', sla: ['avg', 'met', 'n'].includes(saved.sla) ? saved.sla : 'avg' });

function load(data) {
  D = data; X = extent(); S = defaults();
  $('#notice').hidden = true;
  const skipped = D.files.filter((f) => !f.kind);
  const usable = D.files.filter((f) => f.kind);
  $('#meta').textContent = `${plural(usable.length, 'report')} · ${plural(D.people.length, 'person')}`.replace('persons', 'people');
  const miss = ['sent', 'received', 'closed', 'sla'].filter((k) => !usable.some((f) => f.kind === k));
  const KN = { sent: 'sent emails (actioned)', received: 'received emails', closed: 'closed cases', sla: 'SLA' };
  const notes = [];
  if (miss.length) notes.push(`No ${miss.map((k) => KN[k]).join(', ')} report was found, so those figures are blank.`);
  if (skipped.length) notes.push(`${plural(skipped.length, 'file')} skipped (not a case or email report): ${skipped.map((f) => f.name.split('/').pop()).join(', ')}.`);
  if (notes.length) { $('#notice').hidden = false; $('#notice').textContent = notes.join(' '); }
  buildFilters(); show('app'); setTab(S.tab);
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
$('#closeBtn').addEventListener('click', () => { D = null; V = null; show('start'); });

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

// ---------- filters ----------
function buildFilters() {
  $('#teamChips').innerHTML = D.teams.map((t, i) => [t, i]).sort((a, b) => byText(a[0], b[0]))
    .map(([t, i]) => `<button class="chip" type="button" data-t="${i}" aria-pressed="${S.teams.includes(i)}">${esc(t)}</button>`).join('');
  fillPeople();
  for (const id of ['#fFrom', '#fTo']) { $(id).min = toInput(X.from); $(id).max = toInput(X.to); }
  $('#fFrom').value = toInput(S.from); $('#fTo').value = toInput(S.to);
}
// People with any activity, grouped by team; only the selected teams are listed.
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
$('#resetBtn').addEventListener('click', () => { const { tab, pct, sla } = S; S = { ...defaults(), tab, pct, sla }; buildFilters(); render(); });

const TABS = ['over', 'ac', 'cust', 'sla', 'about'];
function setTab(t) {
  S.tab = TABS.includes(t) ? t : 'over';
  document.querySelectorAll('.cc-tabs .tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === S.tab)));
  document.querySelectorAll('.panel').forEach((p) => (p.hidden = p.id !== 'p-' + S.tab));
  saveView(); render();
}
$('.cc-tabs').addEventListener('click', (e) => { const b = e.target.closest('.tab'); if (b) setTab(b.dataset.tab); });

// ---------- aggregation for the current filters ----------
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
  const cust = new Map();   // company -> Map(day -> n), customers only (CEVA addresses are their own row)
  for (const r of D.inb) {
    if (!ok(r) || !inR(r[3])) continue;
    const v = slot(r[0], r[1], r[3]); v[2] += r[5]; if (r[4]) v[3] += r[5];
    if (!r[4]) { let m = cust.get(r[2]); if (!m) cust.set(r[2], (m = new Map())); m.set(r[3], (m.get(r[3]) || 0) + r[5]); }
  }
  const sla = new Map(), weeks = new Set();   // team,person -> Map(monday -> [cases, mins, known, met])
  for (const r of D.sla) {
    if (!ok(r) || r[2] + 6 < S.from || r[2] > S.to) continue;
    const k = r[0] + ',' + r[1]; let o = sla.get(k);
    if (!o) sla.set(k, (o = { t: r[0], p: r[1], wk: new Map() }));
    const v = o.wk.get(r[2]) || [0, 0, 0, 0]; for (let i = 0; i < 4; i++) v[i] += r[3 + i]; o.wk.set(r[2], v);
    weeks.add(r[2]);
  }
  const open = new Map();   // team,person -> Map(status -> n); a snapshot, so the dates don't apply
  for (const r of D.open) {
    if (!ok(r)) continue;
    const k = r[0] + ',' + r[1]; let o = open.get(k);
    if (!o) open.set(k, (o = { t: r[0], p: r[1], st: new Map() }));
    o.st.set(r[2], (o.st.get(r[2]) || 0) + r[3]);
  }
  V = { pd, days: [...days].sort((a, b) => a - b), cust, sla, weeks: [...weeks].sort((a, b) => a - b), open };
}
const addVec = (a, b) => { if (!b) return a; if (!a) return b.slice(); for (let i = 0; i < b.length; i++) a[i] += b[i]; return a; };
const sumDays = (m, ds) => { let v = null; for (const d of ds) v = addVec(v, m.get(d)); return v; };
const mergeMaps = (maps) => { const out = new Map(); for (const m of maps) for (const [k, v] of m) out.set(k, addVec(out.get(k), v)); return out; };
const pctOf = (v) => (v && v[2] ? (v[0] + v[1]) / v[2] : null);

// ---------- tables: one model renders to HTML, copies as text and exports to Excel ----------
// model: { cls, cols: [{ h (html), x (plain), cls }], rows: [{ cls, c: [cell] }] }; cell: { v, f, cls }
// f: 's' text, 'n' count (blank when 0), 'p' percent, 'hm' minutes shown as h:mm
const cS = (v, cls) => ({ v, f: 's', cls }), cN = (v, cls) => ({ v, f: 'n', cls }), cP = (v, cls) => ({ v, f: 'p', cls }), cHM = (v, cls) => ({ v, f: 'hm', cls });
const goalCls = (p) => (p == null ? '' : p < GOAL ? 'low' : 'ok');
function cellHtml(c) {
  if (c.f === 's') return esc(c.v ?? '');
  if (c.v == null || (c.f === 'n' && !c.v)) return '';
  return c.f === 'n' ? fmtN(c.v) : c.f === 'p' ? fmtP(c.v) : fmtHM(c.v);
}
const emptyMsg = (t) => `<div class="empty-s">${t}</div>`;
function renderTable(el, m, scrollEnd) {
  if (!m.rows.length) { el.innerHTML = emptyMsg(m.empty || 'Nothing for these filters.'); return; }
  const th = m.cols.map((c) => `<th class="${c.cls || ''}">${c.h}</th>`).join('');
  const body = m.rows.map((r) => `<tr class="${r.cls || ''}">${r.c.map((c, i) => `<td class="${[m.cols[i] && m.cols[i].cls, c.cls].filter(Boolean).join(' ')}">${cellHtml(c)}</td>`).join('')}</tr>`).join('');
  el.innerHTML = `<table class="${m.cls || ''}"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
  if (scrollEnd) el.scrollLeft = el.scrollWidth;
}
const col = (h, cls, x) => ({ h, cls, x: x ?? h.replace(/<small>/g, ' ').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&') });

// Rows grouped by team (a heading row, the people A to Z, a team total), then a grand total.
// With one team the headings and team totals are left out.
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

// Day columns with a weekly total after each week, like the monthly tabs of the workbook.
function dayCols(days) {
  const out = []; let wk = null, cur = [];
  const flush = () => { if (cur.length) out.push({ wk, days: cur }); };
  for (const d of days) { const m = monday(d); if (m !== wk) { flush(); wk = m; cur = []; } cur.push(d); out.push({ d }); }
  flush();
  return out;
}
// rows: [{ label, vals: Map(day -> vec), cellOf?, cls? } | { group }]
function dayMatrix(rows, days, cellOf, first = 'Person') {
  const dc = dayCols(days), now = today();
  const cols = [col(first), col('Total', 'tot'), ...dc.map((c) => (c.d != null
    ? col(`${DOW[dowOf(c.d)]}<small>${md(c.d)}</small>`, c.d === now ? 'now' : '', `${DOW[dowOf(c.d)]} ${md(c.d)}`)
    : col(`W${isoWeek(c.wk)}<small>total</small>`, 'wk', `W${isoWeek(c.wk)} total`)))];
  return { cls: 'mx', cols, rows: rows.map((r) => {
    if (r.group) return groupRow(r.group, cols.length);
    const f = r.cellOf || cellOf;
    return { cls: r.cls, c: [cS(r.label, r.lcls), f(sumDays(r.vals, days)), ...dc.map((c) => f(c.d != null ? r.vals.get(c.d) : sumDays(r.vals, c.days)))] };
  }) };
}
function personDayRows(keep = () => true) {
  const keys = [...V.pd.values()].filter(keep);
  return grouped(keys, (k) => ({ label: pName(k.p), vals: k.days, lcls: k.p < 0 ? 'muted' : '' }),
    (t, ks) => ({ label: `${D.teams[t]} total`, vals: mergeMaps(ks.map((k) => k.days)) }),
    (ks, multi) => ({ label: multi ? 'All teams' : 'Total', vals: mergeMaps(ks.map((k) => k.days)) }));
}

// ---------- overview ----------
function totals(keysPd, keysSla, keysOpen) {
  const t = { act: 0, clo: 0, recv: 0, ceva: 0, cases: 0, mins: 0, known: 0, met: 0, open: 0 };
  for (const k of keysPd) for (const v of k.days.values()) { t.act += v[0]; t.clo += v[1]; t.recv += v[2]; t.ceva += v[3]; }
  for (const k of keysSla) for (const v of k.wk.values()) { t.cases += v[0]; t.mins += v[1]; t.known += v[2]; t.met += v[3]; }
  for (const k of keysOpen) for (const n of k.st.values()) t.open += n;
  return t;
}
function tile(label, value, sub, cls = '') { return `<div class="ktile ${cls}"><div class="eyebrow">${label}</div><div class="v">${value}</div><div class="d">${sub}</div></div>`; }
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
  // every team+person seen in any table
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
// ---------- daily chart ----------
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
  const c = { muted: css('--muted'), grid: css('--o-grid'), axis: css('--o-axis'), card: css('--o-card'), ink: css('--text'), s1: css('--s1'), s3: css('--s3') };
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

// ---------- actioned & closed ----------
// Only people with emails sent (or cases closed) are listed; days are those with any of either, so the
// two tables line up.
function acModel(i) {
  const keys = [...V.pd.values()], has = (k, d) => { const v = k.days.get(d); return v && (v[0] || v[1]); };
  return dayMatrix(personDayRows((k) => [...k.days.values()].some((v) => v[i])),
    V.days.filter((d) => keys.some((k) => has(k, d))), (v) => cN(v ? v[i] : 0));
}

// ---------- customer emails ----------
const PCT = { pct: '% actioned', recv: 'Emails received', worked: 'Actioned + closed' };
function pctModel() {
  const f = S.pct === 'recv' ? (v) => cN(v ? v[2] : 0) : S.pct === 'worked' ? (v) => cN(v ? v[0] + v[1] : 0) : (v) => { const p = pctOf(v); return cP(p, goalCls(p)); };
  return dayMatrix(personDayRows(), V.days, f);
}
// The CUST E-MAILS block: one row per customer, CEVA station emails, total, emails actioned, % actioned.
function custModel(search = '') {
  const q = search.trim().toLowerCase();
  const all = mergeMaps([...V.pd.values()].map((k) => k.days));
  const custRows = [...V.cust.entries()].map(([c, m]) => ({ label: D.companies[c], vals: new Map([...m].map(([d, n]) => [d, [n]])) }))
    .filter((r) => !q || r.label.toLowerCase().includes(q)).sort((a, b) => byText(a.label, b.label));
  const pick = (i) => new Map([...all].map(([d, v]) => [d, [v[i]]]));
  const ceva = pick(3), total = pick(2), workedM = new Map([...all].map(([d, v]) => [d, [v[0] + v[1], v[2]]]));
  const n1 = (v) => cN(v ? v[0] : 0);
  const rows = [...custRows,
    { label: 'E-mails from CEVA stations', vals: ceva, cls: 'key' },
    { label: 'Total received', vals: total, cls: 'sub' },
    { label: 'E-mails actioned (actioned + closed)', vals: workedM, cls: 'key' },
    { label: '% actioned (goal 85%)', vals: workedM, cls: 'key', cellOf: (v) => { const p = v && v[1] ? v[0] / v[1] : null; return cP(p, goalCls(p)); } }];
  return dayMatrix(rows, V.days, n1, 'Customer');
}

// ---------- SLA ----------
const SLA = { avg: 'Avg first response', met: 'Within SLA', n: 'Cases' };
function slaCell(v) {
  if (!v || !v[0]) return S.sla === 'n' ? cN(0) : cS('');
  return S.sla === 'n' ? cN(v[0]) : S.sla === 'met' ? cP(v[2] ? v[3] / v[2] : null) : cHM(v[1] / v[0]);
}
// Team rows average the people's figures, like TEAM AVERAGE on the WEEKLY SLA tab; case counts add up.
function slaGroupCell(vs) {
  if (S.sla === 'n') return cN(vs.reduce((a, v) => a + (v ? v[0] : 0), 0));
  const xs = vs.filter((v) => v && v[0]).map((v) => (S.sla === 'met' ? (v[2] ? v[3] / v[2] : null) : v[1] / v[0])).filter((x) => x != null);
  if (!xs.length) return cS('');
  const avg = xs.reduce((a, b) => a + b, 0) / xs.length;
  return S.sla === 'met' ? cP(avg) : cHM(avg);
}
function slaModel() {
  const weeks = V.weeks;
  const cols = [col('Person'), col('All weeks', 'tot'), ...weeks.map((w) => col(`W${isoWeek(w)}<small>${md(w)}–${md(w + 6)}</small>`, '', `W${isoWeek(w)} (${md(w)}–${md(w + 6)})`))];
  const keys = [...V.sla.values()];
  const all = (k) => { let v = null; for (const x of k.wk.values()) v = addVec(v, x); return v; };
  const rows = grouped(keys,
    (k) => ({ c: [cS(pName(k.p), k.p < 0 ? 'muted' : ''), slaCell(all(k)), ...weeks.map((w) => slaCell(k.wk.get(w)))] }),
    (t, ks) => ({ c: [cS(S.sla === 'n' ? `${D.teams[t]} total` : `${D.teams[t]} average`), slaGroupCell(ks.map(all)), ...weeks.map((w) => slaGroupCell(ks.map((k) => k.wk.get(w))))] }),
    (ks, multi) => ({ c: [cS(S.sla === 'n' ? (multi ? 'All teams' : 'Total') : (multi ? 'All teams average' : 'Team average')), slaGroupCell(ks.map(all)), ...weeks.map((w) => slaGroupCell(ks.map((k) => k.wk.get(w))))] }));
  return { cls: 'mx', cols, rows: rows.map((r) => (r.group ? groupRow(r.group, cols.length) : r)), empty: D.sla.length ? 'No SLA cases for these filters.' : 'Load the SLA report (the export with Elapsed Time and SLA Breached?) to see first response.' };
}
const STATUS_ORDER = ['new', 're-opened', 'answer received', 'in progress', 'on hold', 'escalated'];
function openModel() {
  const used = new Set(); for (const o of V.open.values()) for (const s of o.st.keys()) used.add(s);
  const sts = [...used].sort((a, b) => { const x = STATUS_ORDER.indexOf(D.statuses[a].toLowerCase()), y = STATUS_ORDER.indexOf(D.statuses[b].toLowerCase()); return (x < 0 ? 99 : x) - (y < 0 ? 99 : y) || byText(D.statuses[a], D.statuses[b]); });
  const cols = [col('Person'), ...sts.map((s) => col(esc(D.statuses[s]))), col('Total', 'wk')];
  const line = (label, objs, lcls) => { const n = sts.map((s) => objs.reduce((a, o) => a + (o.st.get(s) || 0), 0)); return [cS(label, lcls), ...n.map((x) => cN(x)), cN(n.reduce((a, b) => a + b, 0))]; };
  const rows = grouped([...V.open.values()], (k) => ({ c: line(pName(k.p), [k], k.p < 0 ? 'muted' : '') }),
    (t, ks) => ({ c: line(`${D.teams[t]} total`, ks) }), (ks, multi) => ({ c: line(multi ? 'All teams' : 'Total', ks) }));
  return { cols, rows: rows.map((r) => (r.group ? groupRow(r.group, cols.length) : r)), empty: 'No open cases for these filters.' };
}

// ---------- how it's calculated ----------
const KIND = { sent: 'Emails sent (actioned)', received: 'Emails received', closed: 'Cases closed', sla: 'SLA' };
function filesModel() {
  return { cols: [col('File'), col('Team'), col('Report'), col('Rows'), col('Dates')],
    rows: D.files.map((f) => ({ cls: f.kind ? '' : 'muted', c: [cS(f.name), cS(f.team || '–'), cS(f.kind ? KIND[f.kind] : f.note || 'Skipped'), cN(f.rows), cS(f.from != null ? span(f.from, f.to) + (f.kind === 'sla' ? ' (opened)' : '') : '')] })) };
}
function about() {
  $('#about').innerHTML = `<p>Rebuilds the <b>ACTIONED &amp; CLOSED CASE COUNT</b> workbook from the Salesforce report exports. Each report type is recognised from its columns, so the file names don't matter.</p>
    <ul>
      <li><b>Cases actioned</b>: emails sent on the cases a person owns, counted on the email's date (every row of the sent-emails report).</li>
      <li><b>Cases closed</b>: cases whose Date/Time Closed falls on that day, credited to the case owner.</li>
      <li><b>Emails received</b>: every email in the received-emails report. Emails from an @cevalogistics.com address are <b>E-mails from CEVA stations</b>; the rest are customer emails, listed by the case's company. Each email counts for the owner of its case, matched by case number from the other reports.</li>
      <li><b>% actioned</b> = (emails actioned + cases closed) ÷ emails received, against the <b>85% goal</b>, as on the CUST E-MAILS tabs.</li>
      <li><b>SLA first response</b>: the average Elapsed Time of the cases opened that week (hours:minutes). <b>Within SLA</b> is the share marked "SLA Met". Team rows average the people's figures, like TEAM AVERAGE on the WEEKLY SLA tab.</li>
      <li><b>Open cases</b>: cases in the SLA export whose status isn't Closed (New, Re-Opened, Answer Received, In Progress, …).</li>
      <li><b>Teams</b> are the folders the reports sit in (US East, US West, Legacy, …). A person belongs to the team where most of their activity is.</li>
      <li>Exports that overlap (for example two weeks with shared days) count each email and case once.</li>
      <li>PTO isn't in Salesforce, so a day off shows as a blank cell rather than PTO.</li>
    </ul>`;
  renderTable($('#tFiles'), filesModel());
}

// ---------- render ----------
function choice(sel, key, opts) {
  const el = $(sel);
  el.innerHTML = Object.entries(opts).map(([k, l]) => `<button type="button" data-v="${k}" aria-pressed="${S[key] === k}">${l}</button>`).join('');
  el.onclick = (e) => { const b = e.target.closest('button'); if (b && S[key] !== b.dataset.v) { S[key] = b.dataset.v; saveView(); render(); } };
}
function renderHeader() {
  const sel = S.teams.length === D.teams.length ? 'all teams' : S.teams.map((t) => D.teams[t]).sort(byText).join(', ');
  $('#eyebrow').textContent = `Cases & emails · ${S.person >= 0 ? D.people[S.person] : sel}`;
  const parts = [];
  if (V.days.length) parts.push(`Emails and closed cases <b>${span(V.days[0], V.days[V.days.length - 1])}</b>`);
  if (V.weeks.length) parts.push(`SLA weeks <b>W${isoWeek(V.weeks[0])}${V.weeks.length > 1 ? '–W' + isoWeek(V.weeks[V.weeks.length - 1]) : ''}</b>`);
  $('#range').innerHTML = parts.join(' · ') || 'No activity in the selected dates.';
}
function render() {
  if (!D) return;
  aggregate(); renderHeader(); hideTip();
  const t = S.tab;
  if (t === 'over') { renderTiles(); dailyChart($('#cDaily')); renderTable($('#tTeam'), summaryModel(true)); renderTable($('#tPerson'), summaryModel(false)); }
  if (t === 'ac') { renderTable($('#mAct'), acModel(0), true); renderTable($('#mClo'), acModel(1), true); }
  if (t === 'cust') {
    choice('#pctMetric', 'pct', PCT);
    $('#pctTitle').textContent = `${PCT[S.pct]} by person and day`;
    renderTable($('#mPct'), pctModel(), true);
    $('#custTitle').textContent = S.person >= 0 ? `${D.people[S.person]}: emails received by customer` : 'Emails received by customer';
    $('#custSub').textContent = `${plural(V.cust.size, 'customer')} · ${S.person >= 0 ? 'cases this person owns' : 'all cases in the selected teams'}`;
    renderTable($('#mCust'), custModel($('#custSearch').value), true);
  }
  if (t === 'sla') {
    choice('#slaMetric', 'sla', SLA);
    $('#slaTitle').textContent = `${S.sla === 'avg' ? 'SLA first response' : S.sla === 'met' ? 'Cases within SLA' : 'Cases opened'} by person and week`;
    $('#slaNote').textContent = S.sla === 'avg' ? 'Average hours:minutes from the case opening to the first response.' : S.sla === 'met' ? 'Share of cases marked "SLA Met".' : 'Cases opened that week, by owner.';
    renderTable($('#mSla'), slaModel(), true);
    renderTable($('#tOpen'), openModel());
  }
  if (t === 'about') about();
}
$('#custSearch').addEventListener('input', () => { if (D && V) renderTable($('#mCust'), custModel($('#custSearch').value), true); });
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
  return c.f === 'n' ? { t: 'n', v: c.v, z: '#,##0' } : c.f === 'p' ? { t: 'n', v: c.v, z: '0%' } : { t: 'n', v: c.v / 1440, z: '[h]:mm' };
}
function filterText() {
  const sel = S.teams.length === D.teams.length ? 'All teams' : S.teams.map((t) => D.teams[t]).sort(byText).join(', ');
  return `${sel}${S.person >= 0 ? ' · ' + D.people[S.person] : ''} · ${span(S.from, S.to)}`;
}
function sheet(XLSX, title, m) {
  const aoa = [[title], [filterText()], [], m.cols.map((c) => c.x), ...m.rows.map((r) => r.c.map(xCell))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = m.cols.map((c, i) => ({ wch: i === 0 ? 34 : Math.max(9, Math.min(18, (c.x || '').length + 2)) }));
  return ws;
}
async function exportExcel() {
  const btn = $('#exportBtn'), label = btn.textContent;
  btn.disabled = true; btn.textContent = 'Preparing…';
  try {
    const XLSX = await loadXlsx(), wb = XLSX.utils.book_new(), keep = S.pct;
    const add = (name, title, m) => XLSX.utils.book_append_sheet(wb, sheet(XLSX, title, m), name);
    add('By person', 'Summary by person', summaryModel(false));
    add('By team', 'Summary by team', summaryModel(true));
    add('Cases actioned', 'CASES ACTIONED', acModel(0));
    add('Cases closed', 'CASES CLOSED', acModel(1));
    S.pct = 'pct'; add('% actioned', '% ACTIONED (GOAL 85%) by person and day', pctModel()); S.pct = keep;
    add('Cust e-mails', S.person >= 0 ? `CUST E-MAILS · ${D.people[S.person]}` : 'CUST E-MAILS', custModel());
    const keepS = S.sla;
    S.sla = 'avg'; add('Weekly SLA', 'SLA 1ST RESPONSE (h:mm) by week opened', slaModel());
    S.sla = 'met'; add('SLA met', 'Cases within SLA by week opened', slaModel());
    S.sla = keepS;
    add('Open cases', 'Open cases by status', openModel());
    const d = new Date(), pad = (n) => String(n).padStart(2, '0');
    XLSX.writeFile(wb, `Case Count ${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}.xlsx`);
  } catch (err) {
    $('#notice').hidden = false; $('#notice').textContent = err.message;
  } finally { btn.disabled = false; btn.textContent = label; }
}
$('#exportBtn').addEventListener('click', () => { if (D) exportExcel(); });
