// Logic for the emails-received tool only. It is an ES module, so nothing here leaks into other pages.
// The page ships with no email data. The visitor picks Salesforce "Emails Received" exports; worker.js reads
// them in the browser (nothing is uploaded) and returns one compact row per email.
import { esc } from '/assets/core/util.js';

const $ = (s) => document.querySelector(s);
const STATE_KEY = 'er.state';   // tab and view choices only; never data

// ---------- dates (whole days since 1970-01-01) + formatting ----------
const DAY = 864e5;
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const dt = (d) => new Date(d * DAY);
const dowOf = (d) => dt(d).getUTCDay();
const md = (d) => `${dt(d).getUTCMonth() + 1}/${dt(d).getUTCDate()}`;
const dLong = (d, year = true) => `${MON[dt(d).getUTCMonth()]} ${dt(d).getUTCDate()}${year ? ', ' + dt(d).getUTCFullYear() : ''}`;
const span = (a, b) => (a === b ? dLong(a) : dt(a).getUTCFullYear() === dt(b).getUTCFullYear() ? `${dLong(a, false)} – ${dLong(b)}` : `${dLong(a)} – ${dLong(b)}`);
const toInput = (d) => dt(d).toISOString().slice(0, 10);
const fromInput = (s) => (/^\d{4}-\d{2}-\d{2}$/.test(s) ? Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) / DAY : null);
const isWeekday = (d) => { const w = dowOf(d); return w > 0 && w < 6; };
const fmtN = (n) => Math.round(n).toLocaleString('en-US');
const fmtP = (p, d = 0) => (p == null || !isFinite(p) ? '–' : (p * 100).toFixed(d) + '%');
const fmtAge = (h) => (h == null ? '' : h < 24 ? `${Math.round(h)}h` : `${Math.floor(h / 24)}d ${Math.round(h % 24)}h`);
const plural = (n, w) => `${fmtN(n)} ${w}${n === 1 ? '' : 's'}`;
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const byText = (a, b) => String(a).localeCompare(String(b), 'en', { sensitivity: 'base', numeric: true });
// A to Z, with placeholders such as "(no company)" or "<< ASSIGN CORRECT COMPANY NOW >>" last
const byName = (a, b) => { const pa = !/^[a-z0-9]/i.test(a), pb = !/^[a-z0-9]/i.test(b); return pa !== pb ? (pa ? 1 : -1) : byText(a, b); };

// ---------- data + state ----------
// D (from worker.js): teams[], companies[], froms[], subjects[], statuses[], domains[], files[], and
//   rows [team, day, company, from, subject, status, caseNumber, domain, fromCeva, caseOpenedDay, caseAgeHours, caseOpen(1/0/-1)]
let D = null, S = null, F = null, X = null, worker = null, ST = null;
const saved = (() => { try { return JSON.parse(localStorage.getItem(STATE_KEY)) || {}; } catch (e) { return {}; } })();
const saveView = () => { try { localStorage.setItem(STATE_KEY, JSON.stringify({ tab: S.tab, custSort: S.custSort, scope: S.scope })); } catch (e) {} };
const NOCOMP = /ASSIGN CORRECT COMPANY|^\(no company\)$/i;
const dflt = () => ({ teams: D.teams.map((_, i) => i), from: X.lo, to: X.hi, cust: -1, sender: 'all' });
function load(data) {
  D = data;
  let lo = Infinity, hi = -Infinity; for (const r of D.rows) { lo = Math.min(lo, r[1]); hi = Math.max(hi, r[1]); }
  X = { lo, hi };
  const sIdx = (re) => D.statuses.findIndex((s) => re.test(s));
  ST = { nw: sIdx(/^new$/i), rd: sIdx(/^read$/i), rp: sIdx(/^replied$/i) };
  S = { ...dflt(), tab: ['over', 'cust', 'unread', 'send', 'about'].includes(saved.tab) ? saved.tab : 'over',
    custSort: saved.custSort === 'most' ? 'most' : 'az', scope: saved.scope === 'open' ? 'open' : 'all', limit: 300 };
  const skipped = D.files.filter((f) => !f.rows);
  const asOf = D.files.map((f) => f.asOf).filter(Boolean).sort().pop();
  $('#meta').textContent = `${plural(D.files.length - skipped.length, 'file')} · ${fmtN(D.rows.length)} emails${asOf ? ' · as of ' + asOf : ''}`;
  $('#notice').hidden = !skipped.length;
  $('#notice').textContent = skipped.length ? `Skipped: ${skipped.map((f) => `${f.name.split('/').pop()} (${f.note})`).join(', ')}.` : '';
  buildFilters(); show('app'); setTab(S.tab);
}

// ---------- screens + loading ----------
function show(view) { for (const v of ['start', 'progress', 'app']) $('#' + v).hidden = v !== view; if (view !== 'app') hideTip(); }
function showStartError(msg) { const e = $('#startError'); e.textContent = msg; e.hidden = !msg; if (msg) show(D ? 'app' : 'start'); }
const REPORT = /\.(xls|xlsx|xlsm|csv)$/i;
// Everything picked so far: more files can be added later and the same file picked twice counts once.
let LOADED = new Map();
const itemKey = (it) => `${it.file.name}|${it.file.size}|${it.file.lastModified}`;
function addItems(items, merge) {
  items = items.filter((it) => REPORT.test(it.file.name) && !/^~\$/.test(it.file.name));
  if (!items.length) {
    const msg = 'No Excel exports there. Choose the .xlsx or .xls file exported from Salesforce.';
    if (merge && D) { $('#notice').hidden = false; $('#notice').textContent = msg; } else showStartError(msg);
    return;
  }
  if (!merge) LOADED = new Map();
  for (const it of items) LOADED.set(itemKey(it), it);
  readFiles([...LOADED.values()]);
}
function readFiles(items) {
  if (typeof Worker === 'undefined') { showStartError('This browser cannot read the files. Use a current version of Edge, Chrome, Firefox or Safari.'); return; }
  showStartError('');
  const list = $('#progressList'), rows = new Map(), done = new Set();
  list.innerHTML = '';
  for (const it of items) {
    const li = document.createElement('li');
    li.innerHTML = `<div class="pf"><b>${esc(it.path)}</b><span>Waiting</span></div><div class="bar"><i></i></div>`;
    list.append(li); rows.set(it.path, li);
  }
  const count = () => { $('#progressCount').innerHTML = `<span class="er-pcount">${done.size} of ${plural(items.length, 'file')} done.</span>`; };
  count(); show('progress');
  if (worker) worker.terminate();
  worker = new Worker(new URL('worker.js', import.meta.url));
  worker.onmessage = (e) => {
    const m = e.data;
    if (m.type === 'progress') {
      const li = rows.get(m.file); if (!li) return;
      li.querySelector('span').textContent = m.rows != null ? `${fmtN(m.rows)} emails` : m.stage;
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
$('#filesInput').addEventListener('change', (e) => { addItems(fromInputEl(e.target.files), !!D); e.target.value = ''; });
$('#folderInput').addEventListener('change', (e) => { addItems(fromInputEl(e.target.files), !!D); e.target.value = ''; });
$('#addFilesInput').addEventListener('change', (e) => { addItems(fromInputEl(e.target.files), true); e.target.value = ''; });
$('#cancelBtn').addEventListener('click', () => { if (worker) { worker.terminate(); worker = null; } show(D ? 'app' : 'start'); });
$('#closeBtn').addEventListener('click', () => { D = null; F = null; LOADED = new Map(); show('start'); });
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
  dropItems(e.dataTransfer).then((items) => addItems(items, !!D), (err) => showStartError('Those files could not be opened: ' + err.message));
});

// ---------- filters ----------
function choice(sel, key, opts, after) {
  const el = $(sel);
  el.innerHTML = Object.entries(opts).map(([k, l]) => `<button type="button" data-v="${k}" aria-pressed="${S[key] === k}">${l}</button>`).join('');
  el.onclick = (e) => { const b = e.target.closest('button'); if (b && S[key] !== b.dataset.v) { S[key] = b.dataset.v; saveView(); if (after) after(); render(); } };
}
function buildFilters() {
  $('#teamRow').hidden = D.teams.length < 2;
  $('#teamChips').innerHTML = D.teams.map((t, i) => [t, i]).sort((a, b) => byText(a[0], b[0]))
    .map(([t, i]) => `<button class="chip" type="button" data-t="${i}" aria-pressed="${S.teams.includes(i)}">${esc(t)}</button>`).join('');
  for (const id of ['#fFrom', '#fTo']) { $(id).min = toInput(X.lo); $(id).max = toInput(X.hi); }
  $('#fFrom').value = toInput(S.from); $('#fTo').value = toInput(S.to);
  fillCustomers();
  choice('#fSender', 'sender', { all: 'Everyone', cust: 'Customers', ceva: 'CEVA addresses' });
}
function fillCustomers() {
  const ts = new Set(S.teams), n = new Map();
  for (const r of D.rows) if (ts.has(r[0])) n.set(r[2], (n.get(r[2]) || 0) + 1);
  if (S.cust >= 0 && !n.has(S.cust)) S.cust = -1;
  $('#fCust').innerHTML = '<option value="-1">All customers</option>' + [...n.keys()].sort((a, b) => byName(D.companies[a], D.companies[b]))
    .map((c) => `<option value="${c}">${esc(D.companies[c])} (${fmtN(n.get(c))})</option>`).join('');
  $('#fCust').value = String(S.cust);
}
$('#teamChips').addEventListener('click', (e) => {
  const b = e.target.closest('.chip'); if (!b) return;
  const i = +b.dataset.t, on = new Set(S.teams);
  if (on.has(i)) on.delete(i); else on.add(i);
  S.teams = on.size ? [...on] : D.teams.map((_, k) => k);
  document.querySelectorAll('#teamChips .chip').forEach((c) => c.setAttribute('aria-pressed', String(S.teams.includes(+c.dataset.t))));
  fillCustomers(); render();
});
$('#fCust').addEventListener('change', (e) => { S.cust = +e.target.value; render(); });
$('#fFrom').addEventListener('change', (e) => { const d = fromInput(e.target.value); S.from = d == null ? X.lo : d; if (S.from > S.to) { S.to = S.from; $('#fTo').value = toInput(S.to); } render(); });
$('#fTo').addEventListener('change', (e) => { const d = fromInput(e.target.value); S.to = d == null ? X.hi : d; if (S.to < S.from) { S.from = S.to; $('#fFrom').value = toInput(S.from); } render(); });
$('#resetBtn').addEventListener('click', () => { Object.assign(S, dflt()); buildFilters(); render(); });
const TABS = ['over', 'cust', 'unread', 'send', 'about'];
function setTab(t) {
  S.tab = TABS.includes(t) ? t : 'over';
  document.querySelectorAll('.er-tabs .tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === S.tab)));
  document.querySelectorAll('.panel').forEach((p) => (p.hidden = p.id !== 'p-' + S.tab));
  saveView(); render();
}
$('.er-tabs').addEventListener('click', (e) => { const b = e.target.closest('.tab'); if (b) setTab(b.dataset.tab); });

// ---------- filtering + counting ----------
function filter() {
  const ts = new Set(S.teams);
  F = D.rows.filter((r) => ts.has(r[0]) && r[1] >= S.from && r[1] <= S.to && (S.cust < 0 || r[2] === S.cust)
    && (S.sender === 'all' || (S.sender === 'ceva' ? r[8] === 1 : r[8] === 0)));
}
// counts for a set of rows: emails, from CEVA, New, Read, Replied, cases, customers
function tally(rows) {
  const t = { n: 0, ceva: 0, nw: 0, rd: 0, rp: 0, other: 0, cases: new Set(), comps: new Set(), openNew: 0, noComp: 0 };
  for (const r of rows) {
    t.n++; if (r[8]) t.ceva++;
    if (r[5] === ST.nw) { t.nw++; if (r[11] === 1) t.openNew++; } else if (r[5] === ST.rd) t.rd++; else if (r[5] === ST.rp) t.rp++; else t.other++;
    t.cases.add(r[6]); t.comps.add(r[2]);
    if (NOCOMP.test(D.companies[r[2]])) t.noComp++;
  }
  return t;
}
const groupBy = (rows, k) => { const m = new Map(); for (const r of rows) { const v = r[k]; let a = m.get(v); if (!a) m.set(v, (a = [])); a.push(r); } return m; };

// ---------- tables: one model renders to HTML, copies as text and exports to Excel ----------
const cS = (v, cls) => ({ v, f: 's', cls }), cN = (v, cls) => ({ v, f: 'n', cls }), cZ = (v, cls) => ({ v, f: 'z', cls }), cP = (v, cls) => ({ v, f: 'p', cls });
function cellHtml(c) {
  if (c.f === 's') return esc(c.v ?? '');
  if (c.v == null || (c.f === 'n' && !c.v)) return '';
  return c.f === 'p' ? fmtP(c.v) : fmtN(c.v);
}
const emptyMsg = (t) => `<div class="empty-s">${t}</div>`;
function renderTable(el, m, scrollEnd) {
  if (!m.rows.length) { el.innerHTML = emptyMsg(m.empty || 'No emails for these filters.'); return; }
  const th = m.cols.map((c) => `<th class="${c.cls || ''}">${c.h}</th>`).join('');
  const body = m.rows.map((r) => `<tr class="${r.cls || ''}">${r.c.map((c, i) => `<td class="${[m.cols[i] && m.cols[i].cls, c.cls].filter(Boolean).join(' ')}">${cellHtml(c)}</td>`).join('')}</tr>`).join('');
  el.innerHTML = `<table class="${m.cls || ''}"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
  if (scrollEnd) el.scrollLeft = el.scrollWidth;
}
const col = (h, cls, x) => ({ h, cls, x: x ?? h.replace(/<small>/g, ' ').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&') });

// ---------- overview ----------
function tile(label, value, sub) { return `<div class="ktile"><div class="eyebrow">${label}</div><div class="v">${value}</div><div class="d">${sub}</div></div>`; }
function renderTiles() {
  const t = tally(F), days = new Set(F.map((r) => r[1])), wd = [...days].filter(isWeekday);
  const wdN = F.filter((r) => isWeekday(r[1])).length;
  $('#tiles').innerHTML =
    tile('Emails received', fmtN(t.n), wd.length ? `${fmtN(wdN / wd.length)} per weekday · ${plural(days.size, 'day')}` : plural(days.size, 'day')) +
    tile('From customers', fmtN(t.n - t.ceva), `${fmtP(t.n ? (t.n - t.ceva) / t.n : null)} · ${fmtN(t.ceva)} from CEVA addresses`) +
    tile('Unread (New)', fmtN(t.nw), `${fmtP(t.n ? t.nw / t.n : null)} of emails · ${fmtN(t.openNew)} on open cases`) +
    tile('Replied', fmtN(t.rp), `${fmtP(t.n ? t.rp / t.n : null)} of emails · ${fmtN(t.rd)} read only`) +
    tile('No company on the case', fmtN(t.noComp), `${fmtP(t.n ? t.noComp / t.n : null)} of emails need a company assigned`) +
    tile('Cases', fmtN(t.cases.size), `${t.cases.size ? (t.n / t.cases.size).toFixed(1) : '–'} emails per case · ${plural(t.comps.size, 'customer')}`);
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
// Stacked columns per day: replied, read, new (unread) on top.
function dailyChart(el) {
  el.innerHTML = '';
  const g = groupBy(F, 1), days = [...g.keys()].sort((a, b) => a - b);
  if (!days.length) { el.innerHTML = emptyMsg('No emails for these filters.'); return; }
  const c = { muted: css('--muted'), grid: css('--o-grid'), axis: css('--o-axis'), s1: css('--s1'), s2: css('--s2'), s3: css('--s3') };
  const T = days.map((d) => tally(g.get(d)));
  const W = el.clientWidth || 600, H = 260, m = { l: 46, r: 12, t: 14, b: 28 }, iw = W - m.l - m.r, ih = H - m.t - m.b, n = days.length, bw = iw / n;
  const max = niceMax(Math.max(...T.map((t) => t.n)));
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'Emails received per day by status' }, el);
  const Y = (v) => m.t + ih * (1 - v / max);
  for (let k = 0; k <= 4; k++) {
    const v = (max / 4) * k;
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: k === 0 ? c.axis : c.grid, 'stroke-width': 1 }, svg);
    svgEl('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = fmtN(v);
  }
  const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(iw / 44)))), bar = Math.max(3, Math.min(30, bw - 4)), groups = [];
  days.forEach((d, i) => {
    const x = m.l + i * bw + (bw - bar) / 2, t = T[i], gEl = svgEl('g', { opacity: isWeekday(d) ? 1 : 0.55 }, svg);
    let base = 0;
    for (const [v, color] of [[t.rp + t.other, c.s3], [t.rd, c.s1], [t.nw, c.s2]]) {
      if (v) svgEl('rect', { x, y: Y(base + v), width: bar, height: Math.max(0.5, ih * v / max), fill: color }, gEl);
      base += v;
    }
    groups.push(gEl);
    if (i % every === 0) svgEl('text', { x: m.l + (i + 0.5) * bw, y: H - 8, 'text-anchor': 'middle', 'font-size': 12, fill: c.muted }, svg).textContent = md(d);
  });
  const hit = svgEl('rect', { x: m.l, y: m.t, width: iw, height: ih, fill: 'transparent' }, svg);
  hit.addEventListener('mousemove', (ev) => {
    const r = svg.getBoundingClientRect(), i = Math.max(0, Math.min(n - 1, Math.floor(((ev.clientX - r.left) * (W / r.width) - m.l) / bw))), t = T[i];
    groups.forEach((gEl, j) => gEl.setAttribute('opacity', j === i ? 1 : 0.45));
    showTip(`<div class="tt">${DOW[dowOf(days[i])]} ${dLong(days[i])}</div>` + tipRow('Emails received', fmtN(t.n)) + tipRow('Replied', fmtN(t.rp + t.other), c.s3) +
      tipRow('Read', fmtN(t.rd), c.s1) + tipRow('New (unread)', fmtN(t.nw), c.s2) + tipRow('From customers', fmtN(t.n - t.ceva)) + tipRow('Cases', fmtN(t.cases.size)), ev);
  });
  hit.addEventListener('mouseleave', () => { groups.forEach((gEl, j) => gEl.setAttribute('opacity', isWeekday(days[j]) ? 1 : 0.55)); hideTip(); });
}
function topCustomers() {
  const g = [...groupBy(F, 2)].map(([c, rs]) => [c, rs.length, rs.filter((r) => r[5] === ST.nw).length]).sort((a, b) => b[1] - a[1]).slice(0, 10);
  if (!g.length) { $('#topCust').innerHTML = emptyMsg('No emails for these filters.'); return; }
  const max = g[0][1];
  $('#topCust').innerHTML = g.map(([c, n, nw]) => `<div class="b" title="${esc(D.companies[c])}: ${fmtN(n)} emails, ${fmtN(nw)} unread"><span>${esc(D.companies[c])}</span>
    <span class="track"><i style="width:${(100 * (n - nw)) / max}%;background:var(--s1)"></i><i style="width:${(100 * nw) / max}%;background:var(--s2)"></i></span><span>${fmtN(n)}</span></div>`).join('');
}
function dayModel() {
  const g = groupBy(F, 1), days = [...g.keys()].sort((a, b) => b - a);
  const cols = [col('Day'), col('Emails'), col('From customers'), col('From CEVA'), col('New (unread)'), col('Read'), col('Replied'), col('Cases'), col('Customers')];
  const line = (label, t, cls) => ({ cls, c: [cS(label), cZ(t.n), cZ(t.n - t.ceva), cZ(t.ceva), cZ(t.nw), cZ(t.rd), cZ(t.rp + t.other), cZ(t.cases.size), cZ(t.comps.size)] });
  const rows = days.map((d) => line(`${DOW[dowOf(d)]} ${md(d)}`, tally(g.get(d)), isWeekday(d) ? '' : 'muted'));
  if (days.length > 1) rows.push(line('Total', tally(F), 'total'));
  return { cols, rows };
}

// ---------- by customer ----------
function custModel(search = '') {
  const q = search.trim().toLowerCase(), g = groupBy(F, 2), days = [...new Set(F.map((r) => r[1]))].sort((a, b) => a - b);
  let list = [...g.keys()].filter((c) => !q || D.companies[c].toLowerCase().includes(q));
  list.sort(S.custSort === 'most' ? (a, b) => g.get(b).length - g.get(a).length || byName(D.companies[a], D.companies[b]) : (a, b) => byName(D.companies[a], D.companies[b]));
  const cols = [col('Customer'), col('Emails', 'tot'), col('New (unread)'), col('Replied'), col('From CEVA'), col('Cases'),
    ...days.map((d) => ({ ...col(`${DOW[dowOf(d)]}<small>${md(d)}</small>`, isWeekday(d) ? '' : 'muted', `${DOW[dowOf(d)]} ${md(d)}`), d }))];
  const line = (label, rs, cls) => {
    const t = tally(rs), byDay = groupBy(rs, 1);
    return { cls, c: [cS(label, NOCOMP.test(label) ? 'low' : ''), cZ(t.n), cN(t.nw, t.nw ? 'low' : ''), cP(t.n ? (t.rp + t.other) / t.n : null), cN(t.ceva), cN(t.cases.size),
      ...days.map((d) => cN((byDay.get(d) || []).length))] };
  };
  const rows = list.map((c) => line(D.companies[c], g.get(c)));
  if (list.length > 1) rows.push(line('Total', list.flatMap((c) => g.get(c)), 'total'));
  return { cls: 'mx', cols, rows, n: list.length };
}

// ---------- unread ----------
function unreadRows() {
  return F.filter((r) => r[5] === ST.nw && (S.scope === 'all' || r[11] === 1))
    .sort((a, b) => b[1] - a[1] || byText(D.companies[a[2]], D.companies[b[2]]) || byText(a[6], b[6]));
}
function unreadModel(limit) {
  const all = unreadRows(), rs = limit ? all.slice(0, limit) : all;
  const cols = [col('Received'), col('Customer'), col('From'), col('Email subject'), col('Case #'), col('Case'), col('Case age')];
  return { cols, total: all.length, empty: S.scope === 'open' ? 'No unread emails on open cases.' : 'No unread emails for these filters.',
    rows: rs.map((r) => ({ c: [cS(`${DOW[dowOf(r[1])]} ${md(r[1])}`), cS(D.companies[r[2]], 'left'), cS(D.froms[r[3]], 'left'), cS(D.subjects[r[4]], 'wrap'), cS(r[6]),
      cS(r[11] === 1 ? 'Open' : r[11] === 0 ? 'Closed' : '', r[11] === 1 ? 'low' : 'muted'), cS(fmtAge(r[10]))] })) };
}

// ---------- senders ----------
function domainModel() {
  const g = groupBy(F, 7), list = [...g.keys()].sort((a, b) => g.get(b).length - g.get(a).length).slice(0, 300);
  const cols = [col('Sender domain'), col('Emails'), col('New (unread)'), col('Customers'), col('Main customer')];
  return { cols, rows: list.map((k) => { const rs = g.get(k), t = tally(rs), top = [...groupBy(rs, 2)].sort((a, b) => b[1].length - a[1].length)[0];
    return { c: [cS(D.domains[k] || '(no address)', 'left'), cZ(t.n), cN(t.nw), cZ(t.comps.size), cS(top ? D.companies[top[0]] : '', 'left')] }; }) };
}
function fromModel() {
  const g = groupBy(F, 3), list = [...g.keys()].sort((a, b) => g.get(b).length - g.get(a).length).slice(0, 300);
  const cols = [col('Sender'), col('Emails'), col('New (unread)'), col('Cases'), col('Main customer')];
  return { cols, rows: list.map((k) => { const rs = g.get(k), t = tally(rs), top = [...groupBy(rs, 2)].sort((a, b) => b[1].length - a[1].length)[0];
    return { c: [cS(D.froms[k], 'left'), cZ(t.n), cN(t.nw), cZ(t.cases.size), cS(top ? D.companies[top[0]] : '', 'left')] }; }) };
}

// ---------- how it's calculated ----------
function about() {
  $('#about').innerHTML = `<p>Every row of the Salesforce emails report is one email received on a case.</p>
    <ul>
      <li><b>Email status</b>: <b>New</b> = not opened in Salesforce yet (unread), <b>Read</b> = opened, <b>Replied</b> = answered from the case.</li>
      <li><b>From CEVA addresses</b>: the sender's address ends in @cevalogistics.com (stations and colleagues); everything else counts as from customers.</li>
      <li><b>No company on the case</b>: the case's company is blank or "&lt;&lt; ASSIGN CORRECT COMPANY NOW &gt;&gt;".</li>
      <li><b>Case</b> open or closed, and <b>case age</b> (hours since the case was opened), are as of when the report was run.</li>
      <li>Days are the email's received date. Weekends are shown lighter. The report's grouping (Subtotal rows) is ignored; each email counts once, and files that overlap don't double count.</li>
      <li>Several exports (for example one per branch) can be loaded together; each file's branch comes from its "Branch Code equals …" filter.</li>
    </ul>`;
  renderTable($('#tFiles'), { cols: [col('File'), col('Branch'), col('Emails'), col('Days'), col('Report run')],
    rows: D.files.map((f) => ({ cls: f.rows ? '' : 'muted', c: [cS(f.name, 'left'), cS(f.team || '–'), cN(f.rows), cS(f.from != null ? span(f.from, f.to) : f.note || ''), cS(f.asOf || '')] })) });
}

// ---------- render ----------
function render() {
  if (!D) return;
  filter(); hideTip();
  const sel = S.teams.length === D.teams.length ? (D.teams.length > 1 ? 'all branches' : D.teams[0]) : S.teams.map((t) => D.teams[t]).sort(byText).join(', ');
  $('#eyebrow').textContent = `Emails received · ${sel}${S.cust >= 0 ? ' · ' + D.companies[S.cust] : ''}`;
  $('#range').innerHTML = `<b>${span(S.from, S.to)}</b> · ${fmtN(F.length)} emails`;
  const t = S.tab;
  if (t === 'over') { renderTiles(); dailyChart($('#cDaily')); topCustomers(); renderTable($('#tDay'), dayModel()); }
  if (t === 'cust') {
    choice('#custSort', 'custSort', { az: 'A to Z', most: 'Most emails' });
    const m = custModel($('#custSearch').value);
    $('#custSub').textContent = `${plural(m.n, 'customer')}; red = unread or no company`;
    renderTable($('#mCust'), m, false);
  }
  if (t === 'unread') {
    choice('#unreadScope', 'scope', { all: 'All unread', open: 'On open cases' }, () => { S.limit = 300; });
    const m = unreadModel(S.limit);
    $('#unreadSub').textContent = `${fmtN(m.total)} email${m.total === 1 ? '' : 's'}, newest first`;
    renderTable($('#tUnread'), m);
    $('#unreadMore').innerHTML = m.total > m.rows.length ? `<button class="er-linkbtn" type="button" id="moreBtn">Show all ${fmtN(m.total)}</button><span class="note">Showing ${fmtN(m.rows.length)}.</span>` : '';
    const mb = $('#moreBtn'); if (mb) mb.onclick = () => { S.limit = 0; render(); };
  }
  if (t === 'send') { renderTable($('#tDomain'), domainModel()); renderTable($('#tFrom'), fromModel()); }
  if (t === 'about') about();
}
$('#custSearch').addEventListener('input', () => { if (D && S.tab === 'cust') render(); });
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
  return c.f === 'p' ? { t: 'n', v: c.v, z: '0%' } : { t: 'n', v: c.v, z: '#,##0' };
}
async function exportExcel() {
  const btn = $('#exportBtn'), label = btn.textContent;
  btn.disabled = true; btn.textContent = 'Preparing…';
  try {
    const XLSX = await loadXlsx(), wb = XLSX.utils.book_new(), filt = $('#eyebrow').textContent + ' · ' + span(S.from, S.to);
    const add = (name, title, m, w = 30) => {
      const ws = XLSX.utils.aoa_to_sheet([[title], [filt], [], m.cols.map((c) => c.x), ...m.rows.map((r) => r.c.map(xCell))]);
      ws['!cols'] = m.cols.map((c, i) => ({ wch: i === 0 ? w : Math.max(9, Math.min(40, (c.x || '').length + 2)) }));
      XLSX.utils.book_append_sheet(wb, ws, name);
    };
    const keepSort = S.custSort;
    add('By day', 'Emails received by day', dayModel(), 12);
    add('By customer', 'Emails received by customer and day', custModel(), 38);
    add('Unread emails', 'Unread (New) emails', unreadModel(0), 12);
    add('Sender domains', 'Emails by sender domain', domainModel());
    add('Senders', 'Emails by sender', fromModel());
    S.custSort = keepSort;
    const d = new Date(), pad = (n) => String(n).padStart(2, '0');
    XLSX.writeFile(wb, `Emails Received ${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}.xlsx`);
  } catch (err) {
    $('#notice').hidden = false; $('#notice').textContent = err.message;
  } finally { btn.disabled = false; btn.textContent = label; }
}
$('#exportBtn').addEventListener('click', () => { if (D) exportExcel(); });
