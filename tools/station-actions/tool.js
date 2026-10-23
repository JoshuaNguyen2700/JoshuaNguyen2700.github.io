// Logic for the station-actions tool only. It is an ES module, so nothing here leaks into other pages.
// The page ships with no shipment data. The visitor picks the folder the "Station Dashboard – daily export"
// Power Automate flow writes to (a SharePoint library folder, synced to OneDrive): one CSV per day with every
// HAWB on six action cards of the Power BI Station / Customer Dashboard (column KPI = the card, Card Total =
// that card's full count). Files are read in the browser; nothing is uploaded.
// Every day's file is kept small (card, HAWB and the filter columns) for the counts and the trend; only the
// day on screen is read in full, for the HAWB list and the export.
import { esc } from '/assets/core/util.js';
import { unzip } from '/assets/core/unzip.js';

const $ = (s) => document.querySelector(s);
const DAY = 864e5;
const STATE_KEY = 'sa.state';   // the selected card only; never data

// The six cards, in the order the Power BI page shows them. KPI values in the file match these names.
const KPIS = ["NSR's Past Due", 'B2C OHD > 24 Hours w/o APPSET', 'B2C Past Appt', 'B2C Return Needed', 'B2C Unable to Contact', 'Beyonds Showing OFD Or Not Scheduled'];
const SHORT = ['NSR past due', 'B2C OHD no APPSET', 'B2C past appt', 'B2C return needed', 'B2C unable to contact', 'Beyonds OFD / not sched.'];
const SHEET = ['NSR Past Due', 'B2C OHD no APPSET', 'B2C Past Appt', 'B2C Return Needed', 'B2C Unable to Contact', 'Beyonds OFD or Not Sched'];
// filter columns: the key in F, the CSV column, and the list's name
const DIMS = [['stn', 'Destination Station Airport Code', 'stations'], ['cust', 'Customer Code', 'customers'], ['div', 'Station Division Code', 'divisions'], ['prod', 'Product', 'products']];

let SNAPS = [];          // one per day, oldest first: { day, name, file, k, hawb, stn, cust, div, prod, totals, rows }
let cur = -1;            // index of the day on screen
let FULL = null;         // { snap, head, rows } the day on screen, every column
let PREV = null;         // { snap, keys } HAWBs per card on the day before the one on screen
let sel = 0;             // selected card
const F = { stn: new Set(), cust: new Set(), div: new Set(), prod: new Set() };
const L = { q: '', newOnly: false, sort: 'Due', dir: 1, shown: 200 };
try { const s = JSON.parse(localStorage.getItem(STATE_KEY) || '{}'); if (s.sel >= 0 && s.sel < KPIS.length) sel = s.sel; } catch (e) {}
const saveState = () => { try { localStorage.setItem(STATE_KEY, JSON.stringify({ sel })); } catch (e) {} };

// ---------- formatting ----------
const fmtN = (n) => (n == null ? '' : n.toLocaleString('en-US'));
const dLabel = (d, year) => new Date(d * DAY).toLocaleDateString('en-US', { month: 'short', day: 'numeric', ...(year ? { year: 'numeric' } : {}), timeZone: 'UTC' });
const dIso = (d) => new Date(d * DAY).toISOString().slice(0, 10);
// "10/20/2026" or "10/21/2026 7:25:03 AM" -> days since 1970-01-01
const dayOf = (s) => { const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s || ''); return m ? Date.UTC(+m[3], m[1] - 1, +m[2]) / DAY : null; };
const signed = (n) => (n > 0 ? '+' : n < 0 ? '−' : '') + fmtN(Math.abs(n));

// ---------- reading files ----------
// RFC 4180 CSV: quoted fields may hold commas, quotes ("") and line breaks (the notes do)
function parseCSV(text) {
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n') { row.push(f); rows.push(row); row = []; f = ''; }
    else if (c !== '\r') f += c;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}
const TABLE = /\.(csv|xlsx|xls)$/i;
async function readTable(file) {
  let rows;
  if (/\.csv$/i.test(file.name)) rows = parseCSV((await file.text()).replace(/^﻿/, ''));
  else {   // a daily file opened and saved in Excel: first sheet, values as Excel shows them
    await loadXlsx();
    const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
    rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '', blankrows: false }).map((r) => r.map(String));
  }
  const head = (rows.shift() || []).map((h) => h.trim().replace(/^\[|\]$/g, ''));
  const col = Object.fromEntries(head.map((h, i) => [h, i]));
  return { head, col, rows: rows.filter((r) => r.length > 1) };
}
// The day a file is for: the date in its name (the flow writes "Station Dashboard Actions 2026-10-23.csv"),
// else the day it was last saved.
function fileDay(file) {
  const m = /(\d{4})-(\d{2})-(\d{2})/.exec(file.name);
  if (m) return Date.UTC(+m[1], m[2] - 1, +m[3]) / DAY;
  const d = new Date(file.lastModified);
  return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / DAY;
}
async function readSnap(file) {
  const t = await readTable(file);
  if (t.col.KPI == null || t.col.HAWB == null) return { skip: 'not a Station Dashboard export (no KPI and HAWB columns)' };
  const s = { day: fileDay(file), name: file.sourcePath || file.webkitRelativePath || file.name, file, rows: t.rows.length,
    k: new Int8Array(t.rows.length), hawb: [], stn: [], cust: [], div: [], prod: [], totals: KPIS.map(() => null), counts: KPIS.map(() => 0), unknown: 0 };
  const ix = DIMS.map(([, c]) => t.col[c]);
  t.rows.forEach((r, i) => {
    const k = KPIS.indexOf(r[t.col.KPI]);
    s.k[i] = k;
    if (k < 0) { s.unknown++; return; }
    s.counts[k]++;
    if (t.col['Card Total'] != null && s.totals[k] == null) s.totals[k] = +r[t.col['Card Total']] || null;
    s.hawb.push(r[t.col.HAWB]); s.stn.push(r[ix[0]] ?? ''); s.cust.push(r[ix[1]] ?? ''); s.div.push(r[ix[2]] ?? ''); s.prod.push(r[ix[3]] ?? '');
  });
  // keep the per-row arrays aligned with k (rows with an unknown card were skipped above)
  if (s.unknown) s.k = Int8Array.from([...s.k].filter((k) => k >= 0));
  return s;
}

// ---------- loading ----------
function showStartError(msg) { const e = $('#startError'); e.textContent = msg; e.hidden = !msg; }
const NO_FILES = 'No .csv or .xlsx files there. Choose the Customer_Station report folder (Customer Service Center-Josh Test SharePoint), the daily files in it, or a .zip of them.';
// A .zip counts as a folder named after it; zips inside zips are opened too. Each file inside remembers
// where it came from ("Week 43.zip/Station Dashboard Actions 2026-10-23.csv") for the list and the Files table.
const ZIP = /\.zip$/i;
async function expandZips(files, depth = 0) {
  const out = [];
  for (const f of files) {
    if (!ZIP.test(f.name) || depth > 3) { out.push(f); continue; }
    let inner;
    try { inner = await unzip(f); } catch (err) { throw new Error(`${f.name} could not be opened: ${err.message}.`); }
    const folder = (f.sourcePath || f.webkitRelativePath || f.name).replace(ZIP, ''), base = folder.split('/').pop();
    // a zip made from a folder of the same name ("Week 43.zip" holding "Week 43/...") doesn't repeat it
    const same = inner.length && inner.every((x) => x.name.toLowerCase().startsWith(base.toLowerCase() + '/'));
    for (const x of inner) x.file.sourcePath = folder + '/' + (same ? x.name.slice(base.length + 1) : x.name);
    out.push(...await expandZips(inner.map((x) => x.file), depth + 1));
  }
  return out;
}
// Folders, files and zips picked or dropped all come through here
function take(files, now) {
  const go = (fs) => (now ? loadFiles(fs) : stage(fs));
  if (!files.some((f) => ZIP.test(f.name))) { go(files); return; }
  expandZips(files).then(go, (err) => (now ? note(err.message) : showStartError(err.message)));
}
const usable = (files) => files.filter((f) => TABLE.test(f.name) && !/^~\$/.test(f.name));
async function loadFiles(files) {
  files = usable(files);
  if (!files.length) { (SNAPS.length ? note : showStartError)(NO_FILES); return; }
  const skipped = [];
  const byDay = new Map(SNAPS.map((s) => [s.day, s]));
  for (const f of files) {
    try {
      const s = await readSnap(f);
      if (s.skip) { skipped.push(`${f.name} (${s.skip})`); continue; }
      // one file per day: two for the same day (a re-run of the flow, a download copy) keep the newest
      const old = byDay.get(s.day);
      if (old && old.file !== f) {
        const [keep, drop] = old.file.lastModified > f.lastModified ? [old, s] : [s, old];
        skipped.push(`${drop.name} (${dLabel(s.day, true)} already has ${keep.name}, which is newer)`);
        if (keep === old) continue;
      }
      byDay.set(s.day, s);
    } catch (e) { skipped.push(`${f.name} (${e.message})`); }
  }
  SNAPS = [...byDay.values()].sort((a, b) => a.day - b.day);
  if (!SNAPS.length) { showStartError('None of those files are Station Dashboard exports: ' + skipped.join(', ')); return; }
  showStartError('');
  $('#start').hidden = true; $('#app').hidden = false;
  NOTES.load = skipped.length ? 'Skipped: ' + skipped.join('; ') + '.' : '';
  await setDay(SNAPS.length - 1);
}
const NOTES = { load: '', short: '' };
function note(msg) { if (msg !== undefined) NOTES.load = msg; const t = [NOTES.load, NOTES.short].filter(Boolean).join(' '); $('#notice').textContent = t; $('#notice').hidden = !t; }

async function setDay(i) {
  cur = i;
  const s = SNAPS[cur];
  if (!FULL || FULL.snap !== s) { const t = await readTable(s.file); FULL = { snap: s, head: t.head, col: t.col, rows: t.rows.filter((r) => KPIS.includes(r[t.col.KPI])) }; }
  const p = SNAPS[cur - 1];
  if (!p) PREV = null;
  else if (!PREV || PREV.snap !== p) { const keys = new Set(); p.hawb.forEach((h, j) => keys.add(p.k[j] + '\u0001' + h)); PREV = { snap: p, keys }; }
  // a card with fewer rows than its Card Total means the Power BI query came back short (it stops at 15 MB)
  const short = KPIS.map((k, j) => (s.totals[j] != null && s.counts[j] < s.totals[j] ? `${k}: ${fmtN(s.counts[j])} of ${fmtN(s.totals[j])}` : '')).filter(Boolean);
  NOTES.short = short.length ? `The ${dLabel(s.day)} file is missing rows, so its numbers are low (${short.join('; ')}). Run the flow again or split the query.` : '';
  L.shown = 200;
  buildFilters();
  render();
}

// ---------- filters ----------
const pass = (s, j) => DIMS.every(([key], d) => !F[key].size || F[key].has(s[key][j]));
const isNew = (k, h) => PREV && !PREV.keys.has(k + '\u0001' + h);
function countsOf(s) { const c = KPIS.map(() => 0); for (let j = 0; j < s.hawb.length; j++) if (pass(s, j)) c[s.k[j]]++; return c; }

function msel(root, key, plural) {
  root.innerHTML = `<button class="ms-btn" type="button" aria-haspopup="true" aria-expanded="false"></button>
    <div class="ms-pop" hidden><input type="search" placeholder="Search" aria-label="Search ${plural}">
    <div class="ms-acts"><button class="sa-linkbtn" type="button" data-act="all">Tick all shown</button><button class="sa-linkbtn" type="button" data-act="none">Clear</button></div>
    <div class="ms-list"></div></div>`;
  const btn = root.querySelector('.ms-btn'), pop = root.querySelector('.ms-pop'), q = root.querySelector('input'), list = root.querySelector('.ms-list');
  let items = [];
  const shown = () => { const t = q.value.trim().toLowerCase(); return items.filter((it) => !t || it.v.toLowerCase().includes(t)); };
  const label = () => {
    const s = F[key];
    btn.textContent = !s.size ? 'All ' + plural : s.size === 1 ? ([...s][0] || '(blank)') : s.size + ' ' + plural;
    btn.classList.toggle('on', s.size > 0);
  };
  const draw = () => {
    const sh = shown();
    list.innerHTML = sh.length ? sh.slice(0, 600).map((it) => `<label><input type="checkbox" data-v="${esc(it.v)}"${F[key].has(it.v) ? ' checked' : ''}><span>${esc(it.v || '(blank)')}</span><small>${fmtN(it.n)}</small></label>`).join('')
      : '<div class="ms-empty">No matches</div>';
  };
  const close = () => { pop.hidden = true; btn.setAttribute('aria-expanded', 'false'); };
  btn.addEventListener('click', () => { const open = pop.hidden; document.querySelectorAll('.ms-pop').forEach((p) => { p.hidden = true; }); if (open) { pop.hidden = false; btn.setAttribute('aria-expanded', 'true'); q.value = ''; draw(); q.focus(); } else close(); });
  q.addEventListener('input', draw);
  list.addEventListener('change', (e) => { const v = e.target.dataset.v; if (v == null) return; if (e.target.checked) F[key].add(v); else F[key].delete(v); label(); changed(); });
  root.querySelector('.ms-acts').addEventListener('click', (e) => {
    const a = e.target.dataset.act; if (!a) return;
    if (a === 'none') F[key].clear(); else for (const it of shown()) F[key].add(it.v);
    draw(); label(); changed();
  });
  document.addEventListener('click', (e) => { if (!root.contains(e.target)) close(); });
  root.addEventListener('keydown', (e) => { if (e.key === 'Escape') { close(); btn.focus(); } });
  return { setItems(x) { items = x; label(); if (!pop.hidden) draw(); }, label };
}
const MS = {};
DIMS.forEach(([key, , plural]) => { MS[key] = msel($('#ms' + key[0].toUpperCase() + key.slice(1)), key, plural); });
// Each list offers the values on the day on screen, with how many HAWBs (all cards) each has
function buildFilters() {
  const s = SNAPS[cur];
  for (const [key] of DIMS) {
    const m = new Map();
    s[key].forEach((v) => m.set(v, (m.get(v) || 0) + 1));
    for (const v of F[key]) if (!m.has(v)) m.set(v, 0);   // a ticked value the day doesn't have stays visible
    MS[key].setItems([...m].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0])).map(([v, n]) => ({ v, n })));
  }
}
let queued = false;
function changed() { L.shown = 200; if (queued) return; queued = true; requestAnimationFrame(() => { queued = false; render(); }); }
$('#resetBtn').addEventListener('click', () => { for (const [key] of DIMS) F[key].clear(); Object.values(MS).forEach((m) => m.label()); buildFilters(); changed(); });
function setOnly(key, v) {
  const s = F[key];
  if (s.size === 1 && s.has(v)) s.clear(); else { s.clear(); s.add(v); }
  MS[key].label(); changed();
}

// ---------- render ----------
function render() {
  const s = SNAPS[cur], p = SNAPS[cur - 1];
  $('#range').textContent = `${dLabel(s.day, true)}` + (SNAPS.length > 1 ? ` · ${SNAPS.length} days loaded (${dLabel(SNAPS[0].day)} – ${dLabel(SNAPS[SNAPS.length - 1].day, true)})` : ' · load more days for the trend') +
    (p ? ` · changes are against ${dLabel(p.day)}` : '');
  $('#snap').innerHTML = SNAPS.map((x, i) => `<option value="${i}"${i === cur ? ' selected' : ''}>${dLabel(x.day, true)}</option>`).reverse().join('');
  note();
  const series = SNAPS.map(countsOf);
  renderKpis(series);
  renderTrend(series);
  renderBars('#byStn', 'stn', 'stations');
  renderBars('#byCust', 'cust', 'customers');
  renderMatrix();
  renderList();
  renderFiles();
}

function renderKpis(series) {
  const now = series[cur], before = series[cur - 1], s = SNAPS[cur];
  const fresh = KPIS.map(() => 0);
  if (PREV) for (let j = 0; j < s.hawb.length; j++) if (pass(s, j) && isNew(s.k[j], s.hawb[j])) fresh[s.k[j]]++;
  $('#kpis').innerHTML = KPIS.map((k, i) => {
    const d = before ? now[i] - before[i] : null;
    const short = s.totals[i] != null && s.counts[i] < s.totals[i];
    return `<button class="kpi" type="button" data-k="${i}" aria-pressed="${i === sel}">
      <span class="t">${esc(k)}</span><span class="v">${fmtN(now[i])}</span>
      <span class="d">${d == null ? '<span>no day before loaded</span>' : `<span class="${d > 0 ? 'up' : d < 0 ? 'down' : ''}" title="Change against ${dLabel(SNAPS[cur - 1].day)}">${d > 0 ? '▲' : d < 0 ? '▼' : '='} ${signed(d)}</span><span>${fmtN(fresh[i])} new</span>`}</span>
      ${short ? `<span class="warn">File short: ${fmtN(s.counts[i])} of ${fmtN(s.totals[i])}</span>` : ''}
      ${SNAPS.length > 1 ? '<svg aria-hidden="true"></svg>' : ''}</button>`;
  }).join('');
  if (SNAPS.length > 1) $('#kpis').querySelectorAll('.kpi').forEach((b, i) => spark(b.querySelector('svg'), series.map((c) => c[i])));
}
$('#kpis').addEventListener('click', (e) => { const b = e.target.closest('.kpi'); if (!b) return; sel = +b.dataset.k; saveState(); L.shown = 200; render(); });

// A one-series sparkline, drawn at the tile's real width so the dot stays round
function spark(svg, vals) {
  const w = svg.clientWidth || 160, h = 30, max = Math.max(1, ...vals), min = Math.min(...vals), pad = 3;
  const x = (i) => pad + (i * (w - 2 * pad)) / Math.max(1, vals.length - 1);
  const y = (v) => h - pad - ((v - min) / Math.max(1, max - min)) * (h - 2 * pad);
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  svg.innerHTML = `<path d="${vals.map((v, i) => (i ? 'L' : 'M') + x(i).toFixed(1) + ',' + y(v).toFixed(1)).join('')}" fill="none" style="stroke:var(--s1)" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>` +
    `<circle cx="${x(cur)}" cy="${y(vals[cur])}" r="3" style="fill:var(--s1);stroke:var(--bg)" stroke-width="1.5"/>`;
}

// ---------- trend chart (the selected card, one series) ----------
const tip = $('#tip');
function showTip(html, ev) {
  tip.innerHTML = html; tip.hidden = false;
  const pad = 14, w = tip.offsetWidth, h = tip.offsetHeight;
  let x = ev.clientX + pad, y = ev.clientY + pad;
  if (x + w > innerWidth - 8) x = ev.clientX - w - pad;
  if (y + h > innerHeight - 8) y = ev.clientY - h - pad;
  tip.style.left = Math.max(8, x) + 'px'; tip.style.top = Math.max(8, y) + 'px';
}
const hideTip = () => { tip.hidden = true; };
const tipRow = (k, v) => `<div class="r"><span>${k}</span><span>${v}</span></div>`;
function niceMax(v) { if (v <= 5) return 5; const p = 10 ** Math.floor(Math.log10(v)); const m = v / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p; }

function renderTrend(series) {
  const box = $('#trend'), vals = series.map((c) => c[sel]);
  $('#trendTitle').textContent = KPIS[sel];
  if (SNAPS.length < 2) {
    $('#trendSub').textContent = '';
    box.innerHTML = '<p class="note">One day is loaded. The flow saves a file every morning; choose the whole folder to see how this card moves day to day.</p>';
    return;
  }
  $('#trendSub').textContent = 'HAWBs on the card each day, with the filters above · click a day to open it';
  const W = Math.max(300, box.clientWidth), H = 220, ml = 46, mr = 40, mt = 12, mb = 26;
  const d0 = SNAPS[0].day, d1 = SNAPS[SNAPS.length - 1].day;
  const X = (d) => ml + ((d - d0) / Math.max(1, d1 - d0)) * (W - ml - mr);
  const top = niceMax(Math.max(...vals) * 1.05), Y = (v) => mt + (1 - v / top) * (H - mt - mb);
  const ticks = [0, 0.25, 0.5, 0.75, 1].map((f) => f * top);
  const every = Math.ceil(SNAPS.length / Math.max(2, Math.floor((W - ml - mr) / 70)));
  let g = ticks.map((t) => `<line x1="${ml}" x2="${W - mr}" y1="${Y(t)}" y2="${Y(t)}" style="stroke:${t ? 'var(--o-grid)' : 'var(--o-axis)'}"/>` +
    `<text x="${ml - 8}" y="${Y(t) + 4}" text-anchor="end" font-size="11" style="fill:var(--muted)">${fmtN(Math.round(t))}</text>`).join('');
  g += SNAPS.map((s, i) => (i % every && i !== SNAPS.length - 1 ? '' : `<text x="${X(s.day)}" y="${H - 6}" text-anchor="middle" font-size="11" style="fill:var(--muted)">${dLabel(s.day)}</text>`)).join('');
  const path = SNAPS.map((s, i) => (i ? 'L' : 'M') + X(s.day).toFixed(1) + ',' + Y(vals[i]).toFixed(1)).join('');
  const dots = SNAPS.length <= 90 ? SNAPS.map((s, i) => `<circle cx="${X(s.day)}" cy="${Y(vals[i])}" r="${i === cur ? 5 : 3.5}" style="fill:var(--s1);stroke:var(--bg)" stroke-width="2"/>`).join('') : '';
  const last = SNAPS.length - 1;
  box.innerHTML = `<svg viewBox="0 0 ${W} ${H}" height="${H}" role="img" aria-label="${esc(KPIS[sel])} per day">${g}
    <path d="${path}" fill="none" style="stroke:var(--s1)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>${dots}
    <text x="${X(SNAPS[last].day) + 8}" y="${Y(vals[last]) + 4}" font-size="12" font-weight="600" style="fill:var(--text)">${fmtN(vals[last])}</text>
    <line class="xh" x1="0" x2="0" y1="${mt}" y2="${H - mb}" style="stroke:var(--o-axis)" visibility="hidden"/>
    <rect x="${ml}" y="0" width="${W - ml - mr}" height="${H}" fill="transparent" style="cursor:pointer"/></svg>`;
  const svg = box.querySelector('svg'), xh = svg.querySelector('.xh'), hit = svg.querySelector('rect');
  const nearest = (ev) => {
    const r = svg.getBoundingClientRect(), px = ((ev.clientX - r.left) / r.width) * W;
    let b = 0; SNAPS.forEach((s, i) => { if (Math.abs(X(s.day) - px) < Math.abs(X(SNAPS[b].day) - px)) b = i; });
    return b;
  };
  hit.addEventListener('mousemove', (ev) => {
    const i = nearest(ev), s = SNAPS[i];
    xh.setAttribute('x1', X(s.day)); xh.setAttribute('x2', X(s.day)); xh.setAttribute('visibility', 'visible');
    showTip(`<div class="tt">${dLabel(s.day, true)}</div>` + tipRow('HAWBs', fmtN(vals[i])) + (i ? tipRow(`vs ${dLabel(SNAPS[i - 1].day)}`, signed(vals[i] - vals[i - 1])) : ''), ev);
  });
  hit.addEventListener('mouseleave', () => { xh.setAttribute('visibility', 'hidden'); hideTip(); });
  hit.addEventListener('click', (ev) => { hideTip(); setDay(nearest(ev)); });
}

// ---------- bar lists (selected card on the day on screen) ----------
function renderBars(el, key, plural) {
  const s = SNAPS[cur], m = new Map();
  let total = 0;
  for (let j = 0; j < s.hawb.length; j++) if (s.k[j] === sel && pass(s, j)) { m.set(s[key][j], (m.get(s[key][j]) || 0) + 1); total++; }
  const all = [...m].sort((a, b) => b[1] - a[1]), top = all.slice(0, 15), rest = all.slice(15);
  const max = top.length ? top[0][1] : 1, restN = rest.reduce((a, x) => a + x[1], 0);
  const bar = (label, n, attrs, cls) => `<button class="b${cls || ''}" type="button" ${attrs}><span title="${esc(label)}">${esc(label)}</span><span class="track"><i style="width:${(n / max) * 100}%"></i></span><span>${fmtN(n)}<small>${total ? Math.round((n / total) * 100) + '%' : ''}</small></span></button>`;
  $(el).innerHTML = !all.length ? '<div class="empty-s">Nothing on this card with these filters.</div>'
    : top.map(([v, n]) => bar(v || '(blank)', n, `data-key="${key}" data-v="${esc(v)}" title="${F[key].has(v) && F[key].size === 1 ? 'Click to clear this filter' : 'Click to filter to ' + esc(v || '(blank)')}"`)).join('') +
      (rest.length ? bar(`Other ${rest.length} ${plural}`, restN, 'tabindex="-1" aria-disabled="true"', ' other') : '');
}
for (const el of ['#byStn', '#byCust']) $(el).addEventListener('click', (e) => { const b = e.target.closest('[data-key]'); if (b) setOnly(b.dataset.key, b.dataset.v); });

// ---------- stations x cards ----------
function renderMatrix() {
  const s = SNAPS[cur], m = new Map();
  for (let j = 0; j < s.hawb.length; j++) if (pass(s, j)) { const v = s.stn[j]; if (!m.has(v)) m.set(v, KPIS.map(() => 0)); m.get(v)[s.k[j]]++; }
  const rows = [...m].map(([v, c]) => [v, c, c.reduce((a, b) => a + b, 0)]).sort((a, b) => b[1][sel] - a[1][sel] || b[2] - a[2]);
  const tot = KPIS.map((_, i) => rows.reduce((a, r) => a + r[1][i], 0));
  $('#mxSub').textContent = `${fmtN(rows.length)} destination stations on ${dLabel(s.day)}, sorted by the selected card · click a card's heading to select it`;
  $('#matrix').innerHTML = !rows.length ? '<div class="empty-s">No HAWBs with these filters.</div>' : `<table><thead><tr><th>Station</th>${SHORT.map((k, i) => `<th class="n click${i === sel ? ' sel' : ''}" data-k="${i}" title="${esc(KPIS[i])}">${esc(k)}</th>`).join('')}<th class="n">All cards</th></tr></thead><tbody>` +
    rows.map(([v, c, t]) => `<tr><td>${esc(v || '(blank)')}</td>${c.map((n) => `<td class="n${n ? '' : ' zero'}">${n ? fmtN(n) : '0'}</td>`).join('')}<td class="n">${fmtN(t)}</td></tr>`).join('') +
    `<tr class="total"><td>Total</td>${tot.map((n) => `<td class="n">${fmtN(n)}</td>`).join('')}<td class="n">${fmtN(tot.reduce((a, b) => a + b, 0))}</td></tr></tbody></table>`;
}
$('#matrix').addEventListener('click', (e) => { const th = e.target.closest('th[data-k]'); if (th) { sel = +th.dataset.k; saveState(); L.shown = 200; render(); } });

// ---------- HAWB list (selected card, every column of the day on screen) ----------
const get = (r, c) => (FULL.col[c] == null ? '' : r[FULL.col[c]] ?? '');
const late = (r) => { const d = dayOf(get(r, 'Due Date')); const n = d == null ? null : SNAPS[cur].day - d; return n > 0 ? n : null; };
const lastStatus = (r) => [get(r, 'Status Code'), get(r, 'Status Date')].filter(Boolean).join(' ');
// h = heading, v = value, t = sort type, cls = cell class
const COLS = [
  { h: 'HAWB', v: (r) => get(r, 'HAWB') },
  { h: 'Dest. station', v: (r) => get(r, 'Destination Station Airport Code') },
  { h: 'Customer', v: (r) => get(r, 'Customer Code') },
  { h: 'Consignee', v: (r) => get(r, 'Consignee Code') },
  { h: 'Service', v: (r) => get(r, 'Service') },
  { h: 'LOS', v: (r) => get(r, 'Level of Service') },
  { h: 'Product', v: (r) => get(r, 'Product') },
  { h: 'Current status', v: (r) => get(r, 'Current Status Code') },
  { h: 'Shipped', v: (r) => get(r, 'Date Shipped'), t: 'd' },
  { h: 'Due', v: (r) => get(r, 'Due Date'), t: 'd' },
  { h: 'Days late', v: late, t: 'n', cls: 'n' },
  { h: 'Appt date', v: (r) => get(r, 'Appointment Delivery Date'), t: 'd' },
  { h: 'Appt set', v: (r) => get(r, 'Delivery Appointment Status') },
  { h: 'Last status', v: lastStatus },
  { h: 'Status by', v: (r) => get(r, 'Status User') },
  { h: 'Note 1', v: (r) => get(r, 'Note 1'), cls: 'note' },
  { h: 'Note 1 time', v: (r) => get(r, 'Note 1 Timestamp'), t: 'd' },
  { h: 'Note 2', v: (r) => get(r, 'Note 2'), cls: 'note' },
];
const fullPass = (r) => DIMS.every(([key, c]) => !F[key].size || F[key].has(get(r, c)));
function listRows() {
  const q = L.q.toLowerCase(), k = KPIS[sel];
  let rows = FULL.rows.filter((r) => get(r, 'KPI') === k && fullPass(r));
  if (L.newOnly && PREV) rows = rows.filter((r) => isNew(sel, get(r, 'HAWB')));
  if (q) rows = rows.filter((r) => ['HAWB', 'Customer Code', 'Consignee Code', 'Shipper Code', 'Note 1', 'Note 2', 'Current Status Code'].some((c) => get(r, c).toLowerCase().includes(q)));
  const c = COLS.find((x) => x.h === L.sort) || COLS[9];
  const key = c.t === 'd' ? (r) => { const v = c.v(r); const d = dayOf(v); if (d == null) return Infinity; const m = /(\d+):(\d+):?(\d*)\s*([AP]M)?/i.exec(v); return m ? d + ((+m[1] % 12 + (/p/i.test(m[4] || '') ? 12 : 0)) * 60 + +m[2]) / 1440 : d; }
    : c.t === 'n' ? (r) => c.v(r) ?? -Infinity : (r) => String(c.v(r)).toLowerCase();
  return rows.map((r) => [key(r), r]).sort((a, b) => (a[0] < b[0] ? -L.dir : a[0] > b[0] ? L.dir : 0)).map((x) => x[1]);
}
function renderList() {
  const rows = listRows(), shown = rows.slice(0, L.shown);
  $('#listTitle').textContent = 'HAWBs: ' + SHORT[sel];
  $('#newOnlyLabel').textContent = PREV ? `New since ${dLabel(PREV.snap.day)} only` : 'New since the day before only';
  $('#newOnly').disabled = !PREV;
  $('#listSub').textContent = `${fmtN(rows.length)} HAWB${rows.length === 1 ? '' : 's'}` + (rows.length > shown.length ? `, first ${fmtN(shown.length)} shown` : '') + (PREV ? ' · New = not on this card the day before' : '');
  $('#list').innerHTML = !rows.length ? '<div class="empty-s">No HAWBs match.</div>' : `<table><thead><tr>${COLS.map((c) => `<th class="click${c.cls === 'n' ? ' n' : ''}" data-sort="${esc(c.h)}"${L.sort === c.h ? ` aria-sort="${L.dir > 0 ? 'ascending' : 'descending'}"` : ''}>${esc(c.h)}</th>`).join('')}</tr></thead><tbody>` +
    shown.map((r) => `<tr>${COLS.map((c, i) => {
      const v = c.v(r);
      if (i === 0) return `<td>${esc(v)}${isNew(sel, v) ? '<span class="tag new">New</span>' : ''}</td>`;
      if (c.h === 'Days late') return `<td class="n${v ? ' late' : ''}">${v ?? ''}</td>`;
      return `<td${c.cls ? ` class="${c.cls}"` : ''}${c.cls === 'note' && v ? ` title="${esc(v)}"` : ''}>${esc(v)}</td>`;
    }).join('')}</tr>`).join('') + '</tbody></table>';
  $('#moreWrap').hidden = rows.length <= shown.length;
}
$('#list').addEventListener('click', (e) => {
  const th = e.target.closest('th[data-sort]'); if (!th) return;
  if (L.sort === th.dataset.sort) L.dir = -L.dir; else { L.sort = th.dataset.sort; L.dir = th.dataset.sort === 'Days late' ? -1 : 1; }
  renderList();
});
$('#more').addEventListener('click', () => { L.shown += 200; renderList(); });
let qT = 0;
$('#q').addEventListener('input', (e) => { clearTimeout(qT); qT = setTimeout(() => { L.q = e.target.value.trim(); L.shown = 200; renderList(); }, 150); });
$('#newOnly').addEventListener('change', (e) => { L.newOnly = e.target.checked; L.shown = 200; renderList(); });
$('#snap').addEventListener('change', (e) => setDay(+e.target.value));

function renderFiles() {
  $('#files').innerHTML = `<table><thead><tr><th>Day</th><th>File</th><th class="n">HAWB rows</th>${SHORT.map((k) => `<th class="n" title="${esc(k)}">${esc(k)}</th>`).join('')}<th></th></tr></thead><tbody>` +
    SNAPS.map((s, i) => `<tr><td>${dLabel(s.day, true)}${i === cur ? '<span class="tag">on screen</span>' : ''}</td><td>${esc(s.name)}</td><td class="n">${fmtN(s.hawb.length)}</td>` +
      s.counts.map((n, j) => `<td class="n${s.totals[j] != null && n < s.totals[j] ? ' late' : ''}"${s.totals[j] != null && n < s.totals[j] ? ` title="Power BI card total ${fmtN(s.totals[j])}"` : ''}>${fmtN(n)}</td>`).join('') +
      `<td>${i === cur ? '' : `<button class="sa-linkbtn" type="button" data-open="${i}">Open</button>`}</td></tr>`).reverse().join('') + '</tbody></table>';
}
$('#files').addEventListener('click', (e) => { const b = e.target.closest('[data-open]'); if (b) setDay(+b.dataset.open); });

// ---------- copy a table as tab-separated text (pastes into Excel as cells) ----------
document.addEventListener('click', async (e) => {
  const b = e.target.closest('[data-copy]'); if (!b) return;
  const t = document.querySelector(b.dataset.copy + ' table'); if (!t) return;
  // (the HAWB cell's "New" badge is left out)
  const text = [...t.rows].map((r) => [...r.cells].map((c) => (c.querySelector('.tag') ? c.firstChild.textContent : c.textContent).trim()).join('\t')).join('\n');
  try { await navigator.clipboard.writeText(text); b.textContent = 'Copied'; } catch (err) { b.textContent = 'Copy failed'; }
  setTimeout(() => { b.textContent = 'Copy'; }, 1500);
});

// ---------- Excel export (filters applied) ----------
function loadXlsx() {
  if (window.XLSX) return Promise.resolve();
  return new Promise((ok, no) => { const sc = document.createElement('script'); sc.src = '/assets/vendor/xlsx.full.min.js'; sc.onload = ok; sc.onerror = () => no(new Error('the Excel library did not load')); document.head.append(sc); });
}
$('#exportBtn').addEventListener('click', async () => {
  const btn = $('#exportBtn'); btn.disabled = true;
  try {
    await loadXlsx();
    const s = SNAPS[cur], p = SNAPS[cur - 1], series = SNAPS.map(countsOf), wb = XLSX.utils.book_new();
    const filt = DIMS.filter(([key]) => F[key].size).map(([key, , plural]) => `${plural}: ${[...F[key]].join(', ')}`).join('; ') || 'none';
    const fresh = KPIS.map((_, k) => FULL.rows.filter((r) => get(r, 'KPI') === KPIS[k] && fullPass(r) && isNew(k, get(r, 'HAWB'))).length);
    const sum = [['Station Actions', dLabel(s.day, true)], ['Filters', filt], [],
      ['Card', 'HAWBs', p ? `HAWBs on ${dLabel(p.day)}` : 'Day before', 'Change', 'New HAWBs', 'Card total in the file (no filters)'],
      ...KPIS.map((k, i) => [k, series[cur][i], p ? series[cur - 1][i] : '', p ? series[cur][i] - series[cur - 1][i] : '', p ? fresh[i] : '', s.totals[i] ?? ''])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sum), 'Summary');
    if (SNAPS.length > 1) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Day', ...KPIS], ...SNAPS.map((x, i) => [dIso(x.day), ...series[i]])]), 'Trend');
    const mx = [['Destination station', ...KPIS, 'All cards']];
    const m = new Map();
    for (let j = 0; j < s.hawb.length; j++) if (pass(s, j)) { const v = s.stn[j]; if (!m.has(v)) m.set(v, KPIS.map(() => 0)); m.get(v)[s.k[j]]++; }
    [...m].sort((a, b) => b[1].reduce((x, y) => x + y, 0) - a[1].reduce((x, y) => x + y, 0)).forEach(([v, c]) => mx.push([v, ...c, c.reduce((x, y) => x + y, 0)]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mx), 'By station');
    const head = FULL.head.filter((h) => h !== 'KPI' && h !== 'Card Total');
    KPIS.forEach((k, i) => {
      const rows = FULL.rows.filter((r) => get(r, 'KPI') === k && fullPass(r));
      const aoa = [[...head, 'Days late', ...(p ? ['New'] : [])], ...rows.map((r) => [...head.map((h) => get(r, h)), late(r) ?? '', ...(p ? [isNew(i, get(r, 'HAWB')) ? 'Yes' : ''] : [])])];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), SHEET[i]);
    });
    XLSX.writeFile(wb, `Station Actions ${dIso(s.day)}.xlsx`);
  } catch (err) { note('Export failed: ' + err.message); }
  btn.disabled = false;
});

// ---------- picking and dropping files ----------
// On the start screen each folder or set of files picked (or dropped) joins a list, and Load reads them all
// together. Once loaded, Add folder / Add files (or a drop) add more straight away. The same file twice counts once.
const fromInput = (fl) => [...(fl || [])];
const fileKey = (f) => `${f.name}|${f.size}|${f.lastModified}`;
const topFolder = (f) => { const p = (f.sourcePath || f.webkitRelativePath || '').split('/'); return p.length > 1 ? p[0] : 'Files you picked'; };
let PENDING = new Map();
function stage(files) {
  files = usable(files);
  if (!files.length) { showStartError(NO_FILES); return; }
  showStartError('');
  for (const f of files) if (!PENDING.has(fileKey(f))) PENDING.set(fileKey(f), f);
  renderPending();
}
function renderPending() {
  const groups = new Map();
  for (const f of PENDING.values()) { const g = topFolder(f); groups.set(g, [...(groups.get(g) || []), f]); }
  $('#pending').hidden = !groups.size; $('#pickRow').hidden = !!groups.size;
  $('#pendingList').innerHTML = [...groups].map(([g, fs]) => `<li><b>${esc(g)}</b><span>${fs.length} file${fs.length === 1 ? '' : 's'}</span><button class="sa-x" type="button" data-unstage="${esc(g)}">Remove</button></li>`).join('');
  $('#loadBtn').textContent = `Load ${PENDING.size} file${PENDING.size === 1 ? '' : 's'}`;
}
$('#pendingList').addEventListener('click', (e) => {
  const b = e.target.closest('[data-unstage]'); if (!b) return;
  for (const [k, f] of PENDING) if (topFolder(f) === b.dataset.unstage) PENDING.delete(k);
  renderPending();
});
$('#loadBtn').addEventListener('click', async () => {
  if (!PENDING.size) return;
  const btn = $('#loadBtn'); btn.disabled = true; btn.textContent = 'Reading…';
  const files = [...PENDING.values()];
  await loadFiles(files);
  btn.disabled = false;
  if (SNAPS.length) PENDING = new Map();
  renderPending();
});
for (const id of ['folderInput', 'filesInput']) $('#' + id).addEventListener('change', (e) => { take(fromInput(e.target.files), false); e.target.value = ''; });
for (const id of ['addFolderInput', 'addFilesInput']) $('#' + id).addEventListener('change', (e) => { take(fromInput(e.target.files), true); e.target.value = ''; });
$('#closeBtn').addEventListener('click', () => {
  SNAPS = []; FULL = null; PREV = null; cur = -1; for (const [key] of DIMS) F[key].clear();
  PENDING = new Map(); renderPending();
  $('#app').hidden = true; $('#start').hidden = false; note('');
});
async function dropFiles(dt) {
  const entries = [...(dt.items || [])].map((i) => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null)).filter(Boolean);
  if (!entries.length) return fromInput(dt.files);
  const out = [];
  const walk = async (entry) => {
    if (entry.isFile) { out.push(await new Promise((ok, no) => entry.file(ok, no))); return; }
    const reader = entry.createReader(); let batch;
    do { batch = await new Promise((ok, no) => reader.readEntries(ok, no)); for (const c of batch) await walk(c); } while (batch.length);
  };
  for (const e of entries) await walk(e);
  return out;
}
let dragDepth = 0;
document.addEventListener('dragenter', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) { e.preventDefault(); dragDepth++; document.body.classList.add('dragging'); } });
document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dragging'); } });
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault(); dragDepth = 0; document.body.classList.remove('dragging');
  if (e.dataTransfer) dropFiles(e.dataTransfer).then((fs) => take(fs, SNAPS.length > 0), (err) => showStartError('Those files could not be opened: ' + err.message));
});
let rT = 0;
addEventListener('resize', () => { clearTimeout(rT); rT = setTimeout(() => { if (cur >= 0) { const series = SNAPS.map(countsOf); renderKpis(series); renderTrend(series); } }, 150); });
