// Logic for the otp tool only. It is an ES module, so nothing here leaks into other pages.
// The page ships with no shipment data. The visitor picks the regional Excel exports (.xlsx); worker.js
// reads them in the browser with parse.js (nothing is uploaded). A saved data file (.json, from
// "Save data file" or build_otp_dashboard.py) opens instantly instead.
import { esc } from '/assets/core/util.js';

const $ = (s) => document.querySelector(s);
const OLD_DATA_KEY = 'otp.data'; // copy kept by the old "Remember" option; deleted on load
const STATE_KEY = 'otp.state';   // filters and tab

// ---------- formatting ----------
const fmtN = (n) => Math.round(n).toLocaleString('en-US');
const fmtP = (p, d = 1) => (p == null || !isFinite(p) ? '–' : (p * 100).toFixed(d) + '%');
const fmt$ = (n) => '$' + Math.round(n).toLocaleString('en-US');
const fmt$c = (n) => n >= 1e6 ? '$' + (n / 1e6).toFixed(2) + 'M' : n >= 1e3 ? '$' + (n / 1e3).toFixed(1) + 'K' : fmt$(n);
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DAY = 86400000;
// "2026-10-01 10:49" -> "Oct 1, 2026, 10:49 AM"
function fmtBuilt(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?/.exec(s || '');
  if (!m) return s || '?';
  let out = `${MON[+m[2] - 1]} ${+m[3]}, ${m[1]}`;
  if (m[4]) { const h = +m[4]; out += `, ${h % 12 || 12}:${m[5]} ${h < 12 ? 'AM' : 'PM'}`; }
  return out;
}
const plural = (n, word) => `${fmtN(n)} ${word}${n === 1 ? '' : 's'}`;
const ymLabel = (k) => MON[(k % 100) - 1] + ' ' + Math.floor(k / 100);
const wkLabel = (k, multi) => 'W' + (k % 100) + (multi ? ' ’' + String(Math.floor(k / 100)).slice(2) : '');
// The calendar week today falls in, numbered like Excel's WEEKNUM(date, 2) (weeks start Monday,
// week 1 contains Jan 1). Uses the viewer's own clock, so it rolls over each Monday.
function thisWeek() {
  const now = new Date(), y = now.getFullYear(), today = new Date(y, now.getMonth(), now.getDate()), jan1 = new Date(y, 0, 1);
  const yday = Math.round((today - jan1) / 864e5) + 1, week = Math.floor((yday + (jan1.getDay() + 6) % 7 - 1) / 7) + 1;
  const mon = new Date(today); mon.setDate(today.getDate() - (today.getDay() + 6) % 7);
  const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
  return { key: y * 100 + week, week, year: y, mon, sun };
}
const nowTag = '<span class="nowtag">this week</span>';
const partTag = '<span class="parttag">partial week</span>';
// A week is incomplete when it is the week in progress (or later), or when New Year cuts it short
// (Excel's WEEKNUM splits Dec 29 - Jan 4 into W53 and W1). Such weeks are drawn faded or dashed.
function isPartial(k) {
  const y = Math.floor(k / 100), w = k % 100, jan1 = Date.UTC(y, 0, 1);
  const start = jan1 - ((new Date(jan1).getUTCDay() + 6) % 7) * DAY + 7 * (w - 1) * DAY;
  return start < jan1 || start + 6 * DAY > Date.UTC(y, 11, 31) || k >= thisWeek().key;
}
const partNote = (k) => (isPartial(k) ? `<div class="tnote">${k >= thisWeek().key ? 'Week in progress' : 'Partial week (split by New Year)'}</div>` : '');
const sortedKeys = (m) => [...m.keys()].sort((a, b) => a - b);
const multiYear = (keys) => new Set(keys.map((k) => Math.floor(k / 100))).size > 1;
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

// ---------- loaded data + view state ----------
let DATA = null, dims = null, NR = 0, months = [], S = null, A = null, RS = new Set();
// A-Z by name; placeholders such as "#N/A (not on CSR tab)" or "(blank)" go last.
const byName = (names) => (a, b) => {
  const x = String(names[a]), y = String(names[b]), px = !/^[a-z0-9]/i.test(x), py = !/^[a-z0-9]/i.test(y);
  return px !== py ? (px ? 1 : -1) : x.localeCompare(y, 'en', { sensitivity: 'base', numeric: true });
};

const defaults = () => ({ regions: Array.from({ length: NR }, (_, i) => i), account: -1, csr: -1, cust: -1,
  from: months[0], to: months[months.length - 1], tab: 'dash', otpMetric: 'gross', csrMetric: 'gross', fcH: 8, fcBasis: 12 });
const saveState = () => { try { localStorage.setItem(STATE_KEY, JSON.stringify(S)); } catch (e) {} };

function validate(d) {
  const ok = d && d.dims && Array.isArray(d.pod) && Array.isArray(d.ship) && Array.isArray(d.delay) &&
    ['region', 'cust', 'account', 'csr', 'delay'].every((k) => Array.isArray(d.dims[k]));
  if (!ok) throw new Error('This is not an OTP data file. Choose the region Excel files, or a data file saved from this page.');
  return d;
}

function load(d) {
  DATA = validate(d);
  dims = DATA.dims; NR = dims.region.length;
  const vol = new Map();
  for (const t of [DATA.pod, DATA.ship]) for (const r of t) { const k = r[4] * 100 + r[6]; vol.set(k, (vol.get(k) || 0) + r[7]); }
  // keep months with real volume, so typo dates (e.g. year 2326) don't stretch the range
  months = [...vol.entries()].filter(([, v]) => v >= 100).map(([k]) => k).sort((a, b) => a - b);
  if (!months.length) months = [...vol.keys()].sort((a, b) => a - b);
  S = defaults();
  try {
    const saved = JSON.parse(localStorage.getItem(STATE_KEY) || 'null');
    if (saved) Object.assign(S, saved, { from: months.includes(saved.from) ? saved.from : S.from, to: months.includes(saved.to) ? saved.to : S.to });
  } catch (e) {}
  S.regions = (Array.isArray(S.regions) ? S.regions : []).filter((i) => Number.isInteger(i) && i >= 0 && i < NR);
  if (!S.regions.length) S.regions = defaults().regions;

  show('app');
  updateBanner();
  const skipped = DATA.skipped || [];
  $('#notice').hidden = !skipped.length;
  $('#notice').textContent = skipped.length ? 'Skipped: ' + skipped.map((s) => `${s.file} ${s.reason}`).join('; ') + '.' : '';
  const nFiles = (DATA.sources || []).length, nRows = Object.values(DATA.checks || {}).reduce((s, v) => s + (v.rows || 0), 0);
  $('#meta').innerHTML = `${plural(nRows, 'HAWB')} from ${plural(nFiles, 'file')} · built ${esc(fmtBuilt(DATA.generated))}`;
  buildFilters(); aggregate(); renderMini(); setTab(S.tab);
}

// ---------- opening files ----------
function show(view) { for (const v of ['start', 'progress', 'app']) $('#' + v).hidden = v !== view; if (view !== 'app') { hideTip(); $('#miniBar').hidden = true; } }
function showStartError(msg) { const e = $('#startError'); e.textContent = msg; e.hidden = !msg; if (msg) show(DATA ? 'app' : 'start'); }
// Data lives only in this page's memory, so tell the viewer a refresh clears it.
function updateBanner() {
  $('#banner').innerHTML = '<span><b>Data is not saved.</b> Refreshing or closing this page clears it. Use Save data file to reopen it later without the Excel files.</span>';
}

function readJson(file) {
  const fr = new FileReader();
  fr.onload = () => {
    try { load(JSON.parse(fr.result)); showStartError(''); }
    catch (e) { DATA = null; showStartError(e instanceof SyntaxError ? 'That file could not be read as a data file.' : e.message); }
  };
  fr.onerror = () => showStartError('That file could not be opened.');
  fr.readAsText(file);
}

let worker = null;
function readExcel(files) {
  if (typeof DecompressionStream === 'undefined' || typeof Worker === 'undefined') {
    showStartError('This browser cannot read Excel files directly. Use a current version of Edge, Chrome, Firefox or Safari.');
    return;
  }
  showStartError('');
  const list = $('#progressList'), rows = new Map();
  list.innerHTML = '';
  for (const f of files) {
    const li = document.createElement('li');
    li.innerHTML = `<div class="pf"><b>${esc(f.name)}</b><span>Waiting</span></div><div class="bar"><i></i></div>`;
    list.append(li); rows.set(f.name, li);
  }
  const done = new Set(), count = () => { $('#progressCount').innerHTML = `<span class="otp-pcount">${done.size} of ${files.length} file${files.length === 1 ? '' : 's'} done.</span>`; };
  count();
  show('progress');
  if (worker) worker.terminate();
  worker = new Worker(new URL('worker.js', import.meta.url), { type: 'module' });
  worker.onmessage = (e) => {
    const m = e.data;
    if (m.type === 'progress') {
      const li = rows.get(m.file); if (!li) return;
      li.querySelector('span').textContent = m.rows != null ? `${fmtN(m.rows)} shipments` : m.stage;
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
  worker.postMessage({ files });
}

// One entry point for the file pickers and drag-and-drop: Excel files are read, a .json opens directly.
function openFiles(fileList) {
  const files = [...(fileList || [])];
  if (!files.length) return;
  const json = files.find((f) => /\.json$/i.test(f.name));
  if (json && files.length === 1) return readJson(json);
  readExcel(files.filter((f) => !/\.json$/i.test(f.name)));
}
$('#xlsxInput').addEventListener('change', (e) => { openFiles(e.target.files); e.target.value = ''; });
$('#fileInput').addEventListener('change', (e) => { openFiles(e.target.files); e.target.value = ''; });
$('#cancelBtn').addEventListener('click', () => { if (worker) { worker.terminate(); worker = null; } show(DATA ? 'app' : 'start'); });
let dragDepth = 0;
document.addEventListener('dragenter', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) { e.preventDefault(); dragDepth++; document.body.classList.add('dragging'); } });
document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dragging'); } });
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => { e.preventDefault(); dragDepth = 0; document.body.classList.remove('dragging'); if (!worker) openFiles(e.dataTransfer?.files); });

$('#saveBtn').addEventListener('click', () => {
  if (!DATA) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([JSON.stringify(DATA)], { type: 'application/json' }));
  a.download = 'otp_dashboard_data.json';
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
});
$('#forgetBtn').addEventListener('click', () => {
  DATA = null; A = null;
  show('start');
});

// ---------- filters ----------
function optionsFor(dimKey, col) {
  const rs = new Set(S.regions), seen = new Set();
  for (const t of [DATA.ship, DATA.pod]) for (const r of t) if (rs.has(r[0])) seen.add(r[col]);
  return [...seen].sort((a, b) => String(dims[dimKey][a]).localeCompare(String(dims[dimKey][b])));
}
function fillSelect(el, dimKey, col, cur, allLabel) {
  const ids = optionsFor(dimKey, col);
  el.innerHTML = `<option value="-1">${allLabel}</option>` + ids.map((i) => `<option value="${i}">${esc(dims[dimKey][i])}</option>`).join('');
  const keep = ids.includes(cur) ? cur : -1;
  el.value = String(keep);
  return keep;
}
function buildFilters() {
  $('#regionChips').innerHTML = `<button class="chip" type="button" data-r="all" aria-pressed="${S.regions.length === NR}">All</button>` +
    dims.region.map((n, i) => `<button class="chip" type="button" data-r="${i}" aria-pressed="${S.regions.length < NR && S.regions.includes(i)}">${esc(n)}</button>`).join('');
  S.account = fillSelect($('#fAccount'), 'account', 2, S.account, 'All accounts');
  S.csr = fillSelect($('#fCsr'), 'csr', 3, S.csr, 'All CSRs');
  S.cust = fillSelect($('#fCust'), 'cust', 1, S.cust, 'All customers');
  const mo = months.map((k) => `<option value="${k}">${ymLabel(k)}</option>`).join('');
  $('#fFrom').innerHTML = mo; $('#fTo').innerHTML = mo;
  $('#fFrom').value = S.from; $('#fTo').value = S.to;
}
$('#regionChips').addEventListener('click', (e) => {
  const b = e.target.closest('.chip'); if (!b) return;
  const all = Array.from({ length: NR }, (_, i) => i);
  if (b.dataset.r === 'all') S.regions = all;
  else {
    const i = +b.dataset.r, set = new Set(S.regions.length === NR ? [] : S.regions);
    set.has(i) ? set.delete(i) : set.add(i);
    S.regions = set.size ? [...set].sort((a, b) => a - b) : all;
  }
  buildFilters(); update();
});
for (const [id, key] of [['#fAccount', 'account'], ['#fCsr', 'csr'], ['#fCust', 'cust']])
  $(id).addEventListener('change', (e) => { S[key] = +e.target.value; update(); });
$('#fFrom').addEventListener('change', (e) => { S.from = +e.target.value; if (S.from > S.to) { S.to = S.from; $('#fTo').value = S.to; } update(); });
$('#fTo').addEventListener('change', (e) => { S.to = +e.target.value; if (S.to < S.from) { S.from = S.to; $('#fFrom').value = S.from; } update(); });
$('#resetBtn').addEventListener('click', () => { const t = S.tab; S = defaults(); S.tab = t; buildFilters(); update(); });

// One-line summary of the filters, shown in a slim bar once the filter panel scrolls off screen.
function renderMini() {
  if (!DATA) return;
  const parts = [S.regions.length === NR ? 'All regions' : S.regions.map((i) => dims.region[i]).join(', ')];
  if (S.account >= 0) parts.push(dims.account[S.account]);
  if (S.csr >= 0) parts.push('CSR: ' + dims.csr[S.csr]);
  if (S.cust >= 0) parts.push(dims.cust[S.cust]);
  parts.push(S.from === S.to ? ymLabel(S.from) : `${ymLabel(S.from)} – ${ymLabel(S.to)}`);
  $('#miniSummary').textContent = parts.join(' · ');
}
$('#miniEdit').addEventListener('click', () => {
  const top = $('.otp-filters').getBoundingClientRect().top + window.scrollY - 64;
  window.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
});
if (typeof IntersectionObserver !== 'undefined') {
  new IntersectionObserver(([e]) => { $('#miniBar').hidden = e.isIntersecting || !DATA || $('#app').hidden; },
    { rootMargin: '-48px 0px 0px 0px' }).observe($('.otp-filters'));
}

// ---------- tabs ----------
function setTab(t) {
  S.tab = ['dash', 'otp', 'ship', 'csr', 'fc', 'about'].includes(t) ? t : 'dash';
  document.querySelectorAll('.tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === S.tab)));
  document.querySelectorAll('.panel').forEach((p) => (p.hidden = p.id !== 'p-' + S.tab));
  saveState(); render();
}
$('.otp-tabs').addEventListener('click', (e) => { const b = e.target.closest('.tab'); if (b) setTab(b.dataset.tab); });

// ---------- aggregation ----------
// fact rows: pod   [region, cust, account, csr, year, week, month, hawb, grossLate, netLate, revenue]
//            ship  [region, cust, account, csr, year, week, month, hawb]
//            delay [region, cust, account, csr, year, week, month, delay, controllable, count]
const pass = (r) => {
  const ym = r[4] * 100 + r[6];
  return ym >= S.from && ym <= S.to && RS.has(r[0]) && (S.account < 0 || r[2] === S.account) && (S.csr < 0 || r[3] === S.csr) && (S.cust < 0 || r[1] === S.cust);
};
function aggregate() {
  RS = new Set(S.regions);
  const podW = new Map(), shipW = new Map(), delay = new Map(), regP = new Map(), regS = new Map();
  const tot = { h: 0, gl: 0, nl: 0, rev: 0, shipped: 0 }, podRows = [], shipRows = [];
  const bump = (m, k) => { let o = m.get(k); if (!o) m.set(k, (o = { h: 0, gl: 0, nl: 0, rev: 0 })); return o; };
  for (const r of DATA.pod) {
    if (!pass(r)) continue;
    podRows.push(r);
    for (const o of [bump(podW, r[4] * 100 + r[5]), bump(regP, r[0]), tot]) { o.h += r[7]; o.gl += r[8]; o.nl += r[9]; o.rev += r[10]; }
  }
  for (const r of DATA.ship) {
    if (!pass(r)) continue;
    shipRows.push(r);
    const wk = r[4] * 100 + r[5];
    shipW.set(wk, (shipW.get(wk) || 0) + r[7]);
    regS.set(r[0], (regS.get(r[0]) || 0) + r[7]);
    tot.shipped += r[7];
  }
  for (const r of DATA.delay) {
    if (!pass(r)) continue;
    let d = delay.get(r[7]); if (!d) delay.set(r[7], (d = [0, 0]));
    d[r[8] ? 0 : 1] += r[9];
  }
  A = { podW, shipW, delay, regP, regS, tot, podRows, shipRows };
}

// ---------- tooltip ----------
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
const row = (k, v, color) => `<div class="r"><span>${color ? `<i class="k" style="background:${color}"></i>` : ''}${k}</span><span>${v}</span></div>`;
const emptyMsg = (t) => `<div class="empty-s">${t}</div>`;

// ---------- SVG charts ----------
const NS = 'http://www.w3.org/2000/svg';
function svgEl(tag, attrs, parent) { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; }
function niceMax(v) { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))); for (const m of [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) if (m * p >= v) return m * p; return 10 * p; }
const topRoundedBar = (x, y, w, h, r) => { r = Math.min(r, w / 2, h); return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`; };
const rightRoundedBar = (x, y, w, h, r) => { r = Math.min(r, h / 2, w); return `M${x},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h - r}Q${x + w},${y + h} ${x + w - r},${y + h}H${x}Z`; };
const xTickEvery = (n, width) => Math.max(1, Math.ceil(n / Math.max(1, Math.floor(width / 44))));
const T = () => ({ ink: css('--text'), ink2: css('--o-ink2'), muted: css('--muted'), grid: css('--o-grid'), axis: css('--o-axis'), card: css('--o-card'), s1: css('--s1'), s2: css('--s2'), s3: css('--s3') });

function lineChart(el, keys, series, tipFn) {
  el.innerHTML = '';
  if (!keys.length) { el.innerHTML = emptyMsg('No delivered shipments for these filters.'); return; }
  const c = T(), W = el.clientWidth || 600, H = 260, m = { l: 42, r: 62, t: 14, b: 28 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = keys.length, bw = iw / n, multi = multiYear(keys);
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'Weekly gross and net on-time percentage' }, el);
  const X = (i) => m.l + (i + 0.5) * bw, Y = (v) => m.t + ih * (1 - v);
  for (const v of [0, .25, .5, .75, 1]) {
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: v === 0 ? c.axis : c.grid, 'stroke-width': 1 }, svg);
    svgEl('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = v * 100 + '%';
  }
  const every = xTickEvery(n, iw), part = keys.map(isPartial);
  keys.forEach((k, i) => { if (i % every === 0) svgEl('text', { x: X(i), y: H - 8, 'text-anchor': 'middle', 'font-size': 12, fill: c.muted }, svg).textContent = wkLabel(k, multi); });
  for (const s of series) {
    // solid between complete weeks; dashed and lighter on any segment touching an incomplete week
    let solid = '', dashed = '';
    for (let i = 0; i < n; i++) {
      const v = s.values[i]; if (v == null) continue;
      const pv = i > 0 ? s.values[i - 1] : null, pt = X(i).toFixed(1) + ',' + Y(v).toFixed(1);
      if (pv == null) { solid += 'M' + pt; continue; }
      const seg = 'M' + X(i - 1).toFixed(1) + ',' + Y(pv).toFixed(1) + 'L' + pt;
      if (part[i] || part[i - 1]) dashed += seg; else solid += seg;
    }
    svgEl('path', { d: solid, fill: 'none', stroke: s.color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }, svg);
    if (dashed) svgEl('path', { d: dashed, fill: 'none', stroke: s.color, 'stroke-width': 2, 'stroke-dasharray': '4 4', opacity: 0.6, 'stroke-linecap': 'round' }, svg);
  }
  const ends = series.map((s) => { const i = s.values.length - 1; return { s, v: s.values[i], y: Y(s.values[i] ?? 0), x: X(i), part: part[i] }; });
  const collide = ends.length > 1 && Math.abs(ends[0].y - ends[1].y) < 16;
  for (const e of ends) {
    if (e.v == null) continue;
    svgEl('circle', { cx: e.x, cy: e.y, r: 4, fill: e.part ? c.card : e.s.color, stroke: e.part ? e.s.color : c.card, 'stroke-width': 2 }, svg);
    if (!collide) svgEl('text', { x: e.x + 8, y: e.y + 4, 'font-size': 12, 'font-weight': 600, fill: c.ink }, svg).textContent = fmtP(e.v, e.s.dec);
  }
  const guide = svgEl('line', { y1: m.t, y2: m.t + ih, stroke: c.axis, 'stroke-width': 1, visibility: 'hidden' }, svg);
  const dots = series.map((s) => svgEl('circle', { r: 4, fill: s.color, stroke: c.card, 'stroke-width': 2, visibility: 'hidden' }, svg));
  const hit = svgEl('rect', { x: m.l, y: m.t, width: iw, height: ih, fill: 'transparent' }, svg);
  hit.addEventListener('mousemove', (ev) => {
    const rect = svg.getBoundingClientRect(), sx = (ev.clientX - rect.left) * (W / rect.width);
    const i = Math.max(0, Math.min(n - 1, Math.floor((sx - m.l) / bw)));
    guide.setAttribute('x1', X(i)); guide.setAttribute('x2', X(i)); guide.setAttribute('visibility', 'visible');
    series.forEach((s, j) => { const v = s.values[i]; dots[j].setAttribute('visibility', v == null ? 'hidden' : 'visible'); if (v != null) { dots[j].setAttribute('cx', X(i)); dots[j].setAttribute('cy', Y(v)); } });
    showTip(tipFn(i), ev);
  });
  hit.addEventListener('mouseleave', () => { guide.setAttribute('visibility', 'hidden'); dots.forEach((d) => d.setAttribute('visibility', 'hidden')); hideTip(); });
}

function columnChart(el, keys, values, tipFn) {
  el.innerHTML = '';
  if (!keys.length) { el.innerHTML = emptyMsg('No shipments for these filters.'); return; }
  const c = T(), W = el.clientWidth || 600, H = 260, m = { l: 50, r: 12, t: 14, b: 28 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = keys.length, bw = iw / n, multi = multiYear(keys);
  const max = niceMax(Math.max(...values));
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'HAWBs shipped per week' }, el);
  const Y = (v) => m.t + ih * (1 - v / max);
  for (let t = 0; t <= 4; t++) {
    const v = (max / 4) * t;
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: t === 0 ? c.axis : c.grid, 'stroke-width': 1 }, svg);
    svgEl('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = v >= 1000 ? (v / 1000).toFixed(v % 1000 ? 1 : 0) + 'K' : fmtN(v);
  }
  const every = xTickEvery(n, iw), bar = Math.max(2, Math.min(24, bw - 2)), bars = [], base = keys.map((k) => (isPartial(k) ? 0.35 : 1));
  keys.forEach((k, i) => {
    const x = m.l + i * bw + (bw - bar) / 2, h = ih * (values[i] / max);
    bars.push(svgEl('path', { d: topRoundedBar(x, Y(values[i]), bar, Math.max(h, 0.5), 4), fill: c.s1, opacity: base[i] }, svg));
    if (i % every === 0) svgEl('text', { x: m.l + (i + 0.5) * bw, y: H - 8, 'text-anchor': 'middle', 'font-size': 12, fill: c.muted }, svg).textContent = wkLabel(k, multi);
  });
  const hit = svgEl('rect', { x: m.l, y: m.t, width: iw, height: ih, fill: 'transparent' }, svg);
  hit.addEventListener('mousemove', (ev) => {
    const rect = svg.getBoundingClientRect(), sx = (ev.clientX - rect.left) * (W / rect.width);
    const i = Math.max(0, Math.min(n - 1, Math.floor((sx - m.l) / bw)));
    bars.forEach((b, j) => b.setAttribute('opacity', j === i ? base[j] : base[j] * 0.55));
    showTip(tipFn(i), ev);
  });
  hit.addEventListener('mouseleave', () => { bars.forEach((b, j) => b.setAttribute('opacity', base[j])); hideTip(); });
}

function delayChart(el) {
  el.innerHTML = '';
  const items = [...A.delay.entries()].map(([k, v]) => ({ label: dims.delay[k], c: v[0], u: v[1], t: v[0] + v[1] })).sort((a, b) => b.t - a.t);
  const total = items.reduce((s, x) => s + x.t, 0);
  if (!total) { el.innerHTML = emptyMsg('No late shipments for these filters.'); return; }
  const TOP = 11, shown = items.slice(0, TOP);
  if (items.length > TOP) { const rest = items.slice(TOP); shown.push({ label: `Other (${rest.length} codes)`, c: rest.reduce((s, x) => s + x.c, 0), u: rest.reduce((s, x) => s + x.u, 0), t: rest.reduce((s, x) => s + x.t, 0) }); }
  const c = T(), W = el.clientWidth || 600, rowH = 24, lw = Math.min(200, Math.max(110, W * 0.36)), mt = 4, mr = 52;
  const H = mt + shown.length * rowH + 4, iw = W - lw - mr, max = shown[0].t / total, maxChars = Math.floor(lw / 6.4);
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'Share of late shipments by delay code' }, el);
  shown.forEach((it, i) => {
    const y = mt + i * rowH, bh = 14, by = y + (rowH - bh) / 2;
    svgEl('text', { x: lw - 8, y: by + 11, 'text-anchor': 'end', 'font-size': 12, fill: c.ink2 }, svg).textContent = it.label.length > maxChars ? it.label.slice(0, maxChars - 1) + '…' : it.label;
    const wC = iw * (it.c / total) / max, wU = iw * (it.u / total) / max;
    if (it.c && it.u) {
      svgEl('rect', { x: lw, y: by, width: Math.max(0, wC - 1), height: bh, fill: c.s2 }, svg);
      svgEl('path', { d: rightRoundedBar(lw + wC + 1, by, Math.max(1, wU - 1), bh, 4), fill: c.s1 }, svg);
    } else {
      svgEl('path', { d: rightRoundedBar(lw, by, Math.max(1.5, wC + wU), bh, 4), fill: it.c ? c.s2 : c.s1 }, svg);
    }
    svgEl('text', { x: lw + wC + wU + 6, y: by + 11, 'font-size': 12, 'font-weight': 600, fill: c.ink }, svg).textContent = fmtP(it.t / total);
    const hit = svgEl('rect', { x: 0, y, width: W, height: rowH, fill: 'transparent' }, svg);
    hit.addEventListener('mousemove', (ev) => showTip(`<div class="tt">${esc(it.label)}</div>` + row('Late HAWBs', fmtN(it.t)) + row('Share of late', fmtP(it.t / total)) + row('Controllable', fmtN(it.c), c.s2) + row('Uncontrollable', fmtN(it.u), c.s1), ev));
    hit.addEventListener('mouseleave', hideTip);
  });
}

// ---------- tables ----------
function pctCell(p, h, metric) {
  const late = 1 - p;
  const a = (metric === 'net' ? Math.min(1, late / 0.05) : Math.min(1, late)) * 0.5;
  return `<td class="${h < 5 ? 'thin' : ''}" style="background-image:linear-gradient(rgb(var(--crit-rgb) / ${a.toFixed(3)}),rgb(var(--crit-rgb) / ${a.toFixed(3)}))">${fmtP(p, metric === 'net' ? 2 : 1)}</td>`;
}

function weekTable() {
  const keys = sortedKeys(A.podW), multi = multiYear(keys), t = A.tot;
  if (!keys.length) { $('#tWeek').innerHTML = emptyMsg('No delivered shipments for these filters.'); return; }
  const line = (label, w) => `<td>${label}</td><td>${fmtN(w.h)}</td><td>${fmt$(w.rev)}</td><td>${fmtN(w.gl)}</td><td>${fmtN(w.nl)}</td><td>${fmtP(1 - w.gl / w.h)}</td><td>${fmtP(1 - w.nl / w.h, 2)}</td>`;
  $('#tWeek').innerHTML = `<table><thead><tr><th>POD week</th><th>HAWBs</th><th>Total revenue</th><th>Gross late</th><th>Net late</th><th>On-time gross %</th><th>On-time net %</th></tr></thead><tbody>` +
    keys.slice().reverse().map((k) => k === thisWeek().key ? `<tr class="now">${line(wkLabel(k, multi) + nowTag, A.podW.get(k))}</tr>`
      : `<tr>${line(wkLabel(k, multi) + (isPartial(k) ? partTag : ''), A.podW.get(k))}</tr>`).join('') +
    `<tr class="total">${line('Grand total', t)}</tr></tbody></table>`;
}

function regionTable() {
  const ids = [...new Set([...A.regP.keys(), ...A.regS.keys()])].sort((a, b) => a - b);
  if (!ids.length) { $('#tRegion').innerHTML = emptyMsg('No data for these filters.'); return; }
  const line = (label, shipped, p) => {
    const g = p.h ? 1 - p.gl / p.h : null;
    return `<td>${label}</td><td>${fmtN(shipped)}</td><td>${fmtN(p.h)}</td><td>${fmt$c(p.rev)}</td><td>${fmtP(g)}<span class="meter"><b style="width:${((g || 0) * 100).toFixed(1)}%"></b></span></td><td>${p.h ? fmtP(1 - p.nl / p.h, 2) : '–'}</td>`;
  };
  $('#tRegion').innerHTML = `<table><thead><tr><th>Region</th><th>Shipped</th><th>Delivered</th><th>Revenue</th><th>Gross OTP</th><th>Net OTP</th></tr></thead><tbody>` +
    ids.map((i) => `<tr>${line(esc(dims.region[i]), A.regS.get(i) || 0, A.regP.get(i) || { h: 0, gl: 0, nl: 0, rev: 0 })}</tr>`).join('') +
    `<tr class="total">${line('Total', A.tot.shipped, A.tot)}</tr></tbody></table>`;
}

// rows grouped by fact column `col` (1 = customer, 3 = CSR), columns = weeks
function matrix(target, moreTarget, facts, col, mode, opts = {}) {
  const dimKey = col === 1 ? 'cust' : 'csr';
  const cells = new Map(), rowsT = new Map(), colT = new Map(), regionOf = new Map();
  const bump = (m, k) => { let o = m.get(k); if (!o) m.set(k, (o = { h: 0, gl: 0, nl: 0 })); return o; };
  for (const r of facts) {
    const wk = r[4] * 100 + r[5], key = r[col];
    regionOf.set(key, regionOf.has(key) && regionOf.get(key) !== r[0] ? -1 : r[0]);
    for (const o of [bump(cells, key + '|' + wk), bump(rowsT, key), bump(colT, wk)]) { o.h += r[7]; if (mode !== 'count') { o.gl += r[8]; o.nl += r[9]; } }
  }
  const el = $(target);
  if (!rowsT.size) { el.innerHTML = emptyMsg('No data for these filters.'); if (moreTarget) $(moreTarget).innerHTML = ''; return; }
  const weeks = sortedKeys(colT), multi = multiYear(weeks), q = (opts.search || '').trim().toLowerCase();
  const keys = [...rowsT.keys()].filter((k) => !q || String(dims[dimKey][k]).toLowerCase().includes(q)).sort(byName(dims[dimKey]));
  const showRegion = RS.size > 1 && col === 1, metric = mode === 'count' ? 'hawb' : opts.metric;
  const cell = (c, bold) => {
    if (!c || !c.h) return '<td></td>';
    const html = metric === 'hawb' ? `<td>${fmtN(c.h)}</td>` : pctCell(1 - (metric === 'net' ? c.nl : c.gl) / c.h, c.h, metric);
    return bold ? html.replace('<td', '<td style="font-weight:600"') : html;
  };
  const all = { h: 0, gl: 0, nl: 0 }; for (const v of colT.values()) { all.h += v.h; all.gl += v.gl; all.nl += v.nl; }
  const tot = (html) => html.replace(/^<td( class="([^"]*)")?/, (m0, a, cls) => `<td class="tot${cls ? ' ' + cls : ''}"`);   // frozen Total column
  const wkHead = (w) => w === thisWeek().key ? `<th class="now" title="This week">${wkLabel(w, multi)}</th>`
    : isPartial(w) ? `<th class="part" title="Partial week">${wkLabel(w, multi)}</th>` : `<th>${wkLabel(w, multi)}</th>`;
  let h = `<table class="mx"><thead><tr><th>${dimKey === 'cust' ? 'Customer' : 'CSR'}</th><th class="tot">Total</th>${weeks.map(wkHead).join('')}</tr></thead><tbody>`;
  for (const k of keys) {
    const reg = regionOf.get(k), name = esc(dims[dimKey][k]);
    h += `<tr data-k="${k}"><td title="${name}">${name}${showRegion && reg >= 0 ? `<span class="rtag">${esc(dims.region[reg])}</span>` : ''}</td>` +
      tot(cell(rowsT.get(k))) + weeks.map((w) => cell(cells.get(k + '|' + w))).join('') + '</tr>';
  }
  h += `<tr class="total"><td>Grand total</td>${tot(cell(all))}${weeks.map((w) => cell(colT.get(w))).join('')}</tr></tbody></table>`;
  el.innerHTML = h;
  el.scrollLeft = el.scrollWidth;   // open at the newest week
  el.onmousemove = (ev) => {
    const td = ev.target.closest('td'); if (!td || td.cellIndex === 0) { hideTip(); return; }
    const tr = td.parentElement, wi = td.cellIndex - 2, wk = wi >= 0 ? weeks[wi] : null;
    let c, title;
    if (tr.classList.contains('total')) { c = wk == null ? all : colT.get(wk); title = 'All ' + (dimKey === 'cust' ? 'customers' : 'CSRs'); }
    else { const k = +tr.dataset.k; c = wk == null ? rowsT.get(k) : cells.get(k + '|' + wk); title = dims[dimKey][k]; }
    if (!c || !c.h) { hideTip(); return; }
    let body = `<div class="tt">${esc(title)}</div>` + row(wk == null ? 'Period' : (mode === 'count' ? 'Ship week' : 'POD week'), wk == null ? 'Total' : wkLabel(wk, true)) + row('HAWBs', fmtN(c.h));
    if (mode !== 'count') body += row('Gross late', fmtN(c.gl)) + row('Net late', fmtN(c.nl)) + row('On-time gross', fmtP(1 - c.gl / c.h)) + row('On-time net', fmtP(1 - c.nl / c.h, 2));
    if (wk != null) body += partNote(wk);
    showTip(body, ev);
  };
  el.onmouseleave = hideTip;
  if (moreTarget) {
    const mt = $(moreTarget);
    const noun = dimKey === 'cust' ? 'customer' : 'CSR';
    mt.innerHTML = `<span class="note">${keys.length} ${noun}${keys.length === 1 ? '' : 's'}${q ? ' matching your search' : ''}, A to Z.</span>`;
  }
}

function segControl(sel, key) {
  const el = $(sel);
  el.innerHTML = [['hawb', 'HAWBs'], ['gross', 'Gross %'], ['net', 'Net %']].map(([v, l]) => `<button type="button" data-v="${v}" aria-pressed="${S[key] === v}">${l}</button>`).join('');
  el.onclick = (e) => { const b = e.target.closest('button'); if (!b) return; S[key] = b.dataset.v; saveState(); render(); };
}

function csrTable() {
  const m = new Map();
  const get = (k) => { let o = m.get(k); if (!o) m.set(k, (o = { s: 0, h: 0, gl: 0, nl: 0, cust: new Set() })); return o; };
  for (const r of A.shipRows) { const o = get(r[3]); o.s += r[7]; o.cust.add(r[1]); }
  for (const r of A.podRows) { const o = get(r[3]); o.h += r[7]; o.gl += r[8]; o.nl += r[9]; o.cust.add(r[1]); }
  const keys = [...m.keys()].sort(byName(dims.csr));
  if (!keys.length) { $('#tCsr').innerHTML = emptyMsg('No data for these filters.'); return; }
  $('#tCsr').innerHTML = `<table><thead><tr><th>CSR</th><th>Customers</th><th>Shipped</th><th>Delivered</th><th>Gross late</th><th>Net late</th><th>On-time gross</th><th>On-time net</th></tr></thead><tbody>` +
    keys.map((k) => { const o = m.get(k); return `<tr><td>${esc(dims.csr[k])}</td><td>${o.cust.size}</td><td>${fmtN(o.s)}</td><td>${fmtN(o.h)}</td><td>${fmtN(o.gl)}</td><td>${fmtN(o.nl)}</td><td>${fmtP(o.h ? 1 - o.gl / o.h : null)}</td><td>${fmtP(o.h ? 1 - o.nl / o.h : null, 2)}</td></tr>`; }).join('') +
    '</tbody></table>';
}

function tiles() {
  const t = A.tot, g = t.h ? 1 - t.gl / t.h : null, n = t.h ? 1 - t.nl / t.h : null;
  const tile = (label, v, d) => `<div class="ktile"><div class="eyebrow">${label}</div><div class="v">${v}</div>${d ? `<div class="d">${d}</div>` : ''}</div>`;
  $('#tiles').innerHTML =
    tile('HAWBs shipped', fmtN(t.shipped), 'by ship date in period') +
    tile('HAWBs delivered', fmtN(t.h), 'with a POD in period') +
    tile('Total revenue', fmt$c(t.rev), 'on delivered HAWBs') +
    tile('On-time gross', fmtP(g), `${fmtN(t.gl)} gross late`) +
    tile('On-time net', fmtP(n, 2), `<span class="pill ${t.nl ? 'bad' : 'ok'}">${fmtN(t.nl)} controllable late</span>`) +
    tile('Late, uncontrollable', fmtN(t.gl - t.nl), t.gl ? fmtP((t.gl - t.nl) / t.gl) + ' of gross late' : '');
}

function otpChart() {
  const keys = sortedKeys(A.podW), c = T();
  const gross = keys.map((k) => { const w = A.podW.get(k); return 1 - w.gl / w.h; });
  const net = keys.map((k) => { const w = A.podW.get(k); return 1 - w.nl / w.h; });
  lineChart($('#cOtp'), keys, [{ values: gross, color: c.s1, dec: 1 }, { values: net, color: c.s3, dec: 2 }], (i) => {
    const w = A.podW.get(keys[i]);
    return `<div class="tt">POD ${wkLabel(keys[i], true)}</div>` + row('HAWBs', fmtN(w.h)) + row('Gross on-time', fmtP(gross[i]), c.s1) + row('Net on-time', fmtP(net[i], 2), c.s3) + row('Gross late', fmtN(w.gl)) + row('Net late', fmtN(w.nl)) + partNote(keys[i]);
  });
}
function shipChart(sel) {
  const keys = sortedKeys(A.shipW), vals = keys.map((k) => A.shipW.get(k));
  columnChart($(sel), keys, vals, (i) => `<div class="tt">Ship ${wkLabel(keys[i], true)}</div>` + row('HAWBs shipped', fmtN(vals[i])) + partNote(keys[i]));
}

function about() {
  const checks = Object.entries(DATA.checks || {}).map(([r, v]) => {
    const mm = Object.entries(v.mismatches || {}).map(([k, n]) => `${fmtN(n)} ${esc(k)}`).join(', ');
    return `<li><b>${esc(r)}</b>: ${fmtN(v.rows || 0)} rows. ${mm ? `Differs from Excel's saved values on ${mm} rows; the recalculated value is used.` : "Matches Excel's saved values on every row."}</li>`;
  }).join('');
  $('#about').innerHTML = `
    <h2>How the numbers are calculated</h2>
    <p>The page reads each region file and recalculates the formula columns the same way the workbooks do:</p>
    <ul>
      <li><b>SLA adjusted</b> = first digit of <code>Transit Time</code>, +1 if the destination zone is not A–E, +1 if the origin zone is not A–E.</li>
      <li><b>Due date adjusted</b> = <code>WORKDAY(Ship Date, SLA adjusted, holidays)</code> using the holiday list on each file's CSR tab.</li>
      <li><b>On-time</b> when the POD is on or before the adjusted due date. A POD that carries a time of day on the due date counts as late, as it does in Excel.</li>
      <li><b>Gross late</b> = every late HAWB. <b>Net late</b> = late HAWBs whose defect code is <i>Controllable</i> in the “delay codes in WP” table.</li>
      <li><b>Weeks</b> follow <code>WEEKNUM(date, 2)</code> (weeks start Monday) and are kept apart by year, so W50 ’25 sorts before W1 ’26.</li>
      <li><b>CSR and Account</b> come from the value on each row, which keeps CSR history. Rows without a value use the CSR tab lookup.</li>
      <li>Delivery views (OTP, delay codes, revenue) filter by POD month; shipment counts filter by ship month, matching the Excel pivots.</li>
    </ul>
    <h2>How projections work</h2>
    <ul>
      <li><b>HAWBs and revenue</b> follow a straight-line trend through the basis weeks (the last 8, 12 or 26 full weeks). The trend eases off further out, so a short run of growth or decline is not carried on forever.</li>
      <li><b>On-time %</b> is the volume-weighted average of the basis weeks.</li>
      <li>The <b>likely range</b> comes from how far actual weeks strayed from that pattern; about 8 in 10 weeks should land inside it. It widens the further out you look.</li>
      <li>The newest week is left out when it has under 60% of the usual volume, since that means the export was pulled mid-week.</li>
      <li>There is less than a year of history, so seasonal peaks and holiday weeks are not built in. Treat projections as a guide for the next few weeks, not a budget.</li>
    </ul>
    ${checks ? `<h2>Source check</h2><ul>${checks}</ul>` : ''}
    <p class="note">Data built ${esc(fmtBuilt(DATA.generated))} from ${(DATA.sources || []).map(esc).join(', ')}.</p>`;
}

// ---------- projections ----------
// Weekly series ignore the month range (projections always start from the latest data) but use the
// other filters. Weeks are keyed by their Monday, so the W53 and W1 halves of the New Year week merge.
const FC_Z = 1.28;   // likely range covers about 8 in 10 weeks
const FC_DAMP = 0.8; // the trend eases off further out instead of running away
const mondayIdx = (y, w) => { const j = Date.UTC(y, 0, 1); return Math.round(((j - ((new Date(j).getUTCDay() + 6) % 7) * DAY + 7 * (w - 1) * DAY) / DAY - 4) / 7); };
const idxDate = (i) => new Date((i * 7 + 4) * DAY);
const idxLabel = (i) => { const d = idxDate(i); return MON[d.getUTCMonth()] + ' ' + d.getUTCDate(); };
const passNoMonth = (r, valid) => valid.has(r[4] * 100 + r[6]) && RS.has(r[0]) && (S.account < 0 || r[2] === S.account) && (S.csr < 0 || r[3] === S.csr) && (S.cust < 0 || r[1] === S.cust);

// Sums the given fact columns per Monday-week, filling missing weeks with zeros.
function weekly(table, cols) {
  const valid = new Set(months), m = new Map();
  for (const r of table) {
    if (!passNoMonth(r, valid)) continue;
    const k = mondayIdx(r[4], r[5]);
    let a = m.get(k); if (!a) m.set(k, (a = cols.map(() => 0)));
    cols.forEach((c, j) => (a[j] += r[c]));
  }
  if (!m.size) return [];
  const ks = [...m.keys()], lo = Math.min(...ks), hi = Math.max(...ks), out = [];
  for (let k = lo; k <= hi; k++) out.push({ k, v: m.get(k) || cols.map(() => 0) });
  return out;
}
// Drops the newest week when it is far below the weeks before it (an export pulled mid-week).
function trimPartial(series) {
  if (series.length < 5) return { series, dropped: null };
  const last = series[series.length - 1], prev = series.slice(-5, -1).map((p) => p.v[0]).sort((a, b) => a - b);
  const med = (prev[1] + prev[2]) / 2;
  return last.v[0] < 0.6 * med ? { series: series.slice(0, -1), dropped: last } : { series, dropped: null };
}
// Straight-line fit over the basis weeks, projected with a trend that eases off.
function trendProject(ys, H) {
  const n = ys.length, xm = (n - 1) / 2, ym = ys.reduce((s, y) => s + y, 0) / n;
  let sxy = 0, sxx = 0;
  ys.forEach((y, x) => { sxy += (x - xm) * (y - ym); sxx += (x - xm) ** 2; });
  const b = sxx ? sxy / sxx : 0, level = ym + b * (n - 1 - xm);
  const sse = ys.reduce((s, y, x) => s + (y - (ym + b * (x - xm))) ** 2, 0), sd = Math.sqrt(sse / Math.max(1, n - 2));
  const out = [];
  for (let h = 1; h <= H; h++) {
    const mid = Math.max(0, level + b * FC_DAMP * (1 - FC_DAMP ** h) / (1 - FC_DAMP)), w = FC_Z * sd * Math.sqrt(1 + h / n);
    out.push({ mid, lo: Math.max(0, mid - w), hi: mid + w });
  }
  return { out, slope: b, mean: ym };
}
// On-time %: volume-weighted average of the basis weeks, with a range from the week-to-week swing.
function rateProject(ok, tot, H) {
  const T = tot.reduce((s, v) => s + v, 0);
  if (!T) return null;
  const p = ok.reduce((s, v) => s + v, 0) / T;
  const rates = tot.map((t, i) => (t ? ok[i] / t : null)).filter((v) => v != null);
  const sd = Math.sqrt(rates.reduce((s, r) => s + (r - p) ** 2, 0) / Math.max(1, rates.length - 1));
  return Array.from({ length: H }, (_, h) => {
    const w = FC_Z * sd * Math.sqrt(1 + (h + 1) / rates.length);
    return { mid: p, lo: Math.max(0, p - w), hi: Math.min(1, p + w) };
  });
}

function fcAxisLabel(v, pct, money) {
  if (pct) return Math.round(v * 100) + '%';
  const pre = money ? '$' : '';
  if (v >= 1e6) return pre + (v / 1e6).toFixed(1) + 'M';
  if (v >= 1000) return pre + (v / 1000).toFixed(v % 1000 ? 1 : 0) + 'K';
  return pre + fmtN(v);
}

// Actual line, then a dashed projection with its likely-range band, in a shaded "Projected" zone.
function fcChart(el, keys, nActual, series, { pct = false, money = false, tip }) {
  el.innerHTML = '';
  const c = T(), W = el.clientWidth || 600, H = 260, m = { l: pct ? 42 : 56, r: 14, t: 20, b: 28 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = keys.length, bw = iw / n;
  let max = 1;
  if (!pct) { max = 0; for (const s of series) for (let i = 0; i < n; i++) max = Math.max(max, s.actual[i] ?? 0, s.hi[i] ?? 0); max = niceMax(max); }
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'Actual and projected weekly values' }, el);
  const X = (i) => m.l + (i + 0.5) * bw, Y = (v) => m.t + ih * (1 - v / max);
  const fx = X(nActual - 1);
  svgEl('rect', { x: fx, y: m.t, width: W - m.r - fx, height: ih, fill: c.grid, opacity: 0.45 }, svg);
  for (let t = 0; t <= 4; t++) {
    const v = (max / 4) * t;
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: t === 0 ? c.axis : c.grid, 'stroke-width': 1 }, svg);
    svgEl('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = fcAxisLabel(v, pct, money);
  }
  const every = xTickEvery(n, iw);
  keys.forEach((k, i) => { if (i % every === 0) svgEl('text', { x: X(i), y: H - 8, 'text-anchor': 'middle', 'font-size': 12, fill: c.muted }, svg).textContent = idxLabel(k); });
  svgEl('line', { x1: fx, x2: fx, y1: m.t - 8, y2: m.t + ih, stroke: c.axis, 'stroke-width': 1 }, svg);
  svgEl('text', { x: fx + 6, y: m.t - 6, 'font-size': 12, fill: c.muted }, svg).textContent = 'Projected';
  for (const s of series) {
    const col = c[s.color];
    let top = '', bottom = '';
    for (let i = nActual - 1; i < n; i++) { top += (top ? 'L' : 'M') + X(i).toFixed(1) + ',' + Y(s.hi[i]).toFixed(1); bottom = 'L' + X(i).toFixed(1) + ',' + Y(s.lo[i]).toFixed(1) + bottom; }
    svgEl('path', { d: top + bottom + 'Z', fill: col, opacity: 0.16 }, svg);
    let d = '', pen = false;
    for (let i = 0; i < nActual; i++) { const v = s.actual[i]; if (v == null) { pen = false; continue; } d += (pen ? 'L' : 'M') + X(i).toFixed(1) + ',' + Y(v).toFixed(1); pen = true; }
    svgEl('path', { d, fill: 'none', stroke: col, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }, svg);
    let p = '';
    for (let i = nActual - 1; i < n; i++) p += (p ? 'L' : 'M') + X(i).toFixed(1) + ',' + Y(s.mid[i]).toFixed(1);
    svgEl('path', { d: p, fill: 'none', stroke: col, 'stroke-width': 2, 'stroke-dasharray': '5 4', 'stroke-linecap': 'round' }, svg);
    svgEl('circle', { cx: X(n - 1), cy: Y(s.mid[n - 1]), r: 4, fill: col, stroke: c.card, 'stroke-width': 2 }, svg);
  }
  const guide = svgEl('line', { y1: m.t, y2: m.t + ih, stroke: c.axis, 'stroke-width': 1, visibility: 'hidden' }, svg);
  const hit = svgEl('rect', { x: m.l, y: m.t, width: iw, height: ih, fill: 'transparent' }, svg);
  hit.addEventListener('mousemove', (ev) => {
    const rect = svg.getBoundingClientRect(), sx = (ev.clientX - rect.left) * (W / rect.width);
    const i = Math.max(0, Math.min(n - 1, Math.floor((sx - m.l) / bw)));
    guide.setAttribute('x1', X(i)); guide.setAttribute('x2', X(i)); guide.setAttribute('visibility', 'visible');
    showTip(tip(i), ev);
  });
  hit.addEventListener('mouseleave', () => { guide.setAttribute('visibility', 'hidden'); hideTip(); });
}

function projections() {
  const H = S.fcH || 8, N = S.fcBasis || 12, c = T();
  [['#fcHorizon', 'fcH', [4, 8, 12]], ['#fcBasis', 'fcBasis', [8, 12, 26]]].forEach(([sel, key, opts]) => {
    const el = $(sel);
    el.innerHTML = opts.map((v) => `<button type="button" data-v="${v}" aria-pressed="${(S[key] || (key === 'fcH' ? 8 : 12)) === v}">${v} weeks</button>`).join('');
    el.onclick = (e) => { const b = e.target.closest('button'); if (!b) return; S[key] = +b.dataset.v; saveState(); render(); };
  });
  const clear = (msg) => { for (const s of ['#fcTiles', '#fcVol', '#fcOtp', '#fcRev', '#fcTable']) $(s).innerHTML = ''; $('#fcNote').textContent = msg; };
  const shipT = trimPartial(weekly(DATA.ship, [7])), podT = trimPartial(weekly(DATA.pod, [7, 8, 9, 10]));
  const ship = shipT.series, pod = podT.series;
  if (ship.length < 4 || pod.length < 4) { clear('Not enough weekly history for these filters. Projections need at least 4 full weeks.'); return; }
  const sb = ship.slice(-N), pb = pod.slice(-N);
  const vol = trendProject(sb.map((p) => p.v[0]), H), rev = trendProject(pb.map((p) => p.v[3]), H);
  const gross = rateProject(pb.map((p) => p.v[0] - p.v[1]), pb.map((p) => p.v[0]), H);
  const net = rateProject(pb.map((p) => p.v[0] - p.v[2]), pb.map((p) => p.v[0]), H);
  if (!gross) { clear('No delivered shipments in the basis weeks for these filters.'); return; }

  const notes = [`Based on ${sb.length} full ship weeks (week of ${idxLabel(sb[0].k)} to week of ${idxLabel(sb[sb.length - 1].k)}).`];
  if (shipT.dropped) notes.push(`The week of ${idxLabel(shipT.dropped.k)} is left out because it has only ${fmtN(shipT.dropped.v[0])} HAWBs so far, which looks like a partial week.`);
  if (sb.length < N) notes.push(`Only ${sb.length} weeks of history are available for these filters.`);
  notes.push('Weeks with a holiday usually come in lower than projected.');
  $('#fcNote').textContent = notes.join(' ');

  const sum = (a, f) => a.reduce((s, x) => s + x[f], 0);
  const trendPct = vol.mean ? vol.slope / vol.mean : 0;
  const tile = (label, v, d) => `<div class="ktile proj"><div class="eyebrow">Projected · ${label}</div><div class="v">${v}</div>${d ? `<div class="d">${d}</div>` : ''}</div>`;
  $('#fcTiles').innerHTML =
    tile(`HAWBs, next ${H} weeks`, fmtN(sum(vol.out, 'mid')), `likely ${fmtN(sum(vol.out, 'lo'))} – ${fmtN(sum(vol.out, 'hi'))}`) +
    tile('recent trend', (trendPct >= 0 ? '+' : '') + (trendPct * 100).toFixed(1) + '% a week', `${trendPct >= 0 ? 'more' : 'fewer'} HAWBs each week over the last ${sb.length} weeks`) +
    tile('gross OTP', fmtP(gross[0].mid), `likely ${fmtP(gross[0].lo)} – ${fmtP(gross[0].hi)} in a given week`) +
    tile('net OTP', fmtP(net[0].mid, 2), `likely ${fmtP(net[0].lo, 2)} – ${fmtP(net[0].hi, 2)} in a given week`) +
    tile(`revenue, next ${H} weeks`, fmt$c(sum(rev.out, 'mid')), `likely ${fmt$c(sum(rev.out, 'lo'))} – ${fmt$c(sum(rev.out, 'hi'))}`);

  // chart series: up to 26 actual weeks, then H projected weeks joined at the last actual week
  const build = (hist, val, proj) => {
    const a = hist.slice(-Math.min(26, hist.length)), last = a[a.length - 1];
    const keys = a.map((p) => p.k).concat(proj.map((_, h) => last.k + h + 1));
    const actual = a.map(val), lv = actual[actual.length - 1];
    const pad = (f) => Array(a.length - 1).fill(null).concat([lv], proj.map((p) => p[f]));
    return { keys, nActual: a.length, actual, mid: pad('mid'), lo: pad('lo'), hi: pad('hi') };
  };
  const head = (keys, i, isP) => `<div class="tt">Week of ${idxLabel(keys[i])}${isP ? ' · projected' : ''}</div>`;
  const v = build(ship, (p) => p.v[0], vol.out);
  fcChart($('#fcVol'), v.keys, v.nActual, [{ color: 's1', ...v }], { tip: (i) => {
    const isP = i >= v.nActual;
    return head(v.keys, i, isP) + (isP ? row('Projected HAWBs', fmtN(v.mid[i]), c.s1) + row('Likely range', `${fmtN(v.lo[i])} – ${fmtN(v.hi[i])}`) : row('HAWBs shipped', fmtN(v.actual[i]), c.s1));
  } });
  const rate = (f) => (p) => (p.v[0] ? 1 - p.v[f] / p.v[0] : null);
  const g = build(pod, rate(1), gross), nn = build(pod, rate(2), net);
  fcChart($('#fcOtp'), g.keys, g.nActual, [{ color: 's1', ...g }, { color: 's3', ...nn }], { pct: true, tip: (i) => {
    const isP = i >= g.nActual;
    return head(g.keys, i, isP) + (isP
      ? row('Gross on-time', `${fmtP(g.mid[i])} (${fmtP(g.lo[i])} – ${fmtP(g.hi[i])})`, c.s1) + row('Net on-time', `${fmtP(nn.mid[i], 2)} (${fmtP(nn.lo[i], 2)} – ${fmtP(nn.hi[i], 2)})`, c.s3)
      : row('Gross on-time', fmtP(g.actual[i]), c.s1) + row('Net on-time', fmtP(nn.actual[i], 2), c.s3));
  } });
  const r = build(pod, (p) => p.v[3], rev.out);
  fcChart($('#fcRev'), r.keys, r.nActual, [{ color: 's1', ...r }], { money: true, tip: (i) => {
    const isP = i >= r.nActual;
    return head(r.keys, i, isP) + (isP ? row('Projected revenue', fmt$(r.mid[i]), c.s1) + row('Likely range', `${fmt$(r.lo[i])} – ${fmt$(r.hi[i])}`) : row('Delivered revenue', fmt$(r.actual[i]), c.s1));
  } });

  const lastShip = ship[ship.length - 1].k;
  const tw = thisWeek(), isNowIdx = (i) => i === mondayIdx(tw.year, tw.week);
  const wkName = (h) => `${idxLabel(lastShip + h + 1)}, ${idxDate(lastShip + h + 1).getUTCFullYear()}`;
  const fig = (v, lo, hi) => `<td>${v}<span class="rg">${lo} – ${hi}</span></td>`;
  $('#fcTable').innerHTML = `<table><thead><tr><th>Week of</th><th>HAWBs shipped</th><th>Gross OTP</th><th>Net OTP</th><th>Revenue</th></tr></thead><tbody>` +
    vol.out.map((p, h) => `<tr${isNowIdx(lastShip + h + 1) ? ' class="now"' : ''}><td>${wkName(h)}${isNowIdx(lastShip + h + 1) ? nowTag : ''}</td>` +
      fig(fmtN(p.mid), fmtN(p.lo), fmtN(p.hi)) + fig(fmtP(gross[h].mid), fmtP(gross[h].lo), fmtP(gross[h].hi)) +
      fig(fmtP(net[h].mid, 2), fmtP(net[h].lo, 2), fmtP(net[h].hi, 2)) + fig(fmt$(rev.out[h].mid), fmt$(rev.out[h].lo), fmt$(rev.out[h].hi)) + '</tr>').join('') +
    `<tr class="total"><td>Total</td>${fig(fmtN(sum(vol.out, 'mid')), fmtN(sum(vol.out, 'lo')), fmtN(sum(vol.out, 'hi')))}<td></td><td></td>` +
    `${fig(fmt$(sum(rev.out, 'mid')), fmt$(sum(rev.out, 'lo')), fmt$(sum(rev.out, 'hi')))}</tr></tbody></table>`;
}

// ---------- render ----------
function renderThisWeek() {
  const w = thisWeek(), f = (d, yr) => MON[d.getMonth()] + ' ' + d.getDate() + (yr ? ', ' + d.getFullYear() : '');
  $('#thisWeek').innerHTML = `<b>Week ${w.week}</b><span>This week: ${f(w.mon, w.mon.getFullYear() !== w.sun.getFullYear())} – ${f(w.sun, true)}</span>`;
}
function render() {
  if (!A || $('#app').hidden) return;
  renderThisWeek();
  if (S.tab === 'dash') { tiles(); otpChart(); shipChart('#cShip'); delayChart($('#cDelay')); regionTable(); weekTable(); }
  if (S.tab === 'otp') { segControl('#otpMetric', 'otpMetric'); matrix('#mOtp', '#mOtpMore', A.podRows, 1, 'otp', { metric: S.otpMetric, search: $('#otpSearch').value }); }
  if (S.tab === 'ship') { shipChart('#cShip2'); matrix('#mShip', '#mShipMore', A.shipRows, 1, 'count', { search: $('#shipSearch').value }); }
  if (S.tab === 'csr') { csrTable(); matrix('#mCsrShip', null, A.shipRows, 3, 'count'); segControl('#csrMetric', 'csrMetric'); matrix('#mCsrOtp', null, A.podRows, 3, 'otp', { metric: S.csrMetric }); }
  if (S.tab === 'fc') projections();
  if (S.tab === 'about') about();
}
function update() { saveState(); aggregate(); renderMini(); render(); }
$('#otpSearch').addEventListener('input', render);
$('#shipSearch').addEventListener('input', render);
let rz, lastW = 0;
new ResizeObserver((en) => { const w = Math.round(en[0].contentRect.width); if (w === lastW) return; lastW = w; clearTimeout(rz); rz = setTimeout(render, 120); }).observe($('#app'));
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', render);
// If the page stays open past midnight on Sunday, move the week badge and highlights to the new week.
let shownWeek = thisWeek().key;
setInterval(() => { const k = thisWeek().key; if (k !== shownWeek) { shownWeek = k; render(); } }, 60000);
new MutationObserver(render).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });

// ---------- clear any copy saved by the removed "Remember on this computer" option ----------
try { localStorage.removeItem(OLD_DATA_KEY); } catch (e) {}
