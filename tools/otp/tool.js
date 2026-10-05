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
const nowTag = '<span class="nowtag">In-Progress</span>';
const partTag = '<span class="parttag">partial week</span>';
// A week is incomplete when it is the week in progress (or later), or when New Year cuts it short
// (Excel's WEEKNUM splits Dec 29 - Jan 4 into W53 and W1). Such weeks are drawn faded or dashed.
function isPartial(k) {
  const y = Math.floor(k / 100), w = k % 100, jan1 = Date.UTC(y, 0, 1);
  const start = jan1 - ((new Date(jan1).getUTCDay() + 6) % 7) * DAY + 7 * (w - 1) * DAY;
  return start < jan1 || start + 6 * DAY > Date.UTC(y, 11, 31) || k >= thisWeek().key;
}
const partNote = (k) => (isPartial(k) ? `<div class="tnote">${k >= thisWeek().key ? 'In-Progress' : 'Partial week (split by New Year)'}</div>` : '');
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
  from: months[0], to: months[months.length - 1], tab: 'dash', otpMetric: 'gross', csrMetric: 'gross', csrPer: 'week', fcH: 8, fcBasis: 12, revOn: true });
// Revenue can be hidden everywhere (screens and exports), e.g. when sharing the screen.
const showRev = () => S.revOn !== false;
function revUi() { $('#revBtn').textContent = showRev() ? 'Hide revenue' : 'Show revenue'; }
const saveState = () => { try { localStorage.setItem(STATE_KEY, JSON.stringify(S)); } catch (e) {} };

function validate(d) {
  const ok = d && d.dims && Array.isArray(d.pod) && Array.isArray(d.ship) && Array.isArray(d.delay) &&
    ['region', 'cust', 'account', 'csr', 'delay'].every((k) => Array.isArray(d.dims[k]));
  if (!ok) throw new Error('This is not an OTP data file. Choose the region Excel files, or a data file saved from this page.');
  return d;
}

// Region names as the team uses them. Applied on load, so saved data files and the Python
// script's output show the same names.
const REGION_NAMES = { 'COSTCO': 'COSTCO+', 'APPLE/DELL': 'APPLE/DELL+' };
const regionLabel = (n) => REGION_NAMES[n] || n;
// These regions number weeks the ISO way, like their OTP summary workbooks: week 1 is the full
// Monday-Sunday week containing Jan 1 (in 2026, Dec 29 - Jan 4) instead of Excel's split W53/W1.
const ISO_WEEK_REGIONS = ['COSTCO+', 'AMAZON'];
// CSR volume split while the CSRs shared accounts: for ship weeks before 2026 W41, 1/3 of Jesus Quiroga's
// COSTCO+ shipments count for Mindy Wilson. Applied to CSR shipment volume only.
const CSR_SPLITS = [{ region: 'COSTCO+', from: 'Jesus Quiroga', to: 'Mindy Wilson', share: 1 / 3, beforeWeek: 202641 }];
function isoWeekKey(y, w) {   // Excel WEEKNUM(,2) year/week -> ISO year*100+week (weeks start Monday in both)
  const jan1 = Date.UTC(y, 0, 1), start = jan1 - ((new Date(jan1).getUTCDay() + 6) % 7) * DAY + 7 * (w - 1) * DAY;
  const thu = new Date(start + 3 * DAY), iy = thu.getUTCFullYear();
  return iy * 100 + Math.floor((thu - Date.UTC(iy, 0, 1)) / DAY / 7) + 1;
}

function load(d) {
  DATA = validate(d);
  DATA.dims.region = DATA.dims.region.map(regionLabel);
  // Each fact row gets its month key (ym) and week key (wk) once. Stored years/weeks are left as they
  // are, so a saved data file stays in the original numbering and is re-keyed the same way next load.
  const isoIdx = new Set(ISO_WEEK_REGIONS.map((n) => DATA.dims.region.indexOf(n)).filter((i) => i >= 0));
  for (const t of [DATA.pod, DATA.ship, DATA.delay]) for (const r of t) {
    r.ym = r[4] * 100 + r[6];
    r.wk = isoIdx.has(r[0]) ? isoWeekKey(r[4], r[5]) : r[4] * 100 + r[5];
  }
  if (DATA.checks) DATA.checks = Object.fromEntries(Object.entries(DATA.checks).map(([k, v]) => [regionLabel(k), v]));
  dims = DATA.dims; NR = dims.region.length;
  const vol = new Map();
  for (const t of [DATA.pod, DATA.ship]) for (const r of t) vol.set(r.ym, (vol.get(r.ym) || 0) + r[7]);
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
  revUi(); buildFilters(); aggregate(); renderMini(); setTab(S.tab);
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
$('#revBtn').addEventListener('click', () => { S.revOn = !showRev(); saveState(); revUi(); render(); });
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
  const ym = r.ym;
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
    for (const o of [bump(podW, r.wk), bump(regP, r[0]), tot]) { o.h += r[7]; o.gl += r[8]; o.nl += r[9]; o.rev += r[10]; }
  }
  for (const r of DATA.ship) {
    if (!pass(r)) continue;
    shipRows.push(r);
    const wk = r.wk;
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
  const line = (label, w) => `<td>${label}</td><td>${fmtN(w.h)}</td>${showRev() ? `<td>${fmt$(w.rev)}</td>` : ''}<td>${fmtN(w.gl)}</td><td>${fmtN(w.nl)}</td><td>${fmtP(1 - w.gl / w.h)}</td><td>${fmtP(1 - w.nl / w.h, 2)}</td>`;
  $('#tWeek').innerHTML = `<table><thead><tr><th>POD week</th><th>HAWBs</th>${showRev() ? '<th>Total revenue</th>' : ''}<th>Gross late</th><th>Net late</th><th>On-time gross %</th><th>On-time net %</th></tr></thead><tbody>` +
    keys.slice().reverse().map((k) => k === thisWeek().key ? `<tr class="now">${line(wkLabel(k, multi) + nowTag, A.podW.get(k))}</tr>`
      : `<tr>${line(wkLabel(k, multi) + (isPartial(k) ? partTag : ''), A.podW.get(k))}</tr>`).join('') +
    `<tr class="total">${line('Grand total', t)}</tr></tbody></table>`;
}

function regionTable() {
  const ids = [...new Set([...A.regP.keys(), ...A.regS.keys()])].sort((a, b) => a - b);
  if (!ids.length) { $('#tRegion').innerHTML = emptyMsg('No data for these filters.'); return; }
  const line = (label, shipped, p) => {
    const g = p.h ? 1 - p.gl / p.h : null;
    return `<td>${label}</td><td>${fmtN(shipped)}</td><td>${fmtN(p.h)}</td>${showRev() ? `<td>${fmt$c(p.rev)}</td>` : ''}<td>${fmtP(g)}<span class="meter"><b style="width:${((g || 0) * 100).toFixed(1)}%"></b></span></td><td>${p.h ? fmtP(1 - p.nl / p.h, 2) : '–'}</td>`;
  };
  $('#tRegion').innerHTML = `<table><thead><tr><th>Region</th><th>Shipped</th><th>Delivered</th>${showRev() ? '<th>Revenue</th>' : ''}<th>Gross OTP</th><th>Net OTP</th></tr></thead><tbody>` +
    ids.map((i) => `<tr>${line(esc(dims.region[i]), A.regS.get(i) || 0, A.regP.get(i) || { h: 0, gl: 0, nl: 0, rev: 0 })}</tr>`).join('') +
    `<tr class="total">${line('Total', A.tot.shipped, A.tot)}</tr></tbody></table>`;
}

// rows grouped by fact column `col` (1 = customer, 3 = CSR), columns = weeks
function matrix(target, moreTarget, facts, col, mode, opts = {}) {
  const dimKey = col === 1 ? 'cust' : 'csr';
  const cells = new Map(), rowsT = new Map(), colT = new Map(), regionOf = new Map();
  const bump = (m, k) => { let o = m.get(k); if (!o) m.set(k, (o = { h: 0, gl: 0, nl: 0 })); return o; };
  const byMonth = opts.period === 'month';
  for (const r of facts) {
    const wk = byMonth ? r.ym : r.wk, key = r[col];
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
  const wkHead = (w) => byMonth ? `<th>${ymLabel(w)}</th>` : w === thisWeek().key ? `<th class="now" title="In-Progress">${wkLabel(w, multi)}</th>`
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
    const per = byMonth ? (mode === 'count' ? 'Ship month' : 'POD month') : (mode === 'count' ? 'Ship week' : 'POD week');
    let body = `<div class="tt">${esc(title)}</div>` + row(wk == null ? 'Period' : per, wk == null ? 'Total' : byMonth ? ymLabel(wk) : wkLabel(wk, true)) + row('HAWBs', fmtN(c.h));
    if (mode !== 'count') body += row('Gross late', fmtN(c.gl)) + row('Net late', fmtN(c.nl)) + row('On-time gross', fmtP(1 - c.gl / c.h)) + row('On-time net', fmtP(1 - c.nl / c.h, 2));
    if (wk != null && !byMonth) body += partNote(wk);
    showTip(body, ev);
  };
  el.onmouseleave = hideTip;
  if (moreTarget) {
    const mt = $(moreTarget);
    const noun = dimKey === 'cust' ? 'customer' : 'CSR';
    mt.innerHTML = `<span class="note">${keys.length} ${noun}${keys.length === 1 ? '' : 's'}${q ? ' matching your search' : ''}, A to Z.</span>`;
  }
}

function choice(sel, key, opts) {
  const el = $(sel);
  el.innerHTML = opts.map(([v, l]) => `<button type="button" data-v="${v}" aria-pressed="${S[key] === v}">${l}</button>`).join('');
  el.onclick = (e) => { const b = e.target.closest('button'); if (!b) return; S[key] = b.dataset.v; saveState(); render(); };
}
function segControl(sel, key) {
  const el = $(sel);
  el.innerHTML = [['hawb', 'HAWBs'], ['gross', 'Gross %'], ['net', 'Net %']].map(([v, l]) => `<button type="button" data-v="${v}" aria-pressed="${S[key] === v}">${l}</button>`).join('');
  el.onclick = (e) => { const b = e.target.closest('button'); if (!b) return; S[key] = b.dataset.v; saveState(); render(); };
}

// Ship rows for CSR volume views, with CSR_SPLITS applied.
function csrShipRows() {
  const rules = CSR_SPLITS.map((x) => ({ ...x, region: dims.region.indexOf(x.region), from: dims.csr.indexOf(x.from), to: dims.csr.indexOf(x.to) }))
    .filter((x) => x.region >= 0 && x.from >= 0 && x.to >= 0);
  if (!rules.length) return A.shipRows;
  const out = [];
  for (const r of A.shipRows) {
    const rule = rules.find((x) => x.region === r[0] && x.from === r[3] && r.wk < x.beforeWeek);
    if (!rule) { out.push(r); continue; }
    const keep = r.slice(), give = r.slice();
    keep[7] = r[7] * (1 - rule.share); give[3] = rule.to; give[7] = r[7] * rule.share;
    for (const x of [keep, give]) { x.ym = r.ym; x.wk = r.wk; }
    out.push(keep, give);
  }
  return out;
}

function csrTable() {
  const m = new Map();
  const get = (k) => { let o = m.get(k); if (!o) m.set(k, (o = { s: 0, h: 0, gl: 0, nl: 0, cust: new Set() })); return o; };
  for (const r of A.shipRows) get(r[3]).cust.add(r[1]);
  for (const r of csrShipRows()) get(r[3]).s += r[7];
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
    (showRev() ? tile('Total revenue', fmt$c(t.rev), 'on delivered HAWBs') : '') +
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
      <li><b>Weeks</b> follow <code>WEEKNUM(date, 2)</code> (weeks start Monday) and are kept apart by year, so W50 ’25 sorts before W1 ’26.
        <b>COSTCO+ and AMAZON</b> use ISO weeks, like their OTP summaries: week 1 is the full week containing Jan 1, so Dec 29, 2025 – Jan 4, 2026 is W1 ’26 for those regions.</li>
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
const passNoMonth = (r, valid) => valid.has(r.ym) && RS.has(r[0]) && (S.account < 0 || r[2] === S.account) && (S.csr < 0 || r[3] === S.csr) && (S.cust < 0 || r[1] === S.cust);

// Sums the given fact columns per Monday-week, filling missing weeks with zeros.
function weekly(table, cols) {
  const valid = new Set(months), m = new Map();
  for (const r of table) {
    if (!passNoMonth(r, valid)) continue;
    const k = mondayIdx(Math.floor(r.wk / 100), r.wk % 100);
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

// Projection numbers for the current filters (used by the Projections tab and the PowerPoint deck).
function fcCompute(H, N) {
  const shipT = trimPartial(weekly(DATA.ship, [7])), podT = trimPartial(weekly(DATA.pod, [7, 8, 9, 10]));
  const ship = shipT.series, pod = podT.series;
  if (ship.length < 4 || pod.length < 4) return { err: 'Not enough weekly history for these filters. Projections need at least 4 full weeks.' };
  const sb = ship.slice(-N), pb = pod.slice(-N);
  const vol = trendProject(sb.map((p) => p.v[0]), H), rev = trendProject(pb.map((p) => p.v[3]), H);
  const gross = rateProject(pb.map((p) => p.v[0] - p.v[1]), pb.map((p) => p.v[0]), H);
  const net = rateProject(pb.map((p) => p.v[0] - p.v[2]), pb.map((p) => p.v[0]), H);
  if (!gross) return { err: 'No delivered shipments in the basis weeks for these filters.' };
  return { shipT, ship, pod, sb, pb, vol, rev, gross, net };
}

function projections() {
  const H = S.fcH || 8, N = S.fcBasis || 12, c = T();
  [['#fcHorizon', 'fcH', [4, 8, 12]], ['#fcBasis', 'fcBasis', [8, 12, 26]]].forEach(([sel, key, opts]) => {
    const el = $(sel);
    el.innerHTML = opts.map((v) => `<button type="button" data-v="${v}" aria-pressed="${(S[key] || (key === 'fcH' ? 8 : 12)) === v}">${v} weeks</button>`).join('');
    el.onclick = (e) => { const b = e.target.closest('button'); if (!b) return; S[key] = +b.dataset.v; saveState(); render(); };
  });
  const clear = (msg) => { for (const s of ['#fcTiles', '#fcVol', '#fcOtp', '#fcRev', '#fcTable']) $(s).innerHTML = ''; $('#fcNote').textContent = msg; };
  const F = fcCompute(H, N);
  if (F.err) { clear(F.err); return; }
  const { shipT, ship, pod, sb, pb, vol, rev, gross, net } = F;

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
    (showRev() ? tile(`revenue, next ${H} weeks`, fmt$c(sum(rev.out, 'mid')), `likely ${fmt$c(sum(rev.out, 'lo'))} – ${fmt$c(sum(rev.out, 'hi'))}`) : '');

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
  $('#fcRevCard').hidden = !showRev();
  const r = build(pod, (p) => p.v[3], rev.out);
  if (showRev()) fcChart($('#fcRev'), r.keys, r.nActual, [{ color: 's1', ...r }], { money: true, tip: (i) => {
    const isP = i >= r.nActual;
    return head(r.keys, i, isP) + (isP ? row('Projected revenue', fmt$(r.mid[i]), c.s1) + row('Likely range', `${fmt$(r.lo[i])} – ${fmt$(r.hi[i])}`) : row('Delivered revenue', fmt$(r.actual[i]), c.s1));
  } });

  const lastShip = ship[ship.length - 1].k;
  const tw = thisWeek(), isNowIdx = (i) => i === mondayIdx(tw.year, tw.week);
  const wkName = (h) => `${idxLabel(lastShip + h + 1)}, ${idxDate(lastShip + h + 1).getUTCFullYear()}`;
  const fig = (v, lo, hi) => `<td>${v}<span class="rg">${lo} – ${hi}</span></td>`;
  $('#fcTable').innerHTML = `<table><thead><tr><th>Week of</th><th>HAWBs shipped</th><th>Gross OTP</th><th>Net OTP</th>${showRev() ? '<th>Revenue</th>' : ''}</tr></thead><tbody>` +
    vol.out.map((p, h) => `<tr${isNowIdx(lastShip + h + 1) ? ' class="now"' : ''}><td>${wkName(h)}${isNowIdx(lastShip + h + 1) ? nowTag : ''}</td>` +
      fig(fmtN(p.mid), fmtN(p.lo), fmtN(p.hi)) + fig(fmtP(gross[h].mid), fmtP(gross[h].lo), fmtP(gross[h].hi)) +
      fig(fmtP(net[h].mid, 2), fmtP(net[h].lo, 2), fmtP(net[h].hi, 2)) + (showRev() ? fig(fmt$(rev.out[h].mid), fmt$(rev.out[h].lo), fmt$(rev.out[h].hi)) : '') + '</tr>').join('') +
    `<tr class="total"><td>Total</td>${fig(fmtN(sum(vol.out, 'mid')), fmtN(sum(vol.out, 'lo')), fmtN(sum(vol.out, 'hi')))}<td></td><td></td>` +
    `${showRev() ? fig(fmt$(sum(rev.out, 'mid')), fmt$(sum(rev.out, 'lo')), fmt$(sum(rev.out, 'hi'))) : ''}</tr></tbody></table>`;
}

// ---------- export to Excel and copy tables ----------
// Export builds a workbook from the current filters, laid out like the OTP summary workbooks
// (Ship Volume by Week, OTP BY WEEK with PODS / RAW% / NET%, monthly versions, CSR tables).
// The spreadsheet library is shared on the bus and only loaded when someone exports.
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
const xN = (v, z = '#,##0') => (v ? { t: 'n', v, z } : '');              // blank for zero, like a pivot
const xP = (v, z = '0.0%') => (v == null || !isFinite(v) ? '' : { t: 'n', v, z });
const xWeek = (k, multi) => 'WK' + (k % 100) + (multi ? " '" + String(Math.floor(k / 100)).slice(2) : '');
const x0 = (v, z = '#,##0') => ({ t: 'n', v: v || 0, z });              // zero shown as 0 in summary tables

function groupFacts(facts, rowKey, colKey, otp) {
  const cells = new Map(), rows = new Map(), cols = new Map(), all = { h: 0, gl: 0, nl: 0 };
  const bump = (m, k) => { let o = m.get(k); if (!o) m.set(k, (o = { h: 0, gl: 0, nl: 0 })); return o; };
  for (const r of facts) {
    const rk = rowKey(r), ck = colKey(r);
    for (const o of [bump(cells, rk + '|' + ck), bump(rows, rk), bump(cols, ck), all]) { o.h += r[7]; if (otp) { o.gl += r[8]; o.nl += r[9]; } }
  }
  return { cells, rows, cols, all, colKeys: [...cols.keys()].sort((a, b) => a - b) };
}
function regionsOf(facts, col) {
  const m = new Map();
  for (const r of facts) { const k = r[col]; m.set(k, m.has(k) && m.get(k) !== r[0] ? -1 : r[0]); }
  return (k) => (m.get(k) >= 0 ? dims.region[m.get(k)] : 'Multiple');
}
// Volume table: one row per customer/CSR, one column per week or month.
function volumeBlock(facts, col, period, label) {
  const byMonth = period === 'month', g = groupFacts(facts, (r) => r[col], (r) => (byMonth ? r.ym : r.wk), false);
  const multi = multiYear(g.colKeys), head = (k) => (byMonth ? ymLabel(k) : xWeek(k, multi));
  const names = dims[col === 1 ? 'cust' : 'csr'], reg = col === 1 ? regionsOf(facts, 1) : null;
  const ids = [...g.rows.keys()].sort(byName(names));
  const aoa = [[label, ...g.colKeys.map(head), 'Total', ...(reg ? ['Region'] : [])]];
  aoa.push(['Grand Total', ...g.colKeys.map((k) => xN(g.cols.get(k).h)), xN(g.all.h), ...(reg ? [''] : [])]);
  for (const id of ids) aoa.push([names[id], ...g.colKeys.map((k) => xN(g.cells.get(id + '|' + k)?.h)), xN(g.rows.get(id).h), ...(reg ? [reg(id)] : [])]);
  return { aoa, merges: [] };
}
// OTP table: per week or month, three columns PODS / RAW% (gross) / NET%.
function otpBlock(facts, col, period, label) {
  const byMonth = period === 'month', g = groupFacts(facts, (r) => r[col], (r) => (byMonth ? r.ym : r.wk), true);
  const multi = multiYear(g.colKeys), head = (k) => (byMonth ? ymLabel(k) : xWeek(k, multi));
  const names = dims[col === 1 ? 'cust' : 'csr'], ids = [...g.rows.keys()].sort(byName(names));
  const trip = (o) => (o && o.h ? [xN(o.h), xP(1 - o.gl / o.h), xP(1 - o.nl / o.h, '0.00%')] : ['', '', '']);
  const top = [byMonth ? 'MONTH' : 'WEEK'], sub = [label], merges = [];
  [...g.colKeys, 'total'].forEach((k, i) => {
    top.push(k === 'total' ? 'TOTAL' : head(k), '', ''); sub.push('PODS', 'RAW%', 'NET%');
    merges.push({ s: { r: 0, c: 1 + i * 3 }, e: { r: 0, c: 3 + i * 3 } });
  });
  const aoa = [top, sub, ['Grand Total', ...g.colKeys.flatMap((k) => trip(g.cols.get(k))), ...trip(g.all)]];
  for (const id of ids) aoa.push([names[id], ...g.colKeys.flatMap((k) => trip(g.cells.get(id + '|' + k))), ...trip(g.rows.get(id))]);
  return { aoa, merges };
}
// Stack titled blocks down one sheet with blank rows between, like the summary's CSR sheets.
function stackBlocks(blocks) {
  const aoa = [], merges = [];
  for (const [title, b] of blocks) {
    if (aoa.length) aoa.push([], []);
    aoa.push([title]);
    const off = aoa.length;
    for (const row of b.aoa) aoa.push(row);
    for (const m of b.merges) merges.push({ s: { r: m.s.r + off, c: m.s.c }, e: { r: m.e.r + off, c: m.e.c } });
  }
  return { aoa, merges };
}
function sheetFrom(XLSX, block, firstCol = 34) {
  const ws = XLSX.utils.aoa_to_sheet(block.aoa);
  const width = Math.max(1, ...block.aoa.map((r) => r.length));
  ws['!cols'] = Array.from({ length: width }, (_, i) => ({ wch: i === 0 ? firstCol : 11 }));
  if (block.merges.length) ws['!merges'] = block.merges;
  return ws;
}
function filterLines() {
  const lines = [
    ['OTP Dashboard export'],
    ['Exported', new Date().toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' })],
    ['Data built', fmtBuilt(DATA.generated)],
    ['Source files', (DATA.sources || []).join(', ')],
    [],
    ['Regions', S.regions.length === NR ? 'All regions' : S.regions.map((i) => dims.region[i]).join(', ')],
    ['Account', S.account >= 0 ? dims.account[S.account] : 'All accounts'],
    ['CSR', S.csr >= 0 ? dims.csr[S.csr] : 'All CSRs'],
    ['Customer', S.cust >= 0 ? dims.cust[S.cust] : 'All customers'],
    ['Months', `${ymLabel(S.from)} – ${ymLabel(S.to)} (ship month for volume, POD month for OTP)`],
    [],
    ['RAW%', 'Gross on-time: delivered on or before the adjusted due date'],
    ['NET%', 'Net on-time: late shipments count against it only when the delay code is Controllable'],
    ['Weeks', `WEEKNUM(date, 2), weeks start Monday. ${ISO_WEEK_REGIONS.join(' and ')} use ISO weeks (Dec 29 – Jan 4 is WK1).`],
  ];
  return { aoa: lines, merges: [] };
}
function exportName() {
  const reg = S.regions.length === NR ? 'All regions' : S.regions.map((i) => dims.region[i]).join(' ');
  const who = [S.account >= 0 ? dims.account[S.account] : '', S.csr >= 0 ? dims.csr[S.csr] : '', S.cust >= 0 ? dims.cust[S.cust] : ''].filter(Boolean).join(' ');
  const name = `OTP export - ${reg}${who ? ' - ' + who : ''} - ${ymLabel(S.from)} to ${ymLabel(S.to)}`;
  return name.replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').slice(0, 150) + '.xlsx';
}
async function exportExcel() {
  const btn = $('#exportBtn'), label = btn.textContent;
  btn.disabled = true; btn.textContent = 'Preparing…';
  try {
    const XLSX = await loadXlsx(), wb = XLSX.utils.book_new();
    const add = (name, block, w) => XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, block, w), name);
    add('Filters', filterLines(), 22);
    add('Ship Volume by Week', volumeBlock(A.shipRows, 1, 'week', 'Customer'));
    add('OTP by Week', otpBlock(A.podRows, 1, 'week', 'CUSTOMER'));
    add('Ship Volume by Month', volumeBlock(A.shipRows, 1, 'month', 'Customer'));
    add('OTP by Month', otpBlock(A.podRows, 1, 'month', 'CUSTOMER'));
    add('Ship Vol CSR', stackBlocks([
      ['BY MONTH', volumeBlock(csrShipRows(), 3, 'month', 'CSR')], ['BY WEEK', volumeBlock(csrShipRows(), 3, 'week', 'CSR')]]), 26);
    add('OTP by CSR', stackBlocks([['BY MONTH', otpBlock(A.podRows, 3, 'month', 'CSR')], ['BY WEEK', otpBlock(A.podRows, 3, 'week', 'CSR')]]), 26);
    // dashboard tables
    const wkKeys = sortedKeys(A.podW), multi = multiYear(wkKeys);
    const R = showRev(), only = (x) => (R ? [x] : []);
    const line = (label, w) => [label, x0(w.h), ...only(x0(w.rev, '$#,##0')), x0(w.gl), x0(w.nl), xP(1 - w.gl / w.h), xP(1 - w.nl / w.h, '0.00%')];
    add('Weekly Summary', { aoa: [['POD week', 'HAWBs', ...only('Total revenue'), 'Gross late', 'Net late', 'On-time gross %', 'On-time net %'],
      ...wkKeys.map((k) => line(xWeek(k, multi), A.podW.get(k))), ...(A.tot.h ? [line('Grand total', A.tot)] : [])], merges: [] }, 14);
    const regIds = [...new Set([...A.regP.keys(), ...A.regS.keys()])].sort((a, b) => a - b), none = { h: 0, gl: 0, nl: 0, rev: 0 };
    const rline = (label, s, p) => [label, x0(s), x0(p.h), ...only(x0(p.rev, '$#,##0')), p.h ? xP(1 - p.gl / p.h) : '', p.h ? xP(1 - p.nl / p.h, '0.00%') : ''];
    add('By Region', { aoa: [['Region', 'Shipped', 'Delivered', ...only('Revenue'), 'Gross OTP', 'Net OTP'],
      ...regIds.map((i) => rline(dims.region[i], A.regS.get(i) || 0, A.regP.get(i) || none)), rline('Total', A.tot.shipped, A.tot)], merges: [] }, 16);
    const late = [...A.delay.entries()].map(([k, v]) => [dims.delay[k], v[0], v[1]]).sort((a, b) => b[1] + b[2] - a[1] - a[2]);
    const lateTot = late.reduce((s, x) => s + x[1] + x[2], 0);
    add('Delay Codes', { aoa: [['Delay code (late shipments)', 'Late HAWBs', 'Controllable', 'Uncontrollable', 'Share of late'],
      ...late.map(([n, c, u]) => [n, x0(c + u), x0(c), x0(u), xP(lateTot ? (c + u) / lateTot : null)])], merges: [] }, 34);
    XLSX.writeFile(wb, exportName());
    btn.textContent = 'Downloaded';
  } catch (err) {
    btn.textContent = 'Export failed';
    $('#notice').hidden = false; $('#notice').textContent = 'Export to Excel failed: ' + (err && err.message ? err.message : err);
  } finally {
    setTimeout(() => { btn.textContent = label; btn.disabled = false; }, 1800);
  }
}
$('#exportBtn').addEventListener('click', () => { if (A) exportExcel(); });

// Copy any table as tab-separated text, which pastes straight into Excel cells.
function tableText(table) {
  return [...table.rows].map((tr) => [...tr.cells].map((td) => {
    const c = td.cloneNode(true);
    c.querySelectorAll('.nowtag, .parttag, .rtag').forEach((e) => e.remove());
    c.querySelectorAll('.rg').forEach((e) => e.replaceWith(' (' + e.textContent.trim() + ')'));
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

// ---------- company PowerPoint template ----------
// Every PowerPoint export is built inside the CEVA template in template/: its first slide (text swapped), the generated slides on its
// "SIMPLE PAGE" layout so the logo and brand graphics appear on every slide, and its last slide as the
// closing page. The template's other slides are dropped.
const DEFAULT_TEMPLATE = { name: 'CEVA template', url: new URL('template/CEVA_template.pptx', import.meta.url).href, builtIn: true };
const TEMPLATE = { ...DEFAULT_TEMPLATE };   // fetched once, on the first export
async function templateBytes(tpl) {
  if (!tpl.buf) {
    const res = await fetch(tpl.url);
    if (!res.ok) throw new Error('The built-in CEVA template could not be loaded. Check your connection, or choose a template file.');
    tpl.buf = await res.arrayBuffer();
  }
  return tpl.buf;
}
const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const xAttr = (tag, name) => (new RegExp('\\s' + name.replace(':', '\\:') + '="([^"]*)"').exec(tag) || [])[1];
function resolvePart(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/'); parts.pop();
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); }
  return parts.join('/');
}
const relsPath = (p) => p.replace(/([^/]+)$/, '_rels/$1.rels');
const relList = (xml) => [...(xml || '').matchAll(/<Relationship\b[^>]*\/?>/g)].map((m) => ({ tag: m[0], id: xAttr(m[0], 'Id'), type: xAttr(m[0], 'Type') || '', target: xAttr(m[0], 'Target'), external: /TargetMode="External"/.test(m[0]) }));

// Replaces the first visible text on the title slide with the deck title, adding smaller lines under it.
function retitle(xml, lines) {
  const m = /<a:t>[^<]*\S[^<]*<\/a:t>/.exec(xml);
  if (!m) return xml;
  const pStart = xml.lastIndexOf('<a:p>', m.index), pEnd = xml.indexOf('</a:p>', m.index) + 6;
  if (pStart < 0 || pEnd < 6) return xml;
  const para = xml.slice(pStart, pEnd);
  const withText = (p, text) => p.replace(/(<a:r>[\s\S]*?<a:t>)[^<]*(<\/a:t>[\s\S]*?<\/a:r>)/, `$1${xmlEsc(text)}$2`).replace(/(<\/a:r>)[\s\S]*?(<a:endParaRPr|<\/a:p>)/, '$1$2');
  const small = (p, sz) => p.replace(/<a:lnSpc>[\s\S]*?<\/a:lnSpc>/, '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>')
    .replace(/<a:spcBef>[\s\S]*?<\/a:spcBef>/, '<a:spcBef><a:spcPts val="600"/></a:spcBef>')
    .replace(/<a:rPr\b([^>]*?)(\/?)>/, (t, attrs, close) => `<a:rPr${attrs.replace(/\ssz="\d+"/, '')} sz="${sz}"${close}>`)
    .replace(/<a:endParaRPr\b([^>]*?)(\/?)>/, (t, attrs, close) => `<a:endParaRPr${attrs.replace(/\ssz="\d+"/, '')} sz="${sz}"${close}>`);
  const out = [withText(para, lines[0]), ...lines.slice(1).map((l, i) => small(withText(para, l), i === 0 ? 2000 : 1400))].join('');
  return xml.slice(0, pStart) + out + xml.slice(pEnd);
}

async function mergeIntoTemplate(JSZip, tplBuf, genBuf, titleLines) {
  const T = await JSZip.loadAsync(tplBuf), G = await JSZip.loadAsync(genBuf);
  const read = (z, p) => (z.file(p) ? z.file(p).async('string') : Promise.resolve(null));
  let ct = await read(T, '[Content_Types].xml'), pres = await read(T, 'ppt/presentation.xml'), presRels = await read(T, 'ppt/_rels/presentation.xml.rels');
  if (!ct || !pres || !presRels) throw new Error('That file is not a PowerPoint (.pptx) template.');
  const pRels = relList(presRels), sld = [...pres.matchAll(/<p:sldId\b[^>]*\/>/g)].map((m) => ({ tag: m[0], id: +xAttr(m[0], 'id'), rid: xAttr(m[0], 'r:id') }));
  if (sld.length < 2) throw new Error('The template needs at least a first and a last slide.');
  const partOf = (rid) => resolvePart('ppt/presentation.xml', pRels.find((r) => r.id === rid).target);
  const firstPart = partOf(sld[0].rid), lastPart = partOf(sld[sld.length - 1].rid);

  // content layout: "SIMPLE PAGE" if the template has one, otherwise the layout of its second slide
  let layout = null;
  for (const f of Object.keys(T.files).filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f)).sort((a, b) => +a.match(/\d+/)[0] - +b.match(/\d+/)[0])) {
    if (/<p:cSld name="SIMPLE PAGE"/i.test(await read(T, f))) { layout = f; break; }
  }
  if (!layout) {
    const r = relList(await read(T, relsPath(partOf(sld[Math.min(1, sld.length - 1)].rid)))).find((x) => x.type.endsWith('/slideLayout'));
    layout = resolvePart(partOf(sld[1].rid), r.target);
  }

  // drop the template's middle slides (and their notes)
  for (const s of sld.slice(1, -1)) {
    const part = partOf(s.rid), rels = relList(await read(T, relsPath(part)));
    for (const r of rels) if (r.type.endsWith('/notesSlide')) { const np = resolvePart(part, r.target); T.remove(np); T.remove(relsPath(np)); ct = ct.replace(new RegExp(`<Override PartName="/${np}"[^>]*/>`), ''); }
    T.remove(part); T.remove(relsPath(part));
    ct = ct.replace(new RegExp(`<Override PartName="/${part}"[^>]*/>`), '');
    presRels = presRels.replace(pRels.find((r) => r.id === s.rid).tag, '');
  }

  // copy the generated slides (all but the generated title slide) with their charts
  const gPres = await read(G, 'ppt/presentation.xml'), gRels = relList(await read(G, 'ppt/_rels/presentation.xml.rels'));
  const gSlides = [...gPres.matchAll(/<p:sldId\b[^>]*\/>/g)].map((m) => resolvePart('ppt/presentation.xml', gRels.find((r) => r.id === xAttr(m[0], 'r:id')).target)).slice(1);
  let slideNo = Math.max(0, ...Object.keys(T.files).map((f) => +(/^ppt\/slides\/slide(\d+)\.xml$/.exec(f) || [])[1] || 0)) + 1;
  let chartNo = 1000, maxId = Math.max(...sld.map((s) => s.id)), newIds = '';
  // Some template layouts carry leftover sample text in plain text boxes (e.g. "CLICK TO CHANGE STYLES OF
  // MASK TEXT"). Remove those from the chosen layout in the exported copy; placeholders and fields stay.
  T.file(layout, (await read(T, layout)).replace(/<p:sp>[\s\S]*?<\/p:sp>/g,
    (sp) => (!sp.includes('<p:ph') && !sp.includes('<a:fld') && /<a:t>[^<]*\S/.test(sp) ? '' : sp)));
  const layoutTarget = '../slideLayouts/' + layout.split('/').pop();
  for (const [k, gp] of gSlides.entries()) {
    const np = `ppt/slides/slide${slideNo++}.xml`;
    T.file(np, await read(G, gp));
    let rels = '';
    for (const r of relList(await read(G, relsPath(gp)))) {
      if (r.type.endsWith('/notesSlide')) continue;
      if (r.type.endsWith('/slideLayout')) { rels += `<Relationship Id="${r.id}" Type="${r.type}" Target="${layoutTarget}"/>`; continue; }
      if (r.type.endsWith('/chart')) {
        const src = resolvePart(gp, r.target), cn = `ppt/charts/chart${chartNo++}.xml`;
        T.file(cn, await read(G, src));
        let crels = '';
        for (const cr of relList(await read(G, relsPath(src)))) {
          const esrc = resolvePart(src, cr.target), edst = `ppt/embeddings/otp_${cn.match(/\d+/)[0]}_${esrc.split('/').pop()}`;
          T.file(edst, await G.file(esrc).async('uint8array'));
          crels += `<Relationship Id="${cr.id}" Type="${cr.type}" Target="../embeddings/${edst.split('/').pop()}"/>`;
        }
        T.file(relsPath(cn), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${crels}</Relationships>`);
        ct = ct.replace('</Types>', `<Override PartName="/${cn}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
        rels += `<Relationship Id="${r.id}" Type="${r.type}" Target="/${cn}"/>`;
        continue;
      }
      if (!r.external) {   // images or other media
        const src = resolvePart(gp, r.target), dst = `ppt/media/otp_${slideNo}_${src.split('/').pop()}`;
        T.file(dst, await G.file(src).async('uint8array'));
        rels += `<Relationship Id="${r.id}" Type="${r.type}" Target="../media/${dst.split('/').pop()}"/>`;
      } else rels += r.tag;
    }
    T.file(relsPath(np), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
    ct = ct.replace('</Types>', `<Override PartName="/${np}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
    const rid = `rIdOtp${k + 1}`;
    presRels = presRels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${np.replace('ppt/', '')}"/></Relationships>`);
    newIds += `<p:sldId id="${++maxId}" r:id="${rid}"/>`;
  }
  if (!/Extension="xlsx"/i.test(ct)) ct = ct.replace('<Default ', '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Default ');
  pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sld[0].tag}${newIds}${sld[sld.length - 1].tag}</p:sldIdLst>`);
  T.file('ppt/presentation.xml', pres); T.file('ppt/_rels/presentation.xml.rels', presRels); T.file('[Content_Types].xml', ct);

  // title slide text
  T.file(firstPart, retitle(await read(T, firstPart), titleLines));

  // remove media nothing refers to any more (the dropped slides' icons)
  const used = new Set();
  for (const f of Object.keys(T.files).filter((f) => f.endsWith('.rels'))) {
    const owner = f.replace(/_rels\/([^/]+)\.rels$/, '$1');
    for (const r of relList(await read(T, f))) if (!r.external && r.target) used.add(resolvePart(owner, r.target));
  }
  for (const f of Object.keys(T.files)) if (/^ppt\/media\//.test(f) && !T.files[f].dir && !used.has(f)) T.remove(f);
  const app = await read(T, 'docProps/app.xml');
  if (app) T.file('docProps/app.xml', app.replace(/<Slides>\d+<\/Slides>/, `<Slides>${gSlides.length + 2}</Slides>`));
  void lastPart;
  return T.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', compression: 'DEFLATE' });
}

function downloadBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
// ---------- export to PowerPoint ----------
// Builds a board-ready deck from the current filters in a strategy-consulting style: full-sentence
// action titles, an executive summary, numbered exhibits with source lines, a navy/blue palette with
// one highlight color, and native (editable) PowerPoint charts. The library is shared on the bus and
// only loaded when someone exports.
// Colors follow the CEVA Logistics brand: navy #1C2546 and red #E30613. "blue" is the accent/highlight
// slot (CEVA red); "cyan" is the secondary series (a slate tint of the navy).
let PPT = { navy: '1C2546', blue: 'E30613', cyan: '7180AE', ink: '1A1A1A', body: '333333', muted: '6F7385', gray: 'B9BCC8', light: 'E4E6ED', pale: 'F3F4F8',
  red: 'E30613', green: '00875A', serif: 'Georgia', sans: 'Arial' };
let pptLoading = null;
function loadPpt() {
  if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
  if (!pptLoading) {
    pptLoading = new Promise((ok, fail) => {
      const s = document.createElement('script');
      s.src = '/assets/vendor/pptxgen.bundle.js';
      s.onload = () => ok(window.PptxGenJS);
      s.onerror = () => { pptLoading = null; fail(new Error('The PowerPoint library could not be loaded. Check your connection and try again.')); };
      document.head.append(s);
    });
  }
  return pptLoading;
}
const pctTxt = (p, d = 1) => (p == null || !isFinite(p) ? '–' : (p * 100).toFixed(d) + '%');
const ptsTxt = (d) => (d == null || !isFinite(d) ? '' : `${d >= 0 ? '+' : '–'}${Math.abs(d * 100).toFixed(1)} pts`);
function filterText() {
  const parts = [S.regions.length === NR ? 'All regions' : S.regions.map((i) => dims.region[i]).join(', ')];
  if (S.account >= 0) parts.push(dims.account[S.account]);
  if (S.csr >= 0) parts.push('CSR: ' + dims.csr[S.csr]);
  if (S.cust >= 0) parts.push(dims.cust[S.cust]);
  parts.push(S.from === S.to ? ymLabel(S.from) : `${ymLabel(S.from)} – ${ymLabel(S.to)}`);
  return parts.join(' · ');
}
// Everything the deck shows, from the already-filtered aggregates.
function deckData() {
  const done = (k) => k < thisWeek().key;                          // leave out the week in progress
  const podKeys = sortedKeys(A.podW).filter(done), shipKeys = sortedKeys(A.shipW).filter(done);
  const rate = (ks, f) => { let h = 0, x = 0; for (const k of ks) { const w = A.podW.get(k); h += w.h; x += w[f]; } return h ? 1 - x / h : null; };
  const recent = podKeys.slice(-4), prior = podKeys.slice(-8, -4);
  const cust = new Map();
  for (const r of A.podRows) { let o = cust.get(r[1]); if (!o) cust.set(r[1], (o = { h: 0, gl: 0, nl: 0, reg: r[0] })); o.h += r[7]; o.gl += r[8]; o.nl += r[9]; if (o.reg !== r[0]) o.reg = -1; }
  const custRows = [...cust.entries()].map(([k, o]) => ({ name: dims.cust[k], reg: o.reg >= 0 ? dims.region[o.reg] : 'Multiple', ...o, g: 1 - o.gl / o.h, n: 1 - o.nl / o.h }));
  const minVol = Math.max(30, Math.round(A.tot.h * 0.002));
  const delays = [...A.delay.entries()].map(([k, v]) => ({ label: dims.delay[k], c: v[0], u: v[1], t: v[0] + v[1] })).sort((a, b) => b.t - a.t);
  return {
    podKeys: podKeys.slice(-13), shipKeys: shipKeys.slice(-13),
    grossRecent: rate(recent, 'gl'), grossPrior: rate(prior, 'gl'), netRecent: rate(recent, 'nl'), netPrior: rate(prior, 'nl'),
    lowest: custRows.filter((c) => c.h >= minVol).sort((a, b) => a.g - b.g).slice(0, 8), minVol,
    mostCtrl: custRows.filter((c) => c.nl > 0).sort((a, b) => b.nl - a.nl).slice(0, 8),
    delays, lateTotal: delays.reduce((s, x) => s + x.t, 0), ctrlTotal: delays.reduce((s, x) => s + x.c, 0),
  };
}
async function exportPpt() {
  const btn = $('#pptBtn'), label = btn.textContent, basePPT = PPT;
  btn.disabled = true; btn.textContent = 'Preparing…';
  // With a company template, follow its theme: navy #051038, red #FF0000, Arial throughout.
  if (TEMPLATE) PPT = { ...basePPT, navy: '051038', blue: 'FF0000', red: 'FF0000', serif: 'Arial' };
  try {
    const PptxGenJS = await loadPpt(), pres = new PptxGenJS();
    pres.layout = 'LAYOUT_WIDE'; pres.title = 'On-time performance review';
    const W = 13.333, M = 0.6, CW = W - 2 * M, D = deckData(), t = A.tot, filt = filterText(), tw = thisWeek();
    const multi = multiYear([...D.podKeys, ...D.shipKeys]), wl = (k) => wkLabel(k, multi).replace('’', "'");
    const g = t.h ? 1 - t.gl / t.h : null, n = t.h ? 1 - t.nl / t.h : null;
    const source = `Source: OTP Dashboard, regional shipment exports (${(DATA.sources || []).length} files), data as of ${fmtBuilt(DATA.generated)}. Scope: ${filt}.`;
    const sans = { fontFace: PPT.sans, color: PPT.body }, serif = { fontFace: PPT.serif, color: PPT.navy };
    let page = 0, exhibit = 0;

    // Standard content slide: section tracker, action title, rule, source line and page number.
    const slide = (tracker, title, note) => {
      const s = pres.addSlide(); page++;
      s.background = { color: 'FFFFFF' };
      const TW = TEMPLATE ? CW - 1.9 : CW;   // keep clear of the template's logo, top right
      s.addText(tracker.toUpperCase(), { ...sans, x: M, y: 0.28, w: TW, h: 0.28, fontSize: 9, bold: true, color: PPT.blue, charSpacing: 1.5 });
      s.addText(title, { ...serif, x: M, y: 0.55, w: TW, h: 0.95, fontSize: 24, bold: !!TEMPLATE, valign: 'top', fit: 'shrink' });
      s.addShape(pres.ShapeType.line, { x: M, y: 1.55, w: CW, h: 0, line: { color: PPT.navy, width: 1 } });
      s.addText((note ? note + '  ' : '') + source, { ...sans, x: M, y: 6.85, w: CW - 0.8, h: 0.42, fontSize: 8, color: PPT.muted, valign: 'top' });
      s.addText(String(page), { ...sans, x: W - M - 0.6, y: 6.85, w: 0.6, h: 0.3, fontSize: 9, color: PPT.muted, align: 'right' });
      return s;
    };
    // Exhibit header above a chart or table: "Exhibit N", what it shows, and the unit.
    const exhibitHead = (s, x, w, what, unit) => {
      exhibit++;
      s.addText([{ text: `Exhibit ${exhibit}`, options: { bold: true, color: PPT.blue, breakLine: true } },
        { text: what, options: { bold: true, color: PPT.ink, breakLine: true } }, { text: unit, options: { color: PPT.muted } }],
        { ...sans, x, y: 1.7, w, h: 0.8, fontSize: 11, valign: 'top', paraSpaceAfter: 1 });
    };
    // Key takeaways panel at the right of a chart.
    const takeaways = (s, items, x = W - M - 3.25, w = 3.25) => {
      s.addShape(pres.ShapeType.rect, { x, y: 1.75, w, h: 4.9, fill: { color: PPT.pale }, line: { color: PPT.pale } });
      s.addShape(pres.ShapeType.line, { x, y: 1.75, w, h: 0, line: { color: PPT.blue, width: 2.5 } });
      s.addText('Key takeaways', { ...serif, x: x + 0.2, y: 1.9, w: w - 0.4, h: 0.4, fontSize: 14, bold: true });
      s.addText(items.map((b) => ({ text: b, options: { bullet: { indent: 12 }, breakLine: true } })),
        { ...sans, x: x + 0.2, y: 2.35, w: w - 0.4, h: 4.15, fontSize: 11.5, valign: 'top', paraSpaceAfter: 8, color: PPT.body });
    };
    const clean = { catAxisLabelFontSize: 10, valAxisLabelFontSize: 10, catAxisLabelColor: PPT.muted, valAxisLabelColor: PPT.muted,
      catAxisLabelFontFace: PPT.sans, valAxisLabelFontFace: PPT.sans, catAxisLineShow: true, catAxisLineColor: PPT.gray, valAxisLineShow: false,
      valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, legendFontFace: PPT.sans, legendFontSize: 10, legendColor: PPT.body,
      dataLabelFontFace: PPT.sans, dataLabelFontSize: 10, dataLabelColor: PPT.ink };
    const cell = (text, o = {}) => ({ text: String(text), options: { fontFace: PPT.sans, fontSize: 11, color: PPT.body, align: 'right', valign: 'middle', ...o } });
    const head = (cols) => cols.map((c, i) => cell(c, { bold: true, color: PPT.navy, align: i ? 'right' : 'left', fontSize: 10, border: [{ type: 'none' }, { type: 'none' }, { pt: 1.25, color: PPT.navy }, { type: 'none' }] }));
    const rowBorder = [{ type: 'none' }, { type: 'none' }, { pt: 0.5, color: PPT.light }, { type: 'none' }];
    const body = (c) => ({ ...c, options: { ...c.options, border: rowBorder } });
    const table = (s, rows, x, y, colW, rowH = 0.36) => s.addTable(rows.map((r, i) => (i ? r.map(body) : r)), { x, y, w: colW.reduce((a, b) => a + b, 0), colW, rowH, autoPage: false });

    // ---- derived facts for titles and summary ----
    const pod = D.podKeys.map((k) => A.podW.get(k)), gw = pod.map((w) => 1 - w.gl / w.h), nw = pod.map((w) => 1 - w.nl / w.h);
    const regIds = [...new Set([...A.regP.keys(), ...A.regS.keys()])].sort((a, b) => a - b), none = { h: 0, gl: 0, nl: 0, rev: 0 };
    const regRows = regIds.filter((i) => (A.regP.get(i) || none).h).map((i) => { const p = A.regP.get(i); return { name: dims.region[i], g: 1 - p.gl / p.h, n: 1 - p.nl / p.h, nl: p.nl, h: p.h, s: A.regS.get(i) || 0, rev: p.rev }; });
    const byG = regRows.slice().sort((a, b) => b.g - a.g), byN = regRows.slice().sort((a, b) => a.n - b.n);
    const top = D.delays[0], ctrlTop = D.delays.filter((x) => x.c).sort((a, b) => b.c - a.c)[0];
    const missing = D.delays.find((x) => /missing delay code/i.test(x.label));
    const dG = D.grossRecent != null && D.grossPrior != null ? D.grossRecent - D.grossPrior : null;
    const dN = D.netRecent != null && D.netPrior != null ? D.netRecent - D.netPrior : null;
    const F = fcCompute(4, 12);

    // 1. title
    {
      const s = pres.addSlide(); page++;
      s.background = { color: PPT.navy };
      s.addShape(pres.ShapeType.line, { x: 0.9, y: 2.55, w: 1.2, h: 0, line: { color: PPT.red, width: 3.5 } });
      s.addText('On-time performance review', { fontFace: PPT.serif, color: 'FFFFFF', x: 0.9, y: 2.75, w: 11, h: 1.0, fontSize: 40 });
      s.addText(filt, { fontFace: PPT.sans, color: 'D9DCE8', x: 0.9, y: 3.8, w: 11, h: 0.45, fontSize: 16 });
      s.addText(`Week ${tw.week} review · ${new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}`,
        { fontFace: PPT.sans, color: 'A9AFC8', x: 0.9, y: 4.3, w: 11, h: 0.4, fontSize: 12 });
      s.addText('CONFIDENTIAL · INTERNAL USE', { fontFace: PPT.sans, color: 'A9AFC8', x: 0.9, y: 6.7, w: 6, h: 0.3, fontSize: 9, charSpacing: 1.5 });
    }

    // 2. executive summary
    {
      const s = slide('Executive summary', 'Net on-time is strong, but gross on-time and missing delay codes need attention');
      const pts = [
        ['Service level. ', `Net on-time was ${pctTxt(n, 2)} and gross on-time ${pctTxt(g)} across ${fmtN(t.h)} delivered HAWBs; ${fmtN(t.nl)} late shipments were controllable.`],
        dG != null ? ['Direction. ', `Over the last 4 full weeks gross on-time moved ${ptsTxt(dG)} and net on-time ${ptsTxt(dN)} versus the prior 4 weeks.`] : null,
        byG.length > 1 ? ['Regions. ', `${byG[0].name} leads gross on-time at ${pctTxt(byG[0].g)}; ${byG[byG.length - 1].name} trails at ${pctTxt(byG[byG.length - 1].g)}. ${byN[0].name} has the lowest net on-time (${pctTxt(byN[0].n, 2)}).`] : null,
        top ? ['Root causes. ', `${top.label} explains ${pctTxt(top.t / D.lateTotal)} of late shipments. Only ${pctTxt(D.ctrlTotal / D.lateTotal)} of late shipments carry a controllable code${ctrlTop ? `, led by ${ctrlTop.label} (${fmtN(ctrlTop.c)})` : ''}.`] : null,
        missing ? ['Data quality. ', `${pctTxt(missing.t / D.lateTotal)} of late shipments (${fmtN(missing.t)}) have no delay code and cannot be classified as controllable or not.`] : null,
        D.lowest.length ? ['Focus accounts. ', `${D.lowest.slice(0, 3).map((c) => c.name).join(', ')} have the lowest gross on-time; ${D.mostCtrl[0] ? D.mostCtrl[0].name + ' has the most controllable late shipments.' : ''}`] : null,
        !F.err ? ['Outlook. ', `Volume is projected at about ${fmtN(F.vol.out[0].mid)} HAWBs a week over the next 4 weeks, with gross on-time near ${pctTxt(F.gross[0].mid)}.`] : null,
      ].filter(Boolean);
      s.addText(pts.map(([lead, txt]) => [{ text: lead, options: { bold: true, color: PPT.navy, bullet: { indent: 14 } } }, { text: txt, options: { breakLine: true } }]).flat(),
        { ...sans, x: M, y: 1.8, w: CW, h: 4.9, fontSize: 14, valign: 'top', paraSpaceAfter: 11 });
    }

    // 3. at a glance
    {
      const s = slide('Performance at a glance', `Net on-time stands at ${pctTxt(n, 2)}; ${pctTxt(t.h ? t.gl / t.h : null)} of deliveries missed the due date, mostly for reasons outside our control`);
      const tiles = [
        ['Net on-time', pctTxt(n, 2), dN == null ? '' : `${ptsTxt(dN)} last 4 wks vs prior 4`, dN],
        ['Gross on-time', pctTxt(g), dG == null ? '' : `${ptsTxt(dG)} last 4 wks vs prior 4`, dG],
        ['Controllable late', fmtN(t.nl), `${pctTxt(t.h ? t.nl / t.h : null, 2)} of deliveries`, null],
        ['HAWBs shipped', fmtN(t.shipped), 'by ship date', null],
        ['HAWBs delivered', fmtN(t.h), 'with a POD', null],
        showRev() ? ['Revenue', fmt$c(t.rev), 'on delivered HAWBs', null] : ['Gross late', fmtN(t.gl), 'missed the adjusted due date', null],
      ];
      const tw3 = (CW - 2 * 0.45) / 3;
      tiles.forEach(([lab, val, note, d], i) => {
        const x = M + (i % 3) * (tw3 + 0.45), y = 1.95 + Math.floor(i / 3) * 2.35;
        s.addShape(pres.ShapeType.line, { x, y, w: tw3, h: 0, line: { color: i < 3 ? PPT.blue : PPT.gray, width: i < 3 ? 2.5 : 1 } });
        s.addText(lab, { ...sans, x, y: y + 0.12, w: tw3, h: 0.35, fontSize: 12, bold: true, color: PPT.navy });
        s.addText(val, { ...serif, x, y: y + 0.5, w: tw3, h: 0.95, fontSize: 44 });
        s.addText(note, { ...sans, x, y: y + 1.45, w: tw3, h: 0.35, fontSize: 11, color: d == null ? PPT.muted : d >= 0 ? PPT.green : PPT.red });
      });
    }

    // 4. on-time trend
    if (D.podKeys.length >= 2) {
      const labels = D.podKeys.map(wl), first = gw[0], last = gw[gw.length - 1];
      const title = `Gross on-time ${last >= first ? 'rose' : 'fell'} from ${pctTxt(first, 0)} to ${pctTxt(last, 0)} over the last ${labels.length} weeks, while net on-time held between ${pctTxt(Math.min(...nw))} and ${pctTxt(Math.max(...nw))}`;
      const s = slide('Service level', title);
      const cw = CW - 3.55;
      exhibitHead(s, M, cw, 'Gross and net on-time by POD week', `% of delivered HAWBs, ${labels[0]} – ${labels[labels.length - 1]}`);
      s.addChart(pres.ChartType.line, [{ name: 'Net on-time', labels, values: nw }, { name: 'Gross on-time', labels, values: gw }], {
        x: M, y: 2.5, w: cw, h: 4.2, ...clean, chartColors: [PPT.cyan, PPT.navy], lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 6,
        valAxisMinVal: 0, valAxisMaxVal: 1, valAxisHidden: true, showLegend: true, legendPos: 'b',
        showValue: true, dataLabelFormatCode: '0%', dataLabelPosition: 't', dataLabelFontSize: 9 });
      const best = gw.indexOf(Math.max(...gw)), worst = gw.indexOf(Math.min(...gw));
      takeaways(s, [
        `Gross on-time peaked at ${pctTxt(gw[best])} in ${labels[best]} and was lowest at ${pctTxt(gw[worst])} in ${labels[worst]}.`,
        dG != null ? `Last 4 weeks vs prior 4: gross ${ptsTxt(dG)}, net ${ptsTxt(dN)}.` : `Net on-time stayed between ${pctTxt(Math.min(...nw))} and ${pctTxt(Math.max(...nw))}.`,
        `The gap between gross and net is mostly appointment and customer-driven delays outside our control.`,
      ]);

      // 5. net close-up
      const lo = Math.min(...nw), min = Math.max(0, Math.floor((lo - 0.002) * 200) / 200), wN = nw.indexOf(lo);
      const s2 = slide('Service level', `Net on-time dipped to ${pctTxt(lo, 2)} in ${labels[wN]} and has since recovered to ${pctTxt(nw[nw.length - 1], 2)}`);
      exhibitHead(s2, M, cw, 'Net on-time by POD week', `% of delivered HAWBs, scale from ${pctTxt(min)} to show week-to-week change`);
      s2.addChart(pres.ChartType.line, [{ name: 'Net on-time', labels, values: nw }], {
        x: M, y: 2.5, w: cw, h: 4.2, ...clean, chartColors: [PPT.navy], lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 7,
        valAxisMinVal: min, valAxisMaxVal: 1, valAxisHidden: true, showLegend: false,
        showValue: true, dataLabelFormatCode: '0.00%', dataLabelPosition: 't', dataLabelFontSize: 9 });
      takeaways(s2, [
        `Lowest week: ${labels[wN]} at ${pctTxt(lo, 2)}.`,
        `Latest full week: ${labels[labels.length - 1]} at ${pctTxt(nw[nw.length - 1], 2)}.`,
        `Each 0.1 pt of net on-time is about ${fmtN(pod.reduce((s, w) => s + w.h, 0) / pod.length * 0.001)} HAWBs a week at recent volume.`,
      ]);
    }

    // 6. volume
    if (D.shipKeys.length) {
      const labels = D.shipKeys.map(wl), vals = D.shipKeys.map((k) => A.shipW.get(k));
      const avg = vals.reduce((a, b) => a + b, 0) / vals.length, lastV = vals[vals.length - 1], d = avg ? lastV / avg - 1 : 0;
      const s = slide('Volume', `Volume averaged ${fmtN(avg)} HAWBs a week; the latest full week was ${Math.abs(d * 100).toFixed(0)}% ${d >= 0 ? 'above' : 'below'} that average`);
      const cw = CW - 3.55;
      exhibitHead(s, M, cw, 'HAWBs shipped per week', `Count by ship week, ${labels[0]} – ${labels[labels.length - 1]}`);
      s.addChart(pres.ChartType.bar, [{ name: 'HAWBs shipped', labels, values: vals }], {
        x: M, y: 2.5, w: cw, h: 4.2, ...clean, barDir: 'col', barGapWidthPct: 45, chartColors: vals.map((_, i) => (i === vals.length - 1 ? PPT.blue : PPT.navy)),
        valAxisHidden: true, showLegend: false, showValue: true, dataLabelFormatCode: '#,##0', dataLabelPosition: 'outEnd', dataLabelFontSize: 9 });
      const hi = vals.indexOf(Math.max(...vals)), loI = vals.indexOf(Math.min(...vals));
      takeaways(s, [`Highest week: ${labels[hi]} with ${fmtN(vals[hi])} HAWBs.`, `Lowest week: ${labels[loI]} with ${fmtN(vals[loI])} HAWBs.`,
        `${fmtN(t.shipped)} HAWBs shipped in the selected period.`]);
    }

    // 7. regions
    if (regRows.length) {
      const title = regRows.length > 1 ? `${byG[0].name} leads gross on-time at ${pctTxt(byG[0].g)}, ${((byG[0].g - byG[byG.length - 1].g) * 100).toFixed(0)} pts ahead of ${byG[byG.length - 1].name}; ${byN[0].name} has the lowest net on-time`
        : `${regRows[0].name}: gross on-time ${pctTxt(regRows[0].g)}, net on-time ${pctTxt(regRows[0].n, 2)}`;
      const s = slide('Regions', title);
      exhibitHead(s, M, 6.0, 'Gross on-time by region', '% of delivered HAWBs, ranked');
      s.addChart(pres.ChartType.bar, [{ name: 'Gross on-time', labels: byG.map((r) => r.name), values: byG.map((r) => r.g) }], {
        x: M, y: 2.5, w: 5.9, h: 4.1, ...clean, barDir: 'bar', catAxisOrientation: 'maxMin', barGapWidthPct: 40, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 1,
        chartColors: byG.map((r) => (r.g < g ? PPT.blue : PPT.navy)), showLegend: false, showValue: true, dataLabelFormatCode: '0%', dataLabelPosition: 'outEnd', catAxisLabelFontSize: 11 });
      const x2 = M + 6.2;
      exhibitHead(s, x2, CW - 6.2, 'Regional scorecard', 'Selected period');
      table(s, [head(['Region', 'Delivered', 'Gross', 'Net', 'Ctrl. late']),
        ...byG.map((r) => [cell(r.name, { align: 'left' }), cell(fmtN(r.h)), cell(pctTxt(r.g), { color: r.g < g ? PPT.blue : PPT.body, bold: r.g < g }), cell(pctTxt(r.n, 2), { color: r.n < n ? PPT.red : PPT.body }), cell(fmtN(r.nl))]),
        [cell('Total', { align: 'left', bold: true }), cell(fmtN(t.h), { bold: true }), cell(pctTxt(g), { bold: true }), cell(pctTxt(n, 2), { bold: true }), cell(fmtN(t.nl), { bold: true })]],
        x2, 2.55, [1.8, 1.2, 0.95, 0.95, 1.0], 0.4);
      s.addText('Red: below the overall average for that measure.', { ...sans, x: x2, y: 6.4, w: CW - 6.2, h: 0.3, fontSize: 9, color: PPT.muted });
    }

    // 8. root causes
    if (D.lateTotal) {
      const s = slide('Root causes', `${top.label} drives ${pctTxt(top.t / D.lateTotal, 0)} of late shipments; controllable delays are only ${pctTxt(D.ctrlTotal / D.lateTotal)}${ctrlTop ? `, led by ${ctrlTop.label}` : ''}`,
        missing ? `${fmtN(missing.t)} late HAWBs (${pctTxt(missing.t / D.lateTotal)}) have no delay code.` : '');
      const half = (CW - 0.5) / 2, top8 = D.delays.slice(0, 8);
      exhibitHead(s, M, half, 'All late shipments by delay reason', '% of late HAWBs, top 8');
      s.addChart(pres.ChartType.bar, [{ name: 'Share', labels: top8.map((x) => x.label), values: top8.map((x) => x.t / D.lateTotal) }], {
        x: M, y: 2.5, w: half, h: 4.2, ...clean, barDir: 'bar', catAxisOrientation: 'maxMin', barGapWidthPct: 40, valAxisHidden: true,
        chartColors: top8.map((x) => (/missing delay code/i.test(x.label) ? PPT.red : x === top ? PPT.navy : PPT.gray)), showLegend: false,
        showValue: true, dataLabelFormatCode: '0%', dataLabelPosition: 'outEnd', catAxisLabelFontSize: 10 });
      const ctrlItems = D.delays.filter((x) => x.c > 0).sort((a, b) => b.c - a.c).slice(0, 8), x2 = M + half + 0.5;
      exhibitHead(s, x2, half, `Controllable late shipments by reason`, `Count of HAWBs, ${fmtN(D.ctrlTotal)} in total`);
      if (ctrlItems.length) s.addChart(pres.ChartType.bar, [{ name: 'Controllable', labels: ctrlItems.map((x) => x.label), values: ctrlItems.map((x) => x.c) }], {
        x: x2, y: 2.5, w: half, h: 4.2, ...clean, barDir: 'bar', catAxisOrientation: 'maxMin', barGapWidthPct: 40, valAxisHidden: true,
        chartColors: ctrlItems.map((_, i) => (i < 3 ? PPT.blue : PPT.gray)), showLegend: false,
        showValue: true, dataLabelFormatCode: '#,##0', dataLabelPosition: 'outEnd', catAxisLabelFontSize: 10 });
    }

    // 9. where to act
    if (D.lowest.length || D.mostCtrl.length) {
      const worst = D.lowest[0], most = D.mostCtrl[0];
      const s = slide('Priorities', `${D.lowest.length} accounts sit below ${pctTxt(D.lowest[D.lowest.length - 1]?.g, 0)} gross on-time${most ? `; ${most.name} carries the most controllable late shipments (${fmtN(most.nl)})` : ''}`);
      const half = (CW - 0.5) / 2, colW = [2.75, 1.0, 1.0, 1.0];
      exhibitHead(s, M, half, 'Lowest gross on-time accounts', `At least ${fmtN(D.minVol)} delivered HAWBs`);
      table(s, [head(['Customer', 'Delivered', 'Gross', 'Net']), ...D.lowest.map((c) => [cell(c.name, { align: 'left', fontSize: 10 }), cell(fmtN(c.h), { fontSize: 10 }),
        cell(pctTxt(c.g), { fontSize: 10, bold: true, color: PPT.blue }), cell(pctTxt(c.n, 2), { fontSize: 10 })])], M, 2.55, colW, 0.42);
      const x2 = M + half + 0.5;
      exhibitHead(s, x2, half, 'Most controllable late shipments', 'Count of late HAWBs with a controllable code');
      table(s, [head(['Customer', 'Delivered', 'Ctrl. late', 'Net']), ...D.mostCtrl.map((c) => [cell(c.name, { align: 'left', fontSize: 10 }), cell(fmtN(c.h), { fontSize: 10 }),
        cell(fmtN(c.nl), { fontSize: 10, bold: true, color: PPT.blue }), cell(pctTxt(c.n, 2), { fontSize: 10 })])], x2, 2.55, colW, 0.42);
      if (worst) s.addText(`Suggested next step: review lane and appointment practices with ${worst.name} and the top controllable-late accounts.`, { ...sans, x: M, y: 6.45, w: CW, h: 0.3, fontSize: 11, italic: true, color: PPT.navy });
    }

    // 10. CSR summary
    {
      const m = new Map();
      const get = (k) => { let o = m.get(k); if (!o) m.set(k, (o = { s: 0, h: 0, gl: 0, nl: 0 })); return o; };
      for (const r of csrShipRows()) get(r[3]).s += r[7];
      for (const r of A.podRows) { const o = get(r[3]); o.h += r[7]; o.gl += r[8]; o.nl += r[9]; }
      const keys = [...m.keys()].sort((a, b) => m.get(b).s - m.get(a).s);
      if (keys.length) {
        const shown = keys.slice(0, 12), lead = m.get(keys[0]);
        const withPod = shown.filter((k) => m.get(k).h), nets = withPod.map((k) => 1 - m.get(k).nl / m.get(k).h);
        const s = slide('Customer service', `${dims.csr[keys[0]]} handles the most volume (${fmtN(lead.s)} HAWBs); net on-time ranges from ${pctTxt(Math.min(...nets), 2)} to ${pctTxt(Math.max(...nets), 2)} across the largest CSRs`);
        exhibitHead(s, M, CW, 'CSR scorecard', `Top ${shown.length} of ${keys.length} by shipped volume`);
        table(s, [head(['CSR', 'Shipped', 'Delivered', 'Gross on-time', 'Net on-time', 'Controllable late']),
          ...shown.map((k) => { const o = m.get(k), nn2 = o.h ? 1 - o.nl / o.h : null; return [cell(dims.csr[k], { align: 'left', fontSize: 10 }), cell(fmtN(o.s), { fontSize: 10 }), cell(fmtN(o.h), { fontSize: 10 }),
            cell(o.h ? pctTxt(1 - o.gl / o.h) : '–', { fontSize: 10 }), cell(pctTxt(nn2, 2), { fontSize: 10, color: nn2 != null && nn2 < n ? PPT.red : PPT.body }), cell(fmtN(o.nl), { fontSize: 10 })]; })],
          M, 2.55, [3.7, 1.6, 1.6, 1.75, 1.75, 1.73], 0.3);
      }
    }

    // 11. outlook
    if (!F.err) {
      const hist = F.ship.slice(-8), lastK = hist[hist.length - 1].k, avg4 = F.vol.out.reduce((s, p) => s + p.mid, 0) / F.vol.out.length;
      const s = slide('Outlook', `Volume is projected at about ${fmtN(avg4)} HAWBs a week over the next 4 weeks, with gross on-time near ${pctTxt(F.gross[0].mid, 0)} and net near ${pctTxt(F.net[0].mid, 1)}`,
        `Projection: trend of the last ${F.sb.length} full weeks; range covers about 8 in 10 weeks. A guide, not a commitment.`);
      const labels = [...hist.map((p) => idxLabel(p.k)), ...F.vol.out.map((_, h) => idxLabel(lastK + h + 1))];
      const cw = 7.4;
      exhibitHead(s, M, cw, 'HAWBs shipped per week, actual and projected', 'Count by ship week');
      s.addChart(pres.ChartType.bar, [{ name: 'Actual', labels, values: [...hist.map((p) => p.v[0]), ...F.vol.out.map(() => null)] },
        { name: 'Projected', labels, values: [...hist.map(() => null), ...F.vol.out.map((p) => Math.round(p.mid))] }], {
        x: M, y: 2.5, w: cw, h: 4.2, ...clean, barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 40, barOverlapPct: 100, chartColors: [PPT.navy, PPT.cyan],
        valAxisMinVal: 0, valAxisHidden: true, showLegend: true, legendPos: 't', showValue: true, dataLabelFormatCode: '#,##0', dataLabelPosition: 'outEnd', dataLabelFontSize: 8, catAxisLabelFontSize: 9 });
      const x2 = M + cw + 0.4;
      exhibitHead(s, x2, CW - cw - 0.4, 'Projected weeks', 'Likely range in brackets');
      table(s, [head(['Week of', 'HAWBs', 'Gross', 'Net']),
        ...F.vol.out.map((p, h) => [cell(idxLabel(lastK + h + 1), { align: 'left', fontSize: 10 }), cell(`${fmtN(p.mid)}\n(${fmtN(p.lo)}–${fmtN(p.hi)})`, { fontSize: 10 }),
          cell(pctTxt(F.gross[h].mid), { fontSize: 10 }), cell(pctTxt(F.net[h].mid, 2), { fontSize: 10 })])], x2, 2.55, [1.05, 1.55, 0.8, 0.93], 0.62);
    }

    // 12. definitions
    {
      const s = slide('Appendix', 'Definitions and methodology');
      const items = [
        ['Gross on-time. ', 'Delivered on or before the adjusted due date: ship date plus SLA workdays, +1 day for each origin or destination zone outside A–E, excluding holidays.'],
        ['Net on-time. ', 'A late shipment counts against net on-time only when its delay code is Controllable.'],
        ['Weeks. ', 'Weeks start Monday (Excel WEEKNUM type 2). COSTCO+ and AMAZON use ISO weeks, so Dec 29 – Jan 4 is week 1 for them. The week in progress is excluded from weekly exhibits.'],
        ['Projections. ', 'Damped linear trend over the last 12 full weeks for volume; volume-weighted average for on-time %. Ranges cover about 8 in 10 weeks.'],
        ['Data. ', `${(DATA.sources || []).join(', ')}; built ${fmtBuilt(DATA.generated)}.`],
      ];
      s.addText(items.map(([lead, txt]) => [{ text: lead, options: { bold: true, color: PPT.navy, bullet: { indent: 14 } } }, { text: txt, options: { breakLine: true } }]).flat(),
        { ...sans, x: M, y: 1.8, w: CW, h: 4.9, fontSize: 13, valign: 'top', paraSpaceAfter: 10 });
    }

    const name = ('OTP review - ' + filt).replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').slice(0, 150) + '.pptx';
    if (TEMPLATE) {
      btn.textContent = 'Applying template…';
      const date = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
      const gen = await pres.write({ outputType: 'arraybuffer' });
      downloadBlob(await mergeIntoTemplate(window.JSZip, await templateBytes(TEMPLATE), gen, ['On-time performance review', filt, `Week ${tw.week} review · ${date}`]), name);
    } else await pres.writeFile({ fileName: name });
    btn.textContent = 'Downloaded';
  } catch (err) {
    btn.textContent = 'Export failed';
    $('#notice').hidden = false; $('#notice').textContent = 'Export to PowerPoint failed: ' + (err && err.message ? err.message : err);
  } finally {
    PPT = basePPT;
    setTimeout(() => { btn.textContent = label; btn.disabled = false; }, 1800);
  }
}
$('#pptBtn').addEventListener('click', () => { if (A) return exportPpt(); });

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
  if (S.tab === 'csr') {
    choice('#csrPer', 'csrPer', [['week', 'By week'], ['month', 'By month']]);
    const byMonth = S.csrPer === 'month';
    $('#csrShipTitle').textContent = `HAWB count by CSR and ship ${byMonth ? 'month' : 'week'}`;
    $('#csrShipSub').textContent = byMonth ? 'calendar months' : 'opens at the newest week';
    csrTable();
    matrix('#mCsrShip', null, csrShipRows(), 3, 'count', { period: S.csrPer });
    segControl('#csrMetric', 'csrMetric');
    matrix('#mCsrOtp', null, A.podRows, 3, 'otp', { metric: S.csrMetric });
  }
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
