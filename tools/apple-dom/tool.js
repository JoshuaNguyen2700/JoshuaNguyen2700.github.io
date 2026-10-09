// Logic for the apple-dom tool only. It is an ES module, so nothing here leaks into other pages.
// The visitor drops the APPLE_AMR export (and, the first time, the Apple Dom Charts Template). worker.js reads them in
// the browser with model.js, writes the week's numbers into the template and builds the PowerPoint tables.
// Nothing is uploaded; the template is kept in this browser's IndexedDB so later weeks only need the AMR file.
import { esc } from '/assets/core/util.js';
import { quarterOf, weekRange } from './model.js';

const $ = (s) => document.querySelector(s);
const fmtN = (n) => (n == null ? '–' : Math.round(n).toLocaleString('en-US'));
const fmtP = (p, d = 1) => (p == null || !isFinite(p) ? '–' : (p * 100).toFixed(d) + '%');
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtD = (ms) => { const d = new Date(ms); return `${MON[d.getUTCMonth()]} ${d.getUTCDate()}`; };
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const unadj = (o) => (o.due ? 1 - o.late / o.due : null), adj = (o) => (o.due ? 1 - o.carrier / o.due : null);
const SHEET_KEY = 'apple-dom.sheet', S8_KEY = 'apple-dom.s8';

// ---------- template kept in this browser ----------
function idb() {
  return new Promise((ok, no) => { const r = indexedDB.open('apple-dom', 1); r.onupgradeneeded = () => r.result.createObjectStore('files'); r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); });
}
async function idbGet(k) { try { const db = await idb(); return await new Promise((ok) => { const q = db.transaction('files').objectStore('files').get(k); q.onsuccess = () => ok(q.result || null); q.onerror = () => ok(null); }); } catch (e) { return null; } }
async function idbSet(k, v) { try { const db = await idb(); await new Promise((ok) => { const t = db.transaction('files', 'readwrite'); t.objectStore('files').put(v, k); t.oncomplete = ok; t.onerror = ok; }); } catch (e) {} }

let TEMPLATE = null;   // { name, blob, saved }
let AMR = null;        // { name, rows, weeks }
let RES = null;        // last worker result
let SHEET = null;      // template tab shown
let S8 = 'no';         // slide 8 numbers: NO AC (as in FY27 FW1) or WITH AC (as in FY26 FW52)
try { SHEET = localStorage.getItem(SHEET_KEY); S8 = localStorage.getItem(S8_KEY) || 'no'; } catch (e) {}

function tplStatus() {
  $('#tplAsk').hidden = !!TEMPLATE;
  $('#tplStatus').innerHTML = TEMPLATE
    ? `Charts Template: <b>${esc(TEMPLATE.name)}</b>, saved in this browser ${TEMPLATE.saved ? 'on ' + esc(new Date(TEMPLATE.saved).toLocaleDateString()) : ''}. <label class="ad-linkbtn" for="tplInput">Replace</label>`
    : '';
}

// ---------- worker ----------
const worker = new Worker(new URL('worker.js', import.meta.url), { type: 'module' });
let pending = null;
const ask = (msg) => new Promise((ok, no) => { pending = { ok, no, type: msg.type }; worker.postMessage(msg); });
worker.onmessage = (e) => {
  const m = e.data;
  if (m.type === 'progress') { $('#progressBar').style.width = Math.min(100, (100 * m.done) / (m.total || 1)) + '%'; return; }
  const p = pending; pending = null; if (!p) return;
  if (m.type === 'error') p.no(new Error(m.message)); else p.ok(m);
};

function showError(where, msg) { const el = $(where); el.textContent = msg; el.hidden = !msg; }

async function handleFiles(list) {
  const files = [...list].filter((f) => /\.xlsx$/i.test(f.name));
  showError('#startError', ''); showError('#appError', '');
  if (!files.length) { showError(AMR ? '#appError' : '#startError', 'Choose .xlsx files: the APPLE_AMR export and, the first time, the Charts Template.'); return; }
  let kinds;
  try { kinds = (await ask({ type: 'classify', files })).kinds; } catch (e) { showError(AMR ? '#appError' : '#startError', e.message); return; }
  const tpl = files.find((f, i) => kinds[i] === 'template'), amrs = files.filter((f, i) => kinds[i] === 'amr');
  if (tpl) { TEMPLATE = { name: tpl.name, blob: tpl, saved: Date.now() }; await idbSet('template', { name: tpl.name, blob: tpl, saved: TEMPLATE.saved }); tplStatus(); }
  if (amrs.length) await loadAmr(amrs[0], amrs.length > 1 ? `${amrs.length} APPLE_AMR files were chosen; using ${amrs[0].name}.` : '');
  else if (AMR && tpl) await run();
}

async function loadAmr(file, note) {
  $('#start').hidden = true; $('#app').hidden = true; $('#progress').hidden = false;
  $('#progressName').textContent = `${file.name} · ${(file.size / 1048576).toFixed(1)} MB`; $('#progressBar').style.width = '0';
  try {
    AMR = await ask({ type: 'amr', file });
  } catch (e) {
    $('#progress').hidden = true; $('#start').hidden = false; AMR = null;
    showError('#startError', `${file.name} ${e.message}`); return;
  }
  $('#progress').hidden = true; $('#app').hidden = false;
  const sel = $('#weekSel');
  sel.innerHTML = AMR.weeks.list.map((w) => `<option value="${w.fy}-${w.wk}">${esc(weekLabel(w))}</option>`).join('');
  sel.value = `${AMR.weeks.def.fy}-${AMR.weeks.def.wk}`;
  $('#meta').textContent = `${AMR.name} · ${fmtN(AMR.rows)} rows`;
  RES = null; await run(note);
}

const weekLabel = (w) => { const q = quarterOf(w.fy, w.wk), [a, b] = weekRange(w.fy, w.wk); return `FW${w.wk} · Q${q.q} W${q.qweek} FY${String(w.fy).slice(2)} (${fmtD(a)} – ${fmtD(b)})`; };
const selected = () => { const [fy, wk] = $('#weekSel').value.split('-').map(Number); return { fy, wk }; };

async function run(note = '') {
  const sel = selected(), q = quarterOf(sel.fy, sel.wk);
  $('#headEyebrow').textContent = `Apple domestic OTP · FY${String(sel.fy).slice(2)} Q${q.q}`;
  $('#headTitle').textContent = `Week ${sel.wk}`;
  const [a, b] = weekRange(sel.fy, sel.wk); $('#weekDates').textContent = `Due ${fmtD(a)} – ${fmtD(b)}, quarter week ${q.qweek}`;
  $('#dlBtn').disabled = true;
  try {
    RES = await ask({ type: 'run', sel, template: TEMPLATE ? TEMPLATE.blob : null });
  } catch (e) { showError('#appError', e.message); return; }
  showError('#appError', '');
  render(note);
}

// ---------- rendering ----------
function render(note) {
  const R = RES, p = R.ppt;
  $('#dlBtn').disabled = !R.blob;
  $('#dlBtn').textContent = R.blob ? `Download ${R.fileName}` : 'Download filled template';
  const notes = [];
  if (note) notes.push(note);
  if (!R.blob) notes.push('<b>Add the Charts Template</b> to fill and preview the charts: <label class="ad-linkbtn" for="tplInput">choose the Apple Dom Charts Template .xlsx</label>. It is kept in this browser for next time.');
  else {
    const q = R.quarter;
    notes.push(`Writes <b>W${R.sel.wk - q.first + 1}</b> (FW${R.sel.wk}) only. Earlier weeks of the quarter keep what was reported in the template, because the export covers only about a month of ship dates and only the current week has its delay causes. Use last week's filled file as the template; downloading saves it for next week automatically.`);
    notes.push('Late units with no delay code are left out of the Paretos, like the pivots. OEM tabs (one Pareto) count FTL and LTL together.');
  }
  for (const w of R.warnings || []) notes.push(esc(w));
  $('#notes').innerHTML = notes.map((t) => `<p>${t}</p>`).join('');
  const o = p.s4[0], h = p.s6[0], n = p.s4[1];
  $('#kpis').innerHTML = [
    ['Units due', fmtN(o.due), 'Overall, with AC'],
    ['Unadjusted OTD', fmtP(unadj(o), 2), `${fmtN(o.late)} units late`],
    ['Adjusted OTD', fmtP(adj(o), 2), `${fmtN(o.carrier)} carrier late`, adj(o)],
    ['Adjusted OTD, no AC', fmtP(adj(n), 2), `${fmtN(n.due)} units due`, adj(n)],
    ['LTL HAWB adjusted OTD', fmtP(adj(h), 1), `${fmtN(h.due)} HAWBs due`, adj(h)],
  ].map(([k, v, d, g]) => `<div class="ktile"><div class="eyebrow">${k}</div><div class="v num${g != null ? ' ' + goalClass(g) : ''}">${v}</div><div class="d">${d}</div></div>`).join('');
  renderSheets(); renderPpt();
}

const goalClass = (v) => (v == null ? '' : v >= 0.98 ? 'g-ok' : v >= 0.93 ? 'g-warn' : 'g-bad');

function renderSheets() {
  const tabs = RES.tabs || [];
  $('#sheetChips').innerHTML = tabs.map((t) => `<button class="chip" type="button" aria-pressed="${t.name === SHEET}" data-sheet="${esc(t.name)}">${esc(t.name)}</button>`).join('');
  if (!tabs.length) { $('#otdCards').innerHTML = ''; $('#paretoCards').innerHTML = ''; return; }
  if (!tabs.some((t) => t.name === SHEET)) { SHEET = tabs[0].name; $('#sheetChips').firstElementChild.setAttribute('aria-pressed', 'true'); }
  const t = tabs.find((x) => x.name === SHEET), qw = RES.sel.wk - RES.quarter.first + 1;
  $('#otdCards').innerHTML = t.blocks.map((b, i) => `<div class="card"><div class="card-h"><h2>${b.tl} OTD</h2><span class="sub">${b.label} due by quarter week</span></div>
    <div class="legend"><span><i class="sq" style="background:var(--s1)"></i>${b.label} due</span><span><i class="sq" style="background:var(--s2)"></i>Late</span><span><i class="sq" style="background:var(--s3)"></i>Carrier late</span><span><i style="background:var(--text)"></i>Unadjusted OTD</span><span><i style="background:var(--good)"></i>Adjusted OTD</span><span><i class="dash" style="--c:var(--muted)"></i>98% goal</span></div>
    <div class="chart" id="otd${i}"></div></div>`).join('') || '<p class="sub">This tab has no OTD tables.</p>';
  t.blocks.forEach((b, i) => otdChart($('#otd' + i), b.series.slice(0, qw), b.label));
  $('#paretoCards').innerHTML = t.paretos.map((p, i) => `<div class="card"><div class="card-h"><h2>${esc(p.title)}</h2><span class="sub">Week ${RES.sel.wk} · ${fmtN(p.total)} late ${p.unit}${p.blank ? ` · ${fmtN(p.blank)} with no code left out` : ''}</span></div><div class="chart" id="par${i}"></div></div>`).join('');
  t.paretos.forEach((p, i) => paretoChart($('#par' + i), p));
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
const hideTip = () => { tip.hidden = true; };
const trow = (k, v, color) => `<div class="r"><span>${color ? `<i class="k" style="background:${color}"></i>` : ''}${k}</span><span>${v}</span></div>`;

// ---------- SVG charts ----------
const NS = 'http://www.w3.org/2000/svg';
function svgEl(tag, attrs, parent) { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; }
function niceMax(v) { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))); for (const m of [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) if (m * p >= v) return m * p; return 10 * p; }
const topRoundedBar = (x, y, w, h, r) => { r = Math.min(r, w / 2, h); return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`; };
const rightRoundedBar = (x, y, w, h, r) => { r = Math.min(r, h / 2, w); return `M${x},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h - r}Q${x + w},${y + h} ${x + w - r},${y + h}H${x}Z`; };
const T = () => ({ ink: css('--text'), muted: css('--muted'), grid: css('--a-grid'), axis: css('--a-axis'), card: css('--a-card'), s1: css('--s1'), s2: css('--s2'), s3: css('--s3'), good: css('--good') });
const kfmt = (v) => (v >= 1e6 ? (v / 1e6).toFixed(v % 1e6 ? 1 : 0) + 'M' : v >= 1000 ? (v / 1000).toFixed(v % 1000 ? 1 : 0) + 'K' : fmtN(v));

// Weekly OTD: bars for due / late / carrier late (left axis), lines for unadjusted and adjusted OTD % (right axis).
function otdChart(el, series, unit) {
  el.innerHTML = '';
  const c = T(), W = el.clientWidth || 560, H = 270, m = { l: 48, r: 48, t: 16, b: 28 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = Math.max(series.length, 1), gw = iw / n;
  const max = niceMax(Math.max(1, ...series.map((s) => s.due || 0)));
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': `${unit} due, late and OTD % by week` }, el);
  const Y = (v) => m.t + ih * (1 - v / max), P = (v) => m.t + ih * (1 - v);
  for (let t = 0; t <= 4; t++) {
    const v = (max / 4) * t;
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: t ? c.grid : c.axis, 'stroke-width': 1 }, svg);
    svgEl('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = kfmt(v);
    svgEl('text', { x: W - m.r + 6, y: P(t / 4) + 4, 'font-size': 12, fill: c.muted }, svg).textContent = t * 25 + '%';
  }
  svgEl('line', { x1: m.l, x2: W - m.r, y1: P(0.98), y2: P(0.98), stroke: c.muted, 'stroke-width': 1, 'stroke-dasharray': '4 4' }, svg);
  const bw = Math.max(3, Math.min(18, (gw * 0.7) / 3)), keys = [['due', c.s1], ['late', c.s2], ['carrier', c.s3]];
  const X = (i) => m.l + (i + 0.5) * gw;
  series.forEach((s, i) => {
    svgEl('text', { x: X(i), y: H - 8, 'text-anchor': 'middle', 'font-size': 12, fill: c.muted }, svg).textContent = 'W' + (i + 1);
    if (s.due == null) return;
    keys.forEach(([k, col], j) => { const v = s[k] || 0, x = X(i) - 1.5 * bw + j * bw; if (v > 0) svgEl('path', { d: topRoundedBar(x + 0.5, Y(v), bw - 1, ih * (v / max), 3), fill: col }, svg); });
  });
  for (const [f, col] of [[unadj, c.ink], [adj, c.good]]) {
    let d = '', last = null;
    series.forEach((s, i) => { const v = s.due ? f(s) : null; if (v == null) return; d += (d ? 'L' : 'M') + X(i).toFixed(1) + ',' + P(v).toFixed(1); last = { i, v }; svgEl('circle', { cx: X(i), cy: P(v), r: 3.5, fill: col, stroke: c.card, 'stroke-width': 1.5 }, svg); });
    if (d) svgEl('path', { d, fill: 'none', stroke: col, 'stroke-width': 2, 'stroke-linejoin': 'round' }, svg);
    if (last) svgEl('text', { x: X(last.i) + 7, y: P(last.v) + (f === adj ? -6 : 14), 'font-size': 12, 'font-weight': 600, fill: col }, svg).textContent = fmtP(last.v);
  }
  const hit = svgEl('rect', { x: m.l, y: m.t, width: iw, height: ih, fill: 'transparent' }, svg);
  hit.addEventListener('mousemove', (ev) => {
    const r = svg.getBoundingClientRect(), i = Math.max(0, Math.min(n - 1, Math.floor(((ev.clientX - r.left) * (W / r.width) - m.l) / gw))), s = series[i];
    if (!s) return;
    showTip(`<div class="tt">W${i + 1} · FW${s.w}</div>` + (s.due == null ? '<div class="tnote">No numbers yet</div>' :
      trow(unit + ' due', fmtN(s.due), c.s1) + trow('Late', fmtN(s.late), c.s2) + trow('Carrier late', fmtN(s.carrier), c.s3) + trow('Unadjusted OTD', fmtP(unadj(s), 2)) + trow('Adjusted OTD', fmtP(adj(s), 2)) +
      `<div class="tnote">${s.fromFile ? 'Written from this APPLE_AMR file' : 'As reported (from the template)'}</div>`), ev);
  });
  hit.addEventListener('mouseleave', hideTip);
}

// Delay code Pareto: horizontal bars, largest first, with the running share.
function paretoChart(el, p) {
  el.innerHTML = '';
  if (!p.items.length) { el.innerHTML = '<div class="empty-s">No late shipments with a delay code this week.</div>'; return; }
  const c = T(), W = el.clientWidth || 560, rowH = 26, lw = Math.min(260, Math.max(120, W * 0.4)), mr = 150, H = p.items.length * rowH + 6;
  const iw = W - lw - mr, max = p.items[0][1], chars = Math.floor(lw / 6.6);
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'Late volume by delay code' }, el);
  let cum = 0;
  p.items.forEach(([code, v], i) => {
    cum += v; const y = 3 + i * rowH, w = Math.max(1, (iw * v) / max), share = cum / p.total;
    const lab = svgEl('text', { x: lw - 8, y: y + rowH / 2 + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.ink }, svg);
    lab.textContent = code.length > chars ? code.slice(0, chars - 1) + '…' : code;
    svgEl('path', { d: rightRoundedBar(lw, y + 4, w, rowH - 8, 4), fill: c.s1 }, svg);
    svgEl('text', { x: lw + w + 6, y: y + rowH / 2 + 4, 'font-size': 12, 'font-weight': 600, fill: c.ink }, svg).textContent = fmtN(v);
    svgEl('text', { x: W - 4, y: y + rowH / 2 + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = fmtP(share) + ' cum.';
    const hit = svgEl('rect', { x: 0, y, width: W, height: rowH, fill: 'transparent' }, svg);
    hit.addEventListener('mousemove', (ev) => showTip(`<div class="tt">${esc(code)}</div>` + trow('Late ' + p.unit, fmtN(v)) + trow('Share', fmtP(v / p.total)) + trow('Cumulative', fmtP(share)), ev));
    hit.addEventListener('mouseleave', hideTip);
  });
}

// ---------- PowerPoint tables ----------
const pctCell = (v, d) => `<td class="num ${goalClass(v)}">${fmtP(v, d)}</td>`;
function tableCard(el, title, sub, head, rows, copies) {
  el.innerHTML = `<div class="card-h"><h2>${title}</h2><span class="sub">${sub}</span><span class="ad-copies">${copies.map((c, i) => `<button class="ad-btn" type="button" data-copy="${i}">${c.label}</button>`).join('')}</span></div>
    <div class="tscroll"><table><thead><tr>${head.map((h, i) => `<th${i ? '' : ' class="l"'}>${h}</th>`).join('')}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
  el.querySelectorAll('[data-copy]').forEach((b) => b.addEventListener('click', () => copyText(b, copies[+b.dataset.copy].text())));
}
async function copyText(btn, text) {
  const old = btn.textContent;
  try { await navigator.clipboard.writeText(text); btn.textContent = 'Copied'; } catch (e) { btn.textContent = 'Copy failed'; }
  setTimeout(() => { btn.textContent = old; }, 1400);
}
const tsv = (rows) => rows.map((r) => r.join('\t')).join('\n');

function renderPpt() {
  const p = RES.ppt, sel = RES.sel, q = RES.quarter || quarterOf(sel.fy, sel.wk), yy = String(sel.fy).slice(2);
  // Slide 4
  tableCard($('#s4'), 'Slide 4 · OTD by origin (units)', `Q${q.q} Week ${q.qweek} FY${yy}`, ['Origin', 'Units Due', 'Unadjusted OTD%', 'Adjusted OTD %'],
    p.s4.map((o) => `<tr${o.label === 'QTD' ? ' class="tot"' : ''}><td class="l">${esc(o.label)}${o.due == null ? ' <span class="sub">(needs the template)</span>' : ''}</td><td class="num">${fmtN(o.due)}</td>${pctCell(unadj(o), 2)}${pctCell(adj(o), 2)}</tr>`),
    [{ label: 'Copy', text: () => tsv(p.s4.map((o) => [o.label, fmtN(o.due), fmtP(unadj(o), 2), fmtP(adj(o), 2)])) }]);
  // Slide 6 (and the Pivot-Hawb J:O table)
  const z = (o, f) => (o.due == null ? '' : o.due ? fmtP(f(o), 1) : f === adj ? '100%' : '0');
  const pivotRows = p.s6.slice(2, -1).concat([{ ...p.s6[0], label: 'TOTAL' }, { ...p.s6[1], label: 'TOTAL(NO AC)' }]);
  tableCard($('#s6'), 'Slide 6 · OTD by HAWB (LTL)', 'Same numbers as the Pivot-Hawb J:O table', ['Origin', 'HAWB Due', 'Late', 'Carrier late', 'Unadjusted OTD%', 'Adjusted OTD %'],
    p.s6.map((o) => `<tr${o.label === 'QTD' ? ' class="tot"' : ''}><td class="l">${esc(o.label)}</td><td class="num">${fmtN(o.due)}</td><td class="num">${fmtN(o.late)}</td><td class="num">${fmtN(o.carrier)}</td>${o.due == null ? '<td class="num">–</td><td class="num">–</td>' : o.due ? pctCell(unadj(o), 1) + pctCell(adj(o), 1) : '<td class="num">0</td><td class="num g-ok">100%</td>'}</tr>`),
    [{ label: 'Copy for slide', text: () => tsv(p.s6.map((o) => [o.label, fmtN(o.due), z(o, unadj), z(o, adj)])) },
     { label: 'Copy for Pivot-Hawb', text: () => tsv(pivotRows.map((o) => [o.label, o.due, o.late, o.carrier])) }]);
  // Slide 5
  tableCard($('#s5'), 'Slide 5 · Carrier delays', `Late and Cause = Carrier, week ${sel.wk}. A1/C shows its top 10 lanes. Lanes are origin city to destination station.`, ['Delay Code with Delay Owner', 'Lane-Pair', '#Hawb', 'Units'],
    p.s5.length ? p.s5.map((r) => `<tr${r.head ? ' class="grp"' : ''}><td class="l wrap">${esc(r.label)}</td><td class="l">${esc(r.lane)}</td><td class="num">${fmtN(r.hawb)}</td><td class="num">${fmtN(r.units)}</td></tr>`) : ['<tr><td class="l" colspan="4">No carrier delays this week.</td></tr>'],
    [{ label: 'Copy', text: () => tsv(p.s5.map((r) => [r.label, r.lane, r.hawb, Math.round(r.units)])) }]);
  // Slide 8
  const s8 = p.s8[S8 === 'with' ? 'withAC' : 'noAC'], top = s8.brands.slice(0, 5);
  const sentence = s8.bkTotal ? `For BK & HB customer appointment issue, ${top.map((b) => `${(100 * b.units / s8.bkTotal).toFixed(1)}% by ${b.brand}`).join('; ')}` : 'No BK or HB delays this week.';
  const lines = s8.pareto.items.filter(([c]) => !/^(BK|HB)\b/i.test(c)).map(([c, v]) => `${Math.round(v)} ${c.replace(/^[^-]*-\s*/, '').replace(/\s*\(.*$/, '')}`);
  $('#s8').innerHTML = `<div class="card-h"><h2>Slide 8 · LTL delay comments</h2><span class="sub">LTL, week ${sel.wk}</span><span class="ad-copies"><span class="chips">${[['no', 'NO AC'], ['with', 'WITH AC']].map(([k, l]) => `<button class="chip" type="button" data-s8="${k}" aria-pressed="${S8 === k}">${l}</button>`).join('')}</span><button class="ad-btn" type="button" id="copy8">Copy</button></span></div>
    <p class="ad-quote">${esc(sentence)}</p>
    <ul class="ad-lines">${lines.map((l) => `<li class="num">${esc(l)}</li>`).join('')}</ul>
    <details><summary>BK &amp; HB consignees (${fmtN(s8.bkTotal)} late units)</summary><div class="tscroll"><table><thead><tr><th class="l">Brand / consignee</th><th>Units</th><th>Share</th></tr></thead><tbody>${s8.brands.map((b) =>
      `<tr class="grp"><td class="l">${esc(b.brand)}</td><td class="num">${fmtN(b.units)}</td><td class="num">${fmtP(b.units / s8.bkTotal)}</td></tr>` + b.names.map(([nm, v]) => `<tr><td class="l sub2">${esc(nm)}</td><td class="num">${fmtN(v)}</td><td class="num">${fmtP(v / s8.bkTotal)}</td></tr>`).join('')).join('')}</tbody></table></div></details>`;
  $('#copy8').addEventListener('click', (e) => copyText(e.currentTarget, [sentence, ...lines].join('\n')));
  $('#s8').querySelectorAll('[data-s8]').forEach((b) => b.addEventListener('click', () => { S8 = b.dataset.s8; try { localStorage.setItem(S8_KEY, S8); } catch (err) {} renderPpt(); }));
}

// ---------- wiring ----------
$('#fileInput').addEventListener('change', (e) => { handleFiles(e.target.files); e.target.value = ''; });
$('#tplInput').addEventListener('change', async (e) => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  let kind; try { kind = (await ask({ type: 'classify', files: [f] })).kinds[0]; } catch (err) { kind = null; }
  if (kind !== 'template') { showError(AMR ? '#appError' : '#startError', `${f.name} is not the Apple Dom Charts Template (it has no "Overall" tabs).`); return; }
  TEMPLATE = { name: f.name, blob: f, saved: Date.now() }; await idbSet('template', TEMPLATE); tplStatus();
  if (AMR) run();
});
$('#weekSel').addEventListener('change', () => run());
$('#dlBtn').addEventListener('click', () => {
  if (!RES || !RES.blob) return;
  const a = document.createElement('a'), url = URL.createObjectURL(RES.blob);
  a.href = url; a.download = RES.fileName; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  // the filled file becomes the template for next week, so the quarter's history carries forward
  TEMPLATE = { name: RES.fileName, blob: RES.blob, saved: Date.now() }; idbSet('template', TEMPLATE); tplStatus();
  const n = $('#dlNote'); n.textContent = `Saved ${RES.fileName} in this browser as the template for next week.`; n.hidden = false;
});
$('#sheetChips').addEventListener('click', (e) => {
  const b = e.target.closest('[data-sheet]'); if (!b) return;
  SHEET = b.dataset.sheet; try { localStorage.setItem(SHEET_KEY, SHEET); } catch (err) {}
  renderSheets();
});
document.querySelectorAll('.ad-tabs .tab').forEach((t) => t.addEventListener('click', () => {
  document.querySelectorAll('.ad-tabs .tab').forEach((x) => x.setAttribute('aria-selected', String(x === t)));
  $('#panel-charts').hidden = t.dataset.tab !== 'charts'; $('#panel-ppt').hidden = t.dataset.tab !== 'ppt';
  if (t.dataset.tab === 'charts' && RES) renderSheets();
}));
// drop anywhere on the page
let depth = 0;
addEventListener('dragenter', (e) => { if (e.dataTransfer?.types?.includes('Files')) { depth++; document.body.classList.add('dragging'); } });
addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; document.body.classList.remove('dragging'); } });
addEventListener('dragover', (e) => e.preventDefault());
addEventListener('drop', (e) => { e.preventDefault(); depth = 0; document.body.classList.remove('dragging'); if (e.dataTransfer?.files?.length) handleFiles(e.dataTransfer.files); });
let rt; addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => { if (RES && !$('#panel-charts').hidden) renderSheets(); }, 150); });

(async () => { const t = await idbGet('template'); if (t && t.blob) TEMPLATE = t; tplStatus(); })();
