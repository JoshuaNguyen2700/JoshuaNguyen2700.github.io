// Logic for the case-count tool only. It is an ES module, so nothing here leaks into other pages.
// The page ships with no case data. The visitor picks the Cview Report folder (team folders of Salesforce
// exports, plus the ACTIONED & CLOSED CASE COUNT workbook); worker.js reads them in the browser and
// model.js turns the counts into the workbook's tables. Nothing is uploaded.
// Every filter applies to every tab: supervisor and team folder (chips), people and customers
// (multi-select lists; nothing ticked = all), and dates.
import { esc } from '/assets/core/util.js';
import {
  GOAL, DOW, SHEET_MON, MON, dt, dowOf, md, dLong, span, isWeekday, monthKey, monthStart, monthEnd, byText,
  prepare, buildLayout, unlisted, casesModel, custModel, summaryModel, daily, topCustomers, visibleBlocks, col, cS, cN,
} from './model.js';

const $ = (s) => document.querySelector(s);
const STATE_KEY = 'cc.state';   // the open tab and the weekends choice only; never data

// ---------- formatting ----------
const toInput = (d) => dt(d).toISOString().slice(0, 10);
const fromInput = (s) => (/^\d{4}-\d{2}-\d{2}$/.test(s) ? Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) / 864e5 : null);
const fmtN = (n) => Math.round(n).toLocaleString('en-US');
const fmtP = (p, d = 0) => (p == null || !isFinite(p) ? '–' : (p * 100).toFixed(d) + '%');
const plural = (n, w) => `${fmtN(n)} ${w}${n === 1 ? '' : 's'}`;
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

// ---------- data + state ----------
// M: prepare(D) (data, indexes, date extent X); L: the layout (sections of people, CUST E-MAILS blocks)
let M = null, L = null, S = null, worker = null;
const OVR = new Map();   // name fixes made on the Names tab: row key -> person (this page only)
const saved = (() => { try { return JSON.parse(localStorage.getItem(STATE_KEY)) || {}; } catch (e) { return {}; } })();
const saveView = () => { try { localStorage.setItem(STATE_KEY, JSON.stringify({ tab: S.tab, weekends: S.weekends })); } catch (e) {} };
const TABS = ['over', 'cases', 'cust', 'names'];
const dflt = () => ({ secs: new Set(L.sections.map((s) => s.name)), teams: new Set(M.D.teams.map((_, i) => i)), people: new Set(), cust: new Set(), from: M.X.lo, to: M.X.hi });

function load(data) {
  M = prepare(data); OVR.clear(); L = buildLayout(M, OVR);
  // after adding or removing files, stay on the tab that was open
  const tab = S ? S.tab : saved.tab, weekends = S ? S.weekends : !!saved.weekends;
  S = { ...dflt(), tab: TABS.includes(tab) ? tab : 'over', weekends };
  const usable = M.D.files.filter((f) => f.kind && f.kind !== 'workbook'), skipped = M.D.files.filter((f) => !f.kind);
  $('#meta').textContent = `${plural(usable.length, 'report')} · ${span(M.X.lo, M.X.hi)}`;
  const KN = { sent: 'sent emails (cases actioned)', received: 'received emails', closed: 'closed cases' }, notes = [];
  const miss = Object.keys(KN).filter((k) => !usable.some((f) => f.kind === k));
  if (!L.fromBook) notes.push('The ACTIONED & CLOSED CASE COUNT workbook wasn\'t in the files, so people are grouped by team folder with their full names. Load the Cview Report folder with the workbook in it (closed in Excel) to get the supervisor sections, names and customer rows.');
  if (miss.length) notes.push(`No ${miss.map((k) => KN[k]).join(', ')} report was found, so those figures are 0.`);
  if (skipped.length) notes.push(`Skipped: ${skipped.map((f) => `${f.name.split('/').pop()} (${f.note})`).join(', ')}.`);
  const missing = L.fromBook ? L.sections.flatMap((s) => s.rows.filter((r) => r.p < 0).map((r) => `${r.label} (${s.name})`)) : [];
  if (missing.length) notes.push(`Not in the loaded exports, so blank: ${missing.join(', ')}.`);
  $('#notice').hidden = !notes.length; $('#notice').textContent = notes.join(' ');
  buildFilters(); show('app'); setTab(S.tab);
}

// ---------- screens + loading ----------
function show(view) { for (const v of ['start', 'progress', 'app']) $('#' + v).hidden = v !== view; if (view !== 'app') hideTip(); }
function showStartError(msg) { const e = $('#startError'); e.textContent = msg; e.hidden = !msg; if (msg) show(M ? 'app' : 'start'); }
const REPORT = /\.(xls|xlsx|xlsm|csv)$/i;
const reports = (items) => items.filter((it) => REPORT.test(it.file.name) && !/^~\$/.test(it.file.name));
const NO_REPORTS = 'No report files in that folder. Choose a team folder with the .xls exports (East, Legacy, …) or the Cview Report folder that holds them.';
// Folders are loaded only as folders: on the start screen each one picked (or dropped) joins a list, and
// Load reads them all together. Choosing the parent folder brings every folder inside it at once. Once
// loaded, Add folder (or a drop) adds more and re-reads the whole set, in this page's memory only.
// The same file picked twice counts once.
let LOADED = new Map(), PENDING = new Map();
const itemKey = (it) => `${it.file.name}|${it.file.size}|${it.file.lastModified}`;
const topFolder = (it) => { const p = it.path.split('/'); return p.length > 1 ? p[0] : 'Loose files'; };
function stage(items) {
  items = reports(items);
  if (!items.length) { showStartError(NO_REPORTS); return; }
  showStartError('');
  for (const it of items) { const k = itemKey(it); if (!PENDING.has(k)) PENDING.set(k, it); }
  renderPending();
}
function renderPending() {
  const groups = new Map();
  for (const it of PENDING.values()) { const f = topFolder(it); groups.set(f, [...(groups.get(f) || []), it]); }
  const n = groups.size, book = [...PENDING.values()].some((it) => /\.xls[xm]$/i.test(it.file.name));
  $('#pending').hidden = !n; $('#pickRow').hidden = !!n;
  $('#pendingList').innerHTML = [...groups].map(([f, its]) => {
    const subs = [...new Set(its.map((it) => it.path.split('/').slice(1, -1)[0]).filter(Boolean))];
    return `<li><b>${esc(f)}</b><span>${plural(its.length, 'file')}${subs.length ? ` in ${esc(subs.join(', '))}` : ''}</span><button class="cc-x" type="button" data-unstage="${esc(f)}">Remove</button></li>`;
  }).join('');
  $('#loadBtn').textContent = n === 1 ? 'Load this folder' : `Load ${n} folders`;
  $('#pendingNote').hidden = book || !n;
}
$('#pendingList').addEventListener('click', (e) => {
  const b = e.target.closest('[data-unstage]'); if (!b) return;
  for (const [k, it] of PENDING) if (topFolder(it) === b.dataset.unstage) PENDING.delete(k);
  renderPending();
});
$('#loadBtn').addEventListener('click', () => { if (!PENDING.size) return; LOADED = new Map(PENDING); PENDING = new Map(); renderPending(); readFiles([...LOADED.values()]); });
// After the first load: add more folders to what's loaded and read everything again.
function addItems(items) {
  items = reports(items);
  if (!items.length) { $('#notice').hidden = false; $('#notice').textContent = NO_REPORTS; return; }
  for (const it of items) { const k = itemKey(it); if (!LOADED.has(k)) LOADED.set(k, it); }
  readFiles([...LOADED.values()]);
}
// The files in the last read, in the order the worker lists them (two files can share a name).
let READ = [];
function readFiles(items) {
  READ = items;
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
  const KIND = { sent: 'emails sent', received: 'emails received', closed: 'cases closed' };
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
$('#folderInput').addEventListener('change', (e) => { stage(fromInputEl(e.target.files)); e.target.value = ''; });
$('#addFolderInput').addEventListener('change', (e) => { addItems(fromInputEl(e.target.files)); e.target.value = ''; });
$('#cancelBtn').addEventListener('click', () => {
  if (worker) { worker.terminate(); worker = null; }
  if (!M) { PENDING = new Map(LOADED); renderPending(); }   // back to the list, ready to load again
  show(M ? 'app' : 'start');
});
$('#closeBtn').addEventListener('click', () => { M = null; L = null; LOADED = new Map(); PENDING = new Map(); renderPending(); show('start'); });

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
  dropItems(e.dataTransfer).then((items) => (M ? addItems(items) : stage(items)), (err) => showStartError('Those folders could not be opened: ' + err.message));
});

// ---------- multi-select list ----------
// A button that opens a checklist with a search box, Select all / Clear, and group headers that tick
// their whole group. Nothing ticked means everything, so the view is never empty by accident.
function multiSelect(root, opts) {
  root.innerHTML = `<button class="ms-btn" type="button" aria-haspopup="true" aria-expanded="false" aria-labelledby="${opts.labelId}"></button>
    <div class="ms-pop" hidden><input type="search" placeholder="Search" aria-label="Search ${esc(opts.plural)}">
    <div class="ms-acts"><button class="cc-linkbtn" type="button" data-act="all">Select all shown</button><button class="cc-linkbtn" type="button" data-act="none">Clear</button></div>
    <div class="ms-list" role="group" aria-labelledby="${opts.labelId}"></div></div>`;
  const btn = root.querySelector('.ms-btn'), pop = root.querySelector('.ms-pop'), list = root.querySelector('.ms-list'), q = root.querySelector('input');
  let groups = [];
  const sel = () => opts.selected();
  const shown = () => { const t = q.value.trim().toLowerCase(); return groups.map((g) => ({ ...g, items: g.items.filter((it) => !t || (it.label + ' ' + (it.sub || '') + ' ' + g.label).toLowerCase().includes(t)) })).filter((g) => g.items.length); };
  const label = () => {
    const s = sel(), all = groups.flatMap((g) => g.items);
    const on = all.filter((it) => s.has(it.id));
    btn.textContent = !on.length ? `All ${opts.plural}` : on.length === 1 ? on[0].label : `${on.length} ${opts.plural}`;
    btn.classList.toggle('on', on.length > 0);
    btn.title = on.length ? on.map((it) => it.label).join(', ') : `Every ${opts.noun} in the other filters`;
  };
  const draw = () => {
    const s = sel(), gs = shown();
    list.innerHTML = gs.length ? gs.map((g, gi) => {
      const n = g.items.filter((it) => s.has(it.id)).length;
      return (g.label ? `<label class="g"><input type="checkbox" data-g="${gi}"${n && n === g.items.length ? ' checked' : ''}><span>${esc(g.label)}</span></label>` : '') +
        g.items.map((it) => `<label class="${g.label ? 'it' : ''}" title="${esc(it.title || it.label)}"><input type="checkbox" data-id="${esc(it.id)}"${s.has(it.id) ? ' checked' : ''}><span>${esc(it.label)}</span>${it.sub ? `<small>${esc(it.sub)}</small>` : ''}</label>`).join('');
    }).join('') : `<div class="ms-empty">Nothing matches.</div>`;
    sync();
  };
  // Tick marks and group states updated in place, so focus and scroll stay put while ticking.
  const sync = () => {
    const s = sel(), gs = shown();
    list.querySelectorAll('input[data-id]').forEach((x) => { x.checked = s.has(x.dataset.id); });
    gs.forEach((g, gi) => {
      const b = list.querySelector(`[data-g="${gi}"]`), n = g.items.filter((it) => s.has(it.id)).length;
      if (b) { b.checked = n > 0 && n === g.items.length; b.indeterminate = n > 0 && n < g.items.length; }
    });
    label();
  };
  const open = (on) => { pop.hidden = !on; btn.setAttribute('aria-expanded', String(on)); if (on) { q.value = ''; draw(); q.focus(); } };
  btn.addEventListener('click', () => open(pop.hidden));
  q.addEventListener('input', draw);
  list.addEventListener('change', (e) => {
    const s = sel(), x = e.target;
    if (x.dataset.g != null) { const g = shown()[+x.dataset.g]; for (const it of g.items) { if (x.checked) s.add(it.id); else s.delete(it.id); } }
    else if (x.dataset.id != null) { if (x.checked) s.add(x.dataset.id); else s.delete(x.dataset.id); }
    sync(); opts.onChange();
  });
  root.querySelector('.ms-acts').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    const s = sel();
    if (b.dataset.act === 'all') for (const g of shown()) for (const it of g.items) s.add(it.id);
    else s.clear();
    sync(); opts.onChange();
  });
  document.addEventListener('click', (e) => { if (!pop.hidden && !root.contains(e.target)) open(false); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !pop.hidden) { open(false); btn.focus(); } });
  return { setGroups(g) { groups = g; const ids = new Set(g.flatMap((x) => x.items.map((it) => it.id))); for (const id of [...sel()]) if (!ids.has(id)) sel().delete(id); if (!pop.hidden) draw(); else label(); } };
}
const msPeople = multiSelect($('#msPeople'), { noun: 'person', plural: 'people', labelId: 'lPeople', selected: () => S.people, onChange: () => { fillCustomers(); render(); } });
const msCust = multiSelect($('#msCust'), { noun: 'customer', plural: 'customers', labelId: 'lCust', selected: () => S.cust, onChange: () => render() });

// ---------- filters ----------
const teamOfRow = (r) => (r.p >= 0 ? M.I.pteam.get(r.p) : null);
// rows allowed by the supervisor and team chips (before the people list)
const inScope = (s, r) => S.secs.has(s.name) && (S.teams.size === M.D.teams.length || (r.p >= 0 && S.teams.has(teamOfRow(r))));
function fillPeople() {
  msPeople.setGroups(L.sections.filter((s) => S.secs.has(s.name)).map((s) => ({ label: L.sections.length > 1 ? s.name : '',
    items: s.rows.filter((r) => inScope(s, r)).map((r) => ({ id: r.key, label: r.label, sub: r.p >= 0 && r.label !== M.D.people[r.p] ? M.D.people[r.p] : r.p < 0 ? 'no data' : '', title: r.p >= 0 ? M.D.people[r.p] : 'Not in the loaded exports' })) }))
    .filter((g) => g.items.length));
}
function fillCustomers() {
  const rows = new Set(filterRows());
  msCust.setGroups(L.cust.filter((b) => b.rowKey && rows.has(b.rowKey)).map((b) => ({ label: b.label,
    items: b.customers.map((c) => ({ id: c.id, label: b.mailbox ? b.label : c.label, title: c.comps.length ? c.comps.map((x) => M.D.companies[x]).join(', ') : 'No Salesforce company matched', sub: c.comps.length ? '' : 'no match' })) })));
}
function buildFilters() {
  $('#secChips').innerHTML = L.sections.map((s) => `<button class="chip" type="button" data-s="${esc(s.name)}" aria-pressed="${S.secs.has(s.name)}">${esc(s.name)}</button>`).join('');
  const multiTeam = M.D.teams.length > 1;
  $('#teamLabel').hidden = $('#teamChips').hidden = !multiTeam;
  $('#teamChips').innerHTML = M.D.teams.map((t, i) => [t, i]).sort((a, b) => byText(a[0], b[0]))
    .map(([t, i]) => `<button class="chip" type="button" data-t="${i}" aria-pressed="${S.teams.has(i)}">${esc(t)}</button>`).join('');
  for (const id of ['#fFrom', '#fTo']) { $(id).min = toInput(M.X.lo); $(id).max = toInput(M.X.hi); }
  $('#fFrom').value = toInput(S.from); $('#fTo').value = toInput(S.to);
  $('#fWeekends').checked = S.weekends;
  fillPeople(); fillCustomers(); quickDates();
}
// Quick ranges: everything loaded, each month the exports cover, and the latest Monday-to-Friday week.
function quickRanges() {
  const out = [['All dates', M.X.lo, M.X.hi]];
  for (let k = monthKey(M.X.lo); k <= monthKey(M.X.hi); k++) out.push([`${MON[k % 12]} ${Math.floor(k / 12)}`, Math.max(monthStart(k), M.X.lo), Math.min(monthEnd(k), M.X.hi)]);
  const mon = M.X.hi - ((M.X.hi + 3) % 7);
  out.push(['Latest week', Math.max(mon, M.X.lo), Math.min(mon + 6, M.X.hi)]);
  return out.length === 3 && out[1][1] === M.X.lo && out[1][2] === M.X.hi ? [out[0], out[2]] : out;
}
function quickDates() {
  $('#quick').innerHTML = quickRanges().map(([l, a, b]) => `<button type="button" data-a="${a}" data-b="${b}" aria-pressed="${S.from === a && S.to === b}">${esc(l)}</button>`).join('');
}
$('#quick').addEventListener('click', (e) => {
  const b = e.target.closest('button'); if (!b) return;
  S.from = +b.dataset.a; S.to = +b.dataset.b; $('#fFrom').value = toInput(S.from); $('#fTo').value = toInput(S.to); quickDates(); render();
});
const toggleChip = (set, v, all) => { if (set.has(v)) set.delete(v); else set.add(v); if (!set.size) for (const x of all) set.add(x); };
$('#secChips').addEventListener('click', (e) => {
  const b = e.target.closest('.chip'); if (!b) return;
  toggleChip(S.secs, b.dataset.s, L.sections.map((s) => s.name));
  document.querySelectorAll('#secChips .chip').forEach((c) => c.setAttribute('aria-pressed', String(S.secs.has(c.dataset.s))));
  fillPeople(); fillCustomers(); render();
});
$('#teamChips').addEventListener('click', (e) => {
  const b = e.target.closest('.chip'); if (!b) return;
  toggleChip(S.teams, +b.dataset.t, M.D.teams.map((_, i) => i));
  document.querySelectorAll('#teamChips .chip').forEach((c) => c.setAttribute('aria-pressed', String(S.teams.has(+c.dataset.t))));
  fillPeople(); fillCustomers(); render();
});
$('#fFrom').addEventListener('change', (e) => { const d = fromInput(e.target.value); S.from = d == null ? M.X.lo : Math.max(M.X.lo, d); if (S.from > S.to) { S.to = S.from; $('#fTo').value = toInput(S.to); } quickDates(); render(); });
$('#fTo').addEventListener('change', (e) => { const d = fromInput(e.target.value); S.to = d == null ? M.X.hi : Math.min(M.X.hi, d); if (S.to < S.from) { S.from = S.to; $('#fFrom').value = toInput(S.from); } quickDates(); render(); });
$('#fWeekends').addEventListener('change', (e) => { S.weekends = e.target.checked; saveView(); render(); });
$('#resetBtn').addEventListener('click', () => { Object.assign(S, dflt()); buildFilters(); render(); });

function filterRows() {
  const out = [];
  for (const s of L.sections) for (const r of s.rows) if (inScope(s, r) && (!S.people.size || S.people.has(r.key))) out.push(r.key);
  return out;
}
// F: what every model reads (see model.js)
function filters() {
  const days = [];
  for (let d = S.from; d <= S.to; d++) if (S.weekends || isWeekday(d)) days.push(d);
  return { rows: new Set(filterRows()), days, cust: S.cust.size ? new Set(S.cust) : null };
}

// ---------- tabs ----------
function setTab(t) {
  S.tab = TABS.includes(t) ? t : 'over';
  document.querySelectorAll('.cc-tabs .tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === S.tab)));
  document.querySelectorAll('.panel').forEach((p) => (p.hidden = p.id !== 'p-' + S.tab));
  saveView(); render();
}
$('.cc-tabs').addEventListener('click', (e) => { const b = e.target.closest('.tab'); if (b) setTab(b.dataset.tab); });

// ---------- tables ----------
function cellHtml(c) {
  if (c.f === 's') return esc(c.v ?? '');
  if (c.v == null || (c.f === 'n' && !c.v)) return c.f === 'p' ? '–' : '';
  return c.f === 'p' ? fmtP(c.v) : fmtN(c.v);
}
const emptyMsg = (t) => `<div class="empty-s">${t}</div>`;
function renderTable(el, m) {
  if (!el) return;
  if (!m.rows.length || (m.n === 0)) { el.innerHTML = emptyMsg(m.empty || 'Nothing for these filters.'); return; }
  const th = m.cols.map((c) => `<th class="${c.cls || ''}">${c.h}</th>`).join('');
  const body = m.rows.map((r) => `<tr class="${r.cls || ''}">${r.c.map((c, i) => `<td class="${[m.cols[i] && m.cols[i].cls, c.cls].filter(Boolean).join(' ')}"${c.title ? ` title="${esc(c.title)}"` : ''}>${cellHtml(c)}</td>`).join('')}</tr>`).join('');
  el.innerHTML = `<table class="${m.cls || ''}"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
}

// ---------- overview ----------
function tile(label, value, sub) { return `<div class="ktile"><div class="eyebrow">${label}</div><div class="v">${value}</div><div class="d">${sub}</div></div>`; }
function renderTiles(series) {
  const t = series.reduce((a, x) => ({ act: a.act + x.act, clo: a.clo + x.clo, recv: a.recv + x.recv, worked: a.worked + x.worked }), { act: 0, clo: 0, recv: 0, worked: 0 });
  const n = series.length, p = t.recv ? t.worked / t.recv : null;
  const pill = p == null ? 'no customer emails in these filters' : `<span class="pill ${p >= GOAL ? 'ok' : 'bad'}">${p >= GOAL ? 'meets' : 'below'} the 85% goal</span>`;
  $('#tiles').innerHTML =
    tile('Cases actioned', fmtN(t.act), n ? `${fmtN(t.act / n)} per day · emails sent on the cases` : '') +
    tile('Cases closed', fmtN(t.clo), n ? `${fmtN(t.clo / n)} per day` : '') +
    tile('Actioned + closed', fmtN(t.act + t.clo), plural(n, S.weekends ? 'day' : 'weekday')) +
    tile('Customer emails', fmtN(t.recv), 'received for the CUST E-MAILS reps') +
    tile('% actioned', fmtP(p), pill);
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
// Stacked columns per day (actioned, closed on top) and a line for customer emails received.
function dailyChart(el, series) {
  el.innerHTML = '';
  if (!series.length) { el.innerHTML = emptyMsg('No days in this date range.'); return; }
  const c = { muted: css('--muted'), grid: css('--o-grid'), axis: css('--o-axis'), s1: css('--s1'), s2: css('--s2'), s3: css('--s3') };
  const W = el.clientWidth || 600, H = 260, m = { l: 46, r: 12, t: 14, b: 28 }, iw = W - m.l - m.r, ih = H - m.t - m.b, n = series.length, bw = iw / n;
  const max = niceMax(Math.max(...series.map((x) => Math.max(x.act + x.clo, x.recv))));
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H, role: 'img', 'aria-label': 'Cases actioned and closed per day, and customer emails received' }, el);
  const Y = (v) => m.t + ih * (1 - v / max), Xc = (i) => m.l + (i + 0.5) * bw;
  for (let k = 0; k <= 4; k++) {
    const v = (max / 4) * k;
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: k === 0 ? c.axis : c.grid, 'stroke-width': 1 }, svg);
    svgEl('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: c.muted }, svg).textContent = fmtN(v);
  }
  const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(iw / 44)))), bar = Math.max(3, Math.min(30, bw - 4)), groups = [];
  series.forEach((x, i) => {
    const gEl = svgEl('g', { opacity: isWeekday(x.d) ? 1 : 0.55 }, svg), left = Xc(i) - bar / 2;
    if (x.act) svgEl('rect', { x: left, y: Y(x.act), width: bar, height: Math.max(0.5, ih * x.act / max), fill: c.s1 }, gEl);
    if (x.clo) svgEl('rect', { x: left, y: Y(x.act + x.clo), width: bar, height: Math.max(0.5, ih * x.clo / max), fill: c.s3 }, gEl);
    groups.push(gEl);
    if (i % every === 0) svgEl('text', { x: Xc(i), y: H - 8, 'text-anchor': 'middle', 'font-size': 12, fill: c.muted }, svg).textContent = md(x.d);
  });
  if (series.some((x) => x.recv)) {
    svgEl('path', { d: series.map((x, i) => (i ? 'L' : 'M') + Xc(i).toFixed(1) + ',' + Y(x.recv).toFixed(1)).join(''), fill: 'none', stroke: c.s2, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }, svg);
    if (n <= 40) series.forEach((x, i) => svgEl('circle', { cx: Xc(i), cy: Y(x.recv), r: 3, fill: c.s2 }, svg));
  }
  const hit = svgEl('rect', { x: m.l, y: m.t, width: iw, height: ih, fill: 'transparent' }, svg);
  hit.addEventListener('mousemove', (ev) => {
    const r = svg.getBoundingClientRect(), i = Math.max(0, Math.min(n - 1, Math.floor(((ev.clientX - r.left) * (W / r.width) - m.l) / bw))), x = series[i];
    groups.forEach((gEl, j) => gEl.setAttribute('opacity', j === i ? 1 : 0.45));
    showTip(`<div class="tt">${DOW[dowOf(x.d)]} ${dLong(x.d)}</div>` + tipRow('Cases actioned', fmtN(x.act), c.s1) + tipRow('Cases closed', fmtN(x.clo), c.s3) +
      tipRow('Actioned + closed', fmtN(x.act + x.clo)) + tipRow('Customer emails', fmtN(x.recv), c.s2) + tipRow('% actioned', fmtP(x.recv ? x.worked / x.recv : null)), ev);
  });
  hit.addEventListener('mouseleave', () => { groups.forEach((gEl, j) => gEl.setAttribute('opacity', isWeekday(series[j].d) ? 1 : 0.55)); hideTip(); });
}
function renderTop(F) {
  const top = topCustomers(M, L, F);
  if (!top.length) { $('#topCust').innerHTML = emptyMsg('No customer emails for these filters.'); return; }
  const max = top[0].n;
  $('#topCust').innerHTML = top.map((x) => `<div class="b" title="${esc(x.label)} (${esc(x.rep)}): ${fmtN(x.n)} emails"><span>${esc(x.label)}${x.rep !== x.label ? `<small>${esc(x.rep)}</small>` : ''}</span>
    <span class="track"><i style="width:${(100 * x.n) / max}%"></i></span><span>${fmtN(x.n)}</span></div>`).join('');
}

// ---------- cases actioned & closed (the month tab) ----------
function renderCases(F) {
  const secs = L.sections.filter((s) => s.rows.some((r) => F.rows.has(r.key)));
  if (!secs.length || !F.days.length) { $('#p-cases').innerHTML = `<div class="card">${emptyMsg('Nobody matches these filters.')}</div>`; return; }
  $('#p-cases').innerHTML = `<p class="note">Like the month tab: one column per ${S.weekends ? 'day' : 'weekday'}, a TOTAL after each week${F.days.length > 5 ? ' and one for all the days shown' : ''}. Hover a name to see who it is in Salesforce. PTO isn't in Salesforce, so a day off shows 0.</p>` +
    secs.map((s, j) => `<div class="card"><div class="card-h"><h2>${esc(s.name)}</h2></div>
      ${[0, 1].map((i) => `<div class="card-h"><span class="sub">${i ? 'CASES CLOSED' : 'CASES ACTIONED'}</span><button class="cc-linkbtn copybtn" type="button" data-copy="#mt-${j}-${i}" title="Copy this table, then paste into Excel">Copy</button></div><div class="tscroll" id="mt-${j}-${i}"></div>`).join('')}
    </div>`).join('');
  secs.forEach((s, j) => [0, 1].forEach((i) => renderTable($(`#mt-${j}-${i}`), casesModel(M, s, i, F))));
}

// ---------- customer emails (CUST E-MAILS) ----------
function renderCust(F) {
  const n = visibleBlocks(L, F).length;
  $('#custSub').textContent = `${plural(n, 'rep')}; goal 85%. Hover a customer to see the Salesforce companies counted.${F.cust ? ' TOTAL and % ACTIONED cover all of the rep\'s customers, as in the workbook.' : ''}`;
  renderTable($('#mCust'), custModel(M, L, F));
}

// ---------- names & sources ----------
function renderNames() {
  const opts = (p) => `<option value="-1">(none)</option>` + M.D.people.map((n, i) => [n, i]).sort((a, b) => byText(a[0], b[0])).map(([n, i]) => `<option value="${i}"${i === p ? ' selected' : ''}>${esc(n)}</option>`).join('');
  $('#namesNote').innerHTML = L.fromBook
    ? `Names come from the <b>${esc(L.month)}</b> tab. Each is matched to a Salesforce case owner by name or nickname, or by comparing the workbook's own numbers with the exports. Pick a different person to fix a match; fixes last until the page is closed. To keep them, add a <b>ROSTER</b> tab to the workbook with the workbook name in column A and the Salesforce name in column B.`
    : 'The workbook wasn\'t loaded, so each team folder is a section and full Salesforce names are used.';
  $('#tNames').innerHTML = `<table><thead><tr><th>Supervisor</th><th>Name in workbook</th><th>Salesforce case owner</th><th>Team folder</th><th>Matched by</th></tr></thead><tbody>` +
    L.sections.flatMap((s) => s.rows.map((r) => `<tr><td>${esc(s.name)}</td><td>${esc(r.label)}</td><td>${L.fromBook ? `<select data-key="${esc(r.key)}" aria-label="Salesforce name for ${esc(r.label)}">${opts(r.p)}</select>` : esc(M.D.people[r.p])}</td><td>${esc(r.p >= 0 ? M.D.teams[M.I.pteam.get(r.p)] || '' : '')}</td><td class="${r.p < 0 || /check/.test(r.how) ? 'low' : ''}">${esc(r.how)}</td></tr>`)).join('') + '</tbody></table>';
  $('#tCustMap').innerHTML = L.cust.length ? `<table><thead><tr><th>Rep</th><th>Customer row</th><th>Salesforce companies counted</th></tr></thead><tbody>` +
    L.cust.flatMap((b) => b.customers.map((c) => `<tr><td>${esc(b.label)}</td><td>${esc(c.label)}</td><td class="${c.comps.length ? '' : 'low'}" style="white-space:normal;text-align:left">${c.comps.length ? esc(c.comps.map((x) => M.D.companies[x]).join(' · ')) : 'none matched'}</td></tr>`)).join('') + '</tbody></table>' : emptyMsg('No customer rows.');
  const un = unlisted(M, L);
  $('#tUnlisted').innerHTML = un.length ? `<table><thead><tr><th>Case owner</th><th>Team folder</th><th style="text-align:right">Actioned + closed</th></tr></thead><tbody>` +
    un.map((x) => `<tr><td>${esc(x.name)}</td><td>${esc(x.team)}</td><td style="text-align:right">${fmtN(x.n)}</td></tr>`).join('') + '</tbody></table>' : emptyMsg('Everyone in the exports has a row.');
  about();
}
$('#p-names').addEventListener('change', (e) => {
  const s = e.target.closest('select[data-key]'); if (!s) return;
  OVR.set(s.dataset.key, +s.value); L = buildLayout(M, OVR); fillPeople(); fillCustomers(); renderNames();
});
const KIND = { sent: 'Emails sent (cases actioned)', received: 'Emails received', closed: 'Cases closed', workbook: 'Workbook layout' };
function about() {
  $('#about').innerHTML = `<h2>How the numbers are calculated</h2>
    <ul>
      <li><b>CASES ACTIONED</b>: every row of the sent-emails report (Sent, Replied and Draft) on the cases a person owns, on the email's date.</li>
      <li><b>CASES CLOSED</b>: cases whose Date/Time Closed falls on that day, credited to the case's current owner. Salesforce keeps only a case's latest close, so closures that were later reopened, or cases now owned by a queue, aren't counted. That's why closed can run lower than numbers typed into the workbook on the day.</li>
      <li><b>TOTAL</b> columns add up each Monday-to-Friday week; <b>TOTALS</b> rows add up everyone shown.</li>
      <li><b>CUST E-MAILS</b>: each customer row counts every email received for the Salesforce companies matched to it (listed above), whoever owns the case. <b>E-MAILS FROM CEVA STATIONS</b> are emails from @cevalogistics.com addresses on the rep's cases. <b>TOTAL</b> = customer rows + CEVA stations. <b>E-MAILS ACTIONED</b> = the rep's cases actioned + cases closed. <b>% ACTIONED</b> = E-MAILS ACTIONED ÷ TOTAL, goal 85%. A mailbox block (a name with dates beside it) counts every email received for that company, against its rep's actioned + closed.</li>
      <li>Every figure is recalculated from the exports, so each cell uses the right day and the right person.</li>
      <li>Weekends are left out unless "Include weekends" is ticked. Ship counts (the WK SHIP CT columns) aren't in these exports and aren't shown.</li>
    </ul>`;
  $('#tFiles').innerHTML = `<table><thead><tr><th>File</th><th>Team</th><th>Report</th><th style="text-align:right">Rows</th><th>Dates</th><th></th></tr></thead><tbody>` +
    M.D.files.map((f, i) => `<tr class="${f.kind ? '' : 'muted'}"><td>${esc(f.name)}</td><td>${esc(f.team || '–')}</td><td style="text-align:left">${esc(f.kind ? (f.kind === 'workbook' ? `${KIND.workbook}: ${f.note}` : KIND[f.kind]) : f.note || 'Skipped')}</td>` +
      `<td style="text-align:right">${f.rows ? fmtN(f.rows) : ''}</td><td style="text-align:left">${f.from != null ? esc(span(f.from, f.to)) : ''}</td>` +
      `<td><button class="cc-x" type="button" data-rm="${i}" title="Take this file out and recalculate">Remove</button></td></tr>`).join('') + '</tbody></table>';
}
// Remove one loaded file and re-read the rest (the last one removed goes back to the start screen).
$('#tFiles').addEventListener('click', (e) => {
  const b = e.target.closest('[data-rm]'); if (!b) return;
  const gone = READ[+b.dataset.rm];
  for (const [k, it] of LOADED) if (it === gone) LOADED.delete(k);
  if (LOADED.size) readFiles([...LOADED.values()]); else { M = null; L = null; renderPending(); show('start'); }
});
$('#filesBtn').addEventListener('click', () => { setTab('names'); requestAnimationFrame(() => $('#tFiles').scrollIntoView({ behavior: 'smooth', block: 'center' })); });

// ---------- render ----------
function render() {
  if (!M) return;
  hideTip();
  const F = filters();
  const secs = [...S.secs], who = S.people.size ? plural(F.rows.size, 'person').replace('persons', 'people') : secs.length === L.sections.length ? (L.fromBook ? 'everyone' : 'all teams') : secs.join(' / ');
  $('#eyebrow').textContent = `Actioned & closed case count · ${who}${F.cust ? ` · ${plural(F.cust.size, 'customer')}` : ''}`;
  $('#range').innerHTML = `<b>${esc(span(S.from, S.to))}</b> · ${plural(F.days.length, S.weekends ? 'day' : 'weekday')}`;
  const t = S.tab;
  if (t === 'over') { const series = daily(M, L, F); renderTiles(series); dailyChart($('#cDaily'), series); renderTable($('#tPerson'), summaryModel(M, L, F)); renderTop(F); }
  else if (t === 'cases') renderCases(F);
  else if (t === 'cust') renderCust(F);
  else if (t === 'names') renderNames();
}
let rz = 0;
addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { if (M && S.tab === 'over') dailyChart($('#cDaily'), daily(M, L, filters())); }, 150); });

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
const xDate = (d) => ({ t: 'n', v: d + 25569, z: 'm/d' });
const dowRow = (m) => ['', ...m.cols.slice(1).map((c) => (c.d != null ? DOW[dowOf(c.d)].toUpperCase() : ''))];
const headRow = (m, first) => [first, ...m.cols.slice(1).map((c) => (c.d != null ? xDate(c.d) : c.x))];
// The export keeps the workbook's layout, so blocks can be pasted straight into it.
async function exportExcel() {
  const btn = $('#exportBtn'), label = btn.textContent;
  btn.disabled = true; btn.textContent = 'Preparing…';
  try {
    const XLSX = await loadXlsx(), wb = XLSX.utils.book_new(), F = filters();
    const add = (name, aoa, widths) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = widths; XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31)); };
    const aoa = [];
    for (const s of L.sections) {
      if (!s.rows.some((r) => F.rows.has(r.key))) continue;
      aoa.push([s.name]);
      for (const i of [0, 1]) { const m = casesModel(M, s, i, F); aoa.push(dowRow(m), headRow(m, i ? 'CASES CLOSED' : 'CASES ACTIONED'), ...m.rows.map((r) => r.c.map(xCell)), []); }
      aoa.push([]);
    }
    const k = monthKey(S.to);
    add(monthKey(S.from) === k ? `${SHEET_MON[k % 12]} ${Math.floor(k / 12)}` : 'CASES', aoa, [{ wch: 18 }]);
    const cm = custModel(M, L, F);
    add(`CUST E-MAILS${monthKey(S.from) === k ? ' ' + SHEET_MON[k % 12] : ''}`, [headRow(cm, 'CUSTOMER / REP'), ...cm.rows.map((r) => r.c.map(xCell))], [{ wch: 30 }]);
    const sm = summaryModel(M, L, F);
    add('BY PERSON', [sm.cols.map((c) => c.x), ...sm.rows.map((r) => r.c.map(xCell))], [{ wch: 24 }, { wch: 15 }, { wch: 13 }, { wch: 17 }, { wch: 16 }, { wch: 12 }]);
    const pad = (n) => String(n).padStart(2, '0'), d = new Date();
    XLSX.writeFile(wb, `ACTIONED & CLOSED CASE COUNT ${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}.xlsx`);
  } catch (err) {
    $('#notice').hidden = false; $('#notice').textContent = err.message;
  } finally { btn.disabled = false; btn.textContent = label; }
}
$('#exportBtn').addEventListener('click', () => { if (M) exportExcel(); });
