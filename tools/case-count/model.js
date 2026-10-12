// The case-count tool's numbers, with no page code, so the same logic can be checked outside the browser.
// Input D comes from worker.js: teams[], people[], companies[], files[], template, and count tables
//   act [team, person, day, n]                   emails sent on the person's cases (CASES ACTIONED)
//   clo [team, person, day, n]                   cases closed (CASES CLOSED)
//   inb [team, person, company, day, ceva, n]    emails received; ceva 1 = from a CEVA address
//   sla [team, person, company, opened, closed, elapsed, target, met, status, segment, parent, case, subject, week]
//                                                one row per case in the SLA export; opened/closed are minutes since
//                                                1970-01-01, elapsed/target minutes, met 1/0, text fields index strs[]
// person -1 = no known case owner; days are whole days since 1970-01-01.
export const GOAL = 0.85;

// ---------- dates ----------
const DAY = 864e5;
export const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export const SHEET_MON = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE', 'JULY', 'AUG', 'SEPT', 'OCT', 'NOV', 'DEC'];   // as the workbook names its tabs
export const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
export const dt = (d) => new Date(d * DAY);
export const dowOf = (d) => dt(d).getUTCDay();
export const monday = (d) => d - ((d + 3) % 7);   // day 0 (1970-01-01) was a Thursday
export const isWeekday = (d) => { const w = dowOf(d); return w > 0 && w < 6; };
export const md = (d) => `${dt(d).getUTCMonth() + 1}/${dt(d).getUTCDate()}`;
export const dLong = (d, year = true) => `${MON[dt(d).getUTCMonth()]} ${dt(d).getUTCDate()}${year ? ', ' + dt(d).getUTCFullYear() : ''}`;
export const span = (a, b) => (a === b ? dLong(a) : dt(a).getUTCFullYear() === dt(b).getUTCFullYear() ? `${dLong(a, false)} – ${dLong(b)}` : `${dLong(a)} – ${dLong(b)}`);
export const monthKey = (d) => dt(d).getUTCFullYear() * 12 + dt(d).getUTCMonth();
export const monthStart = (k) => Date.UTC(Math.floor(k / 12), k % 12, 1) / DAY;
export const monthEnd = (k) => monthStart(k + 1) - 1;
export const sheetName = (k) => `${SHEET_MON[k % 12]} ${Math.floor(k / 12)}`;
export const byText = (a, b) => String(a).localeCompare(String(b), 'en', { sensitivity: 'base', numeric: true });

// ---------- indexes ----------
// I.pday: person -> day -> [actioned, closed, received from CEVA on their cases, received on their cases]
// I.comp: company -> day -> received (any owner); I.compOwn: "person,company" -> day -> received not from CEVA
// X: the days the case-count reports cover; XS: the days SLA cases were opened; XA: both. Only the reports
// that were loaded count, so a single file still gives a dashboard.
export const slaDay = (r) => Math.floor(r[3] / 1440);
export function prepare(D) {
  D.sla = D.sla || []; D.strs = D.strs || [];
  let lo = Infinity, hi = -Infinity;
  for (const r of D.act) { lo = Math.min(lo, r[2]); hi = Math.max(hi, r[2]); }
  for (const r of D.clo) { lo = Math.min(lo, r[2]); hi = Math.max(hi, r[2]); }
  for (const r of D.inb) { lo = Math.min(lo, r[3]); hi = Math.max(hi, r[3]); }
  let slo = Infinity, shi = -Infinity;
  for (const r of D.sla) { const d = slaDay(r); slo = Math.min(slo, d); shi = Math.max(shi, d); }
  const has = { act: D.act.length > 0, clo: D.clo.length > 0, inb: D.inb.length > 0, sla: D.sla.length > 0 };
  has.cc = has.act || has.clo || has.inb;
  if (!has.cc) { lo = slo; hi = shi; }
  if (!has.sla) { slo = lo; shi = hi; }
  const pday = new Map(), comp = new Map(), compOwn = new Map();
  const slot = (p, d) => { let m = pday.get(p); if (!m) pday.set(p, (m = new Map())); let v = m.get(d); if (!v) m.set(d, (v = [0, 0, 0, 0])); return v; };
  for (const r of D.act) if (r[1] >= 0) slot(r[1], r[2])[0] += r[3];
  for (const r of D.clo) if (r[1] >= 0) slot(r[1], r[2])[1] += r[3];
  for (const r of D.inb) {
    if (r[1] >= 0) { const v = slot(r[1], r[3]); v[3] += r[5]; if (r[4]) v[2] += r[5]; }
    let m = comp.get(r[2]); if (!m) comp.set(r[2], (m = new Map())); m.set(r[3], (m.get(r[3]) || 0) + r[5]);
    if (r[1] >= 0 && !r[4]) { const k = r[1] + ',' + r[2]; let o = compOwn.get(k); if (!o) compOwn.set(k, (o = new Map())); o.set(r[3], (o.get(r[3]) || 0) + r[5]); }
  }
  // each person's team folder: where most of their activity is
  const pteam = new Map(), votes = new Map();
  const vote = (t, p, n) => { if (p < 0) return; let m = votes.get(p); if (!m) votes.set(p, (m = new Map())); m.set(t, (m.get(t) || 0) + n); };
  for (const r of D.act) vote(r[0], r[1], r[3]);
  for (const r of D.clo) vote(r[0], r[1], r[3]);
  for (const r of D.inb) vote(r[0], r[1], r[5]);
  for (const r of D.sla) vote(r[0], r[1], 1);
  for (const [p, m] of votes) pteam.set(p, [...m].sort((a, b) => b[1] - a[1])[0][0]);
  const slaBy = new Map();   // person -> SLA rows
  for (const r of D.sla) { let a = slaBy.get(r[1]); if (!a) slaBy.set(r[1], (a = [])); a.push(r); }
  return { D, has, X: { lo, hi }, XS: { lo: slo, hi: shi }, XA: { lo: Math.min(lo, slo), hi: Math.max(hi, shi) }, I: { pday, comp, compOwn, pteam, slaBy } };
}
export const pv = (M, p, d, i) => { const m = M.I.pday.get(p), v = m && m.get(d); return v ? v[i] : 0; };

// ---------- matching workbook names to Salesforce ----------
// Common nicknames, so a short form finds the full name.
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
// every word of the workbook name matches a word of the Salesforce name (a first or middle name)
const nameHit = (label, name) => { const l = toks(label), n = toks(name); return l.length > 0 && l.every((w) => n.some((t) => wordEq(w, t))); };
const labelEq = (a, b) => { const x = toks(a).join(' '), y = toks(b).join(' '); return x === y || (x.length >= 3 && y.startsWith(x)) || (y.length >= 3 && x.startsWith(y)) || nickOf(x).has(y); };
// How well a person's daily counts fit the workbook's own numbers for that name, on the weekdays both
// cover (0 = identical). Used when the name alone doesn't decide (a nickname not in Salesforce).
function fit(M, key, p) {
  const h = M.D.template && M.D.template.hist[key]; if (!h) return null;
  const pd = M.I.pday.get(p); let err = 0, base = 0, n = 0;
  [['act', 0], ['clo', 1]].forEach(([k, i]) => {
    for (const [ds, v] of Object.entries(h[k])) {
      const d = +ds; if (d < M.X.lo || d >= M.X.hi || !isWeekday(d)) continue;
      const a = pd && pd.get(d) ? pd.get(d)[i] : 0; err += Math.abs(v - a); base += v; n++;
    }
  });
  return n >= 4 && base > 0 ? err / base : null;
}
// The "No case owner (team)" row the worker makes for emails on cases with no known owner: it has customer
// emails but nobody actioning them, so it stays out of every % actioned.
export const isUnowned = (M, p) => p >= 0 && /^No case owner \(/.test(M.D.people[p]);
const activity = (M, p) => { let n = 0; const m = M.I.pday.get(p); if (m) for (const v of m.values()) n += v[0] + v[1]; return n; };
const personByName = (M, name) => M.D.people.findIndex((n) => n.toLowerCase() === String(name).trim().toLowerCase());

// ---------- layout ----------
// L.sections [{ name, rows: [{ key, label, p, how }] }]   one section per supervisor on the month tab
// L.cust [{ id, label, p, rowKey, mailbox, ceva, total, ownerOnly, customers: [{ id, label, comps: [] }] }]
// ovr: fixes made on the Names tab, row key -> person
export function buildLayout(M, ovr = new Map()) {
  const D = M.D, T = D.template;
  const L = { fromBook: !!(T && T.sections.length), month: T && T.month, sections: [], cust: [] };
  if (L.fromBook) {
    const all = D.people.map((_, i) => i);
    // 1. fixes from the Names tab or a ROSTER tab, 2. the name, 3. the numbers. Every section is matched by
    // name before any numbers are used, so loading one team's files or both puts people on the same rows.
    const cands = new Map();   // row -> people whose name fits it
    for (const s of T.sections) {
      const rows = s.labels.map((label) => ({ key: s.name + '|' + label, label, sec: s.name, p: -1, how: 'Not in the loaded files' }));
      for (const r of rows) {
        if (ovr.has(r.key)) { r.p = ovr.get(r.key); r.how = 'Picked on this page'; continue; }
        const ro = (T.roster || []).find((x) => labelEq(x.label, r.label) && (!x.section || labelEq(x.section, s.name)));
        const p = ro ? personByName(M, ro.name) : -1;
        if (p >= 0) { r.p = p; r.how = 'ROSTER tab'; continue; }
        const cs = all.filter((i) => nameHit(r.label, D.people[i]));
        if (cs.length) cands.set(r, cs);
      }
      L.sections.push({ name: s.name, rows });
    }
    // Which team folder each section mostly holds, from the rows only one person's name fits (a supervisor
    // section is mostly one team), so a shared first name goes to the section of that person's own team.
    const secTeam = new Map();
    for (const s of L.sections) {
      const m = new Map();
      for (const r of s.rows) { const p = r.p >= 0 ? r.p : cands.has(r) && cands.get(r).length === 1 ? cands.get(r)[0] : -1; const t = M.I.pteam.get(p); if (t != null) m.set(t, (m.get(t) || 0) + 1); }
      secTeam.set(s.name, m);
    }
    const teamFit = (r, p) => secTeam.get(r.sec).get(M.I.pteam.get(p)) || 0;
    const fitOr = (r, p) => { const f = fit(M, r.key, p); return f == null ? Infinity : f; };
    const better = (r, a, b) => teamFit(r, b) - teamFit(r, a) || fitOr(r, a) - fitOr(r, b) || activity(M, b) - activity(M, a);   // < 0: a fits r better
    for (const [r, cs] of cands) {
      r.p = [...cs].sort((a, b) => better(r, a, b))[0];
      r.how = cs.length === 1 ? 'Name' : secTeam.get(r.sec).size ? 'Name and team' : 'Name (several match: check)';
    }
    // One person fills one row. If a name matched the same person in two sections (two people with the
    // same first name, only one of them in the loaded files), the row in that person's team's section keeps them.
    const claims = new Map();
    for (const s of L.sections) for (const r of s.rows) if (cands.has(r) && r.p >= 0) { if (!claims.has(r.p)) claims.set(r.p, []); claims.get(r.p).push(r); }
    for (const [p, rs] of claims) {
      if (rs.length < 2) continue;
      const keep = rs.reduce((a, b) => (teamFit(b, p) - teamFit(a, p) || fitOr(a, p) - fitOr(b, p) || 0) > 0 ? b : a);
      for (const r of rs) if (r !== keep) { r.p = -1; r.how = 'Not in the loaded files'; }
    }
    // Rows no name fits (a nickname that isn't in Salesforce) take the person whose numbers fit best, but
    // never someone already on a row or whose own name fits a row (they belong there, loaded or not).
    const used = new Set(L.sections.flatMap((s) => s.rows.map((r) => r.p)));
    const labels = L.sections.flatMap((s) => s.rows.map((r) => r.label));
    const free = all.filter((i) => !used.has(i) && !labels.some((l) => nameHit(l, D.people[i])));
    for (const s of L.sections) for (const r of s.rows) {
      if (r.p >= 0 || ovr.has(r.key) || cands.has(r)) continue;
      const best = free.filter((i) => !used.has(i)).map((i) => [i, fit(M, r.key, i)]).filter((x) => x[1] != null && x[1] <= 0.35).sort((a, b) => a[1] - b[1])[0];
      if (best) { r.p = best[0]; r.how = 'Numbers match'; used.add(best[0]); }
    }
    const findRow = (sec, label) => { const s = L.sections.find((x) => x.name === sec); const pool = s ? [s] : L.sections; for (const x of pool) { const r = x.rows.find((y) => labelEq(y.label, label)); if (r) return r; } return null; };
    // CUST E-MAILS blocks: the rep comes from the E-MAILS ACTIONED formulas, else from the name
    T.cust.forEach((b, bi) => {
      const r = (b.ref && findRow(b.ref.section, b.ref.label)) || findRow(null, b.label), p = r ? r.p : -1, allowed = teamCompanies(M, p);
      // A rep who isn't in the loaded files (the other team's exports) gets no customer matches, rather than
      // look-alike companies from the team that is loaded.
      const none = p < 0 && !b.mailbox && r;
      L.cust.push({ id: bi, label: b.label, p, rowKey: r ? r.key : null, mailbox: b.mailbox, ceva: b.ceva, total: b.total || !b.mailbox,
        customers: b.customers.map((c, ci) => ({ id: bi + ':' + ci, label: c, comps: none ? [] : matchCompanies(M, c, allowed) })) });
    });
    // A team with nobody on the workbook (a team folder the workbook doesn't track) gets a section of its
    // own with full names, so its numbers still show. Others not on the workbook (queues, supervisors)
    // stay on the Names tab's list.
    const onRow = new Set(L.sections.flatMap((s) => s.rows.map((r) => r.p)).filter((p) => p >= 0));
    const covered = new Set([...onRow].map((p) => M.I.pteam.get(p)));
    teamSections(M, L, (p) => !onRow.has(p) && !covered.has(M.I.pteam.get(p)), 'Team not on the workbook');
  } else {
    // Without the workbook: one section per team folder, full names A to Z.
    teamSections(M, L, () => true, 'Salesforce name');
  }
  return L;
}
// One section per team folder for the people `want` picks, full names A to Z, and per person (with emails
// received) a CUST E-MAILS block of the customers on the cases they own.
function teamSections(M, L, want, how) {
  const D = M.D, by = new Map();
  for (const [tbl, pi] of [[D.act, 1], [D.clo, 1], [D.inb, 1], [D.sla, 1]]) for (const r of tbl) {
    const p = r[pi]; if (p < 0 || !want(p)) continue;
    const t = M.I.pteam.get(p); if (!by.has(t)) by.set(t, new Set()); by.get(t).add(p);
  }
  const taken = new Set(L.sections.map((s) => s.name));
  for (const t of [...by.keys()].sort((a, b) => byText(D.teams[a], D.teams[b]))) {
    const ps = [...by.get(t)].sort((a, b) => byText(D.people[a], D.people[b]));
    let name = D.teams[t].toUpperCase(); if (taken.has(name)) name += ' (TEAM)';
    taken.add(name);
    L.sections.push({ name, rows: ps.map((p) => ({ key: name + '|' + D.people[p], label: D.people[p], p, how })) });
    for (const p of ps) {
      if (!D.inb.some((r) => r[1] === p)) continue;
      const comps = [...new Set(D.inb.filter((r) => r[1] === p && !r[4]).map((r) => r[2]))].sort((a, b) => byText(D.companies[a], D.companies[b]));
      const id = L.cust.length;
      L.cust.push({ id, label: D.people[p], p, rowKey: name + '|' + D.people[p], mailbox: false, ceva: true, total: true, ownerOnly: true, unowned: isUnowned(M, p),
        customers: comps.map((c, ci) => ({ id: id + ':' + ci, label: D.companies[c], comps: [c] })) });
    }
  }
}
// People in the exports who aren't on the layout (queues, supervisors, other teams), most active first.
export function unlisted(M, L) {
  const on = new Set(L.sections.flatMap((s) => s.rows.map((r) => r.p)));
  return M.D.people.map((n, p) => ({ p, name: n, team: M.D.teams[M.I.pteam.get(p)] || '', n: activity(M, p), sla: (M.I.slaBy.get(p) || []).length }))
    .filter((x) => !on.has(x.p) && x.n + x.sla > 0).sort((a, b) => b.n - a.n || b.sla - a.sla);
}

// Customer rows on CUST E-MAILS use the workbook's own customer names; they are matched to Salesforce
// company names by their words, rare words counting more. A company is counted when most of the row's
// words and most of the company's own words agree (a shared word such as "CLUB" alone is not enough).
// If nothing agrees that well, companies containing every known word of the row, starting with its first
// word, are used (a one-word row matching a longer company name). Names & sources lists every match.
const STOP = new Set(['INC', 'LLC', 'LTD', 'CORP', 'CORPORATION', 'CO', 'COMPANY', 'THE', 'OF', 'AND', 'DBA', 'PARENT', 'HOLDINGS', 'GROUP', 'US', 'USA', 'MAILBOX']);
const ctoks = (s) => toks(s).filter((w) => w.length >= 2 && !STOP.has(w));
// same word, an abbreviation of it ("HSN" / "HSNI"), a plural ("HOMES" / "HOME") or a one-letter typo
const typo = (a, b) => a.length >= 5 && a.length === b.length && [...a].filter((x, i) => x !== b[i]).length === 1;
const tokEq = (l, c) => l === c || (l.length >= 3 && c.startsWith(l) && c.length - l.length <= 2) || (c.length >= 4 && l.startsWith(c) && l.length - c.length <= 1) || typo(l, c);
// Only companies with emails in the rep's own team folder are considered.
function teamCompanies(M, p) {
  if (p < 0) return null;
  const t = M.I.pteam.get(p);
  return t == null ? null : new Set(M.D.inb.filter((r) => r[0] === t).map((r) => r[2]));
}
const CT = new WeakMap();
function matchCompanies(M, label, allowed) {
  const D = M.D;
  let ct = CT.get(D);
  if (!ct) {
    const t = D.companies.map(ctoks), df = new Map();
    for (const ts of t) for (const w of new Set(ts)) df.set(w, (df.get(w) || 0) + 1);
    ct = { t, joined: D.companies.map((c) => toks(c).join('')), idfC: (w) => Math.log(1 + t.length / (df.get(w) || 1)) };
    CT.set(D, ct);
  }
  const N = ct.t.length, words = ctoks(label), joined = toks(label).filter((w) => !STOP.has(w)).join('');
  const seen = words.map((w) => [w, ct.t.reduce((n, ts) => n + (ts.some((c) => tokEq(w, c)) ? 1 : 0), 0)]).filter((x) => x[1] > 0).map(([w, n]) => [w, Math.log(1 + N / n)]);
  const strong = [], loose = [];
  ct.t.forEach((ts, c) => {
    if (allowed && !allowed.has(c)) return;
    if (joined.length >= 5 && ct.joined[c].includes(joined)) { strong.push(c); return; }   // two words written as one
    if (!seen.length || !ts.length) return;
    let lm = 0, lt = 0, all = true;
    for (const [w, idf] of seen) { lt += idf; if (ts.some((x) => tokEq(w, x))) lm += idf; else all = false; }
    let cm = 0, ctot = 0;
    for (const x of ts) { const i = ct.idfC(x); ctot += i; if (seen.some(([w]) => tokEq(w, x))) cm += i; }
    // half of the company's words is enough only when every known word of the row is in it
    // ("ACME WIDGETS" = Acme Trading Co, but "ZETA SPA" is not Delta Spas)
    const cf = cm / ctot;
    if (lm / lt >= 0.5 && (cf > 0.5 || (cf >= 0.5 && all)) && seen.some(([w]) => w === words[0])) strong.push(c);
    else if (all && words.length && seen[0][0] === words[0]) loose.push(c);
  });
  const out = strong.length ? strong : loose;
  // Also count companies that start with the row's own first word when few companies share it: the
  // same customer under a second Salesforce name ("ACME FREIGHT" also counts "Acme.com (Americas)").
  const first = words[0], shared = first ? ct.t.filter((ts) => ts.some((x) => tokEq(first, x))).length : 0;
  if (out.length && shared > 0 && shared <= 3) {
    ct.t.forEach((ts, c) => { if ((!allowed || allowed.has(c)) && ts.length && tokEq(first, ts[0]) && !out.includes(c)) out.push(c); });
  }
  return out;
}

// ---------- filters ----------
// F: { rows: Set of layout row keys shown, days: [weekdays (or every day) in the date range], cust: Set of customer ids or null for all }
export function weeksOf(days) {
  const out = [];
  for (const d of days) { const w = out[out.length - 1]; if (w && monday(w[0]) === monday(d)) w.push(d); else out.push([d]); }
  return out;
}
export const visibleBlocks = (L, F) => L.cust.filter((b) => b.rowKey && F.rows.has(b.rowKey) && (!F.cust || b.customers.some((c) => F.cust.has(c.id))));

// ---------- tables: one model renders to HTML, copies as text and exports to Excel ----------
// model: { cls, cols: [{ h (html), x (plain), cls, d (day), wk }], rows: [{ cls, c: [cell] }] }; cell: { v, f, cls, title }
// f: 's' text, 'n' count (blank when 0), 'z' count (0 shown), 'p' percent against the 85% goal,
//    'q' percent with one decimal (no goal), 't' minutes shown as "Hh MMm"
export const cS = (v, cls, title) => ({ v, f: 's', cls, title }), cN = (v, cls) => ({ v, f: 'n', cls }), cZ = (v, cls) => ({ v, f: 'z', cls }), cP = (v, cls) => ({ v, f: 'p', cls });
export const cQ = (v, cls) => ({ v, f: 'q', cls }), cT = (v, cls) => ({ v, f: 't', cls });
// Like the SLA Dashboard (and its Power BI measure): ROUND(AVERAGE(minutes), 0) shown as "Hh MMm"
export const fmtHM = (mins) => (mins == null || !isFinite(mins) ? '' : `${Math.floor(Math.round(mins) / 60)}h ${String(Math.round(mins) % 60).padStart(2, '0')}m`);
export const goalCls = (p) => (p == null ? '' : p < GOAL ? 'low' : 'ok');
export const col = (h, cls, x) => ({ h, cls, x: x ?? h.replace(/<small>/g, ' ').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&') });
export const dayCol = (d) => ({ ...col(`${DOW[dowOf(d)].toUpperCase()}<small>${md(d)}</small>`, isWeekday(d) ? '' : 'muted', md(d)), d });
const sumBy = (days, f) => { let s = 0; for (const d of days) s += f(d) || 0; return s; };

// CASES ACTIONED (i = 0) or CASES CLOSED (i = 1) for one section, like the month tab: the days in the
// range, a TOTAL after each week, and a TOTAL for the whole range when it spans more than one week.
export function casesModel(M, sec, i, F) {
  const rows0 = sec.rows.filter((r) => F.rows.has(r.key)), weeks = weeksOf(F.days), many = weeks.length > 1;
  const cols = [col(i ? 'CASES CLOSED' : 'CASES ACTIONED')];
  for (const w of weeks) { for (const d of w) cols.push(dayCol(d)); cols.push({ ...col(`TOTAL<small>${md(w[0])}–${md(w[w.length - 1])}</small>`, 'wk', 'TOTAL'), wk: w }); }
  if (many) cols.push({ ...col('TOTAL<small>all days</small>', 'wk all', 'TOTAL ALL DAYS'), wk: F.days });
  const v = (p, ds) => (p < 0 ? null : sumBy(ds, (d) => pv(M, p, d, i)));
  const cell = (p, c) => { const x = v(p, c.d != null ? [c.d] : c.wk); return x == null ? cS('') : cZ(x); };
  const rows = rows0.map((r) => ({ c: [cS(r.label, r.p < 0 ? 'muted' : '', r.p >= 0 ? M.D.people[r.p] : 'Not in the loaded exports'), ...cols.slice(1).map((c) => cell(r.p, c))] }));
  rows.push({ cls: 'total', c: [cS('TOTALS'), ...cols.slice(1).map((c) => cZ(rows0.reduce((a, r) => a + (v(r.p, c.d != null ? [c.d] : c.wk) || 0), 0)))] });
  return { cls: 'mx mirror', cols, rows, n: rows0.length };
}

// CUST E-MAILS lines for one block, with a value function per day:
//   customer rows: every email received for the companies matched to the row (any case owner), or, without
//   the workbook, the emails on the rep's own cases; E-MAILS FROM CEVA STATIONS: emails from CEVA addresses
//   on the rep's cases; TOTAL: the sum of those; E-MAILS ACTIONED: the rep's actioned + closed;
//   % ACTIONED: E-MAILS ACTIONED / TOTAL (goal 85%).
export function blockLines(M, b) {
  const compDay = (cs, d) => cs.reduce((a, c) => { const m = M.I.comp.get(c); return a + ((m && m.get(d)) || 0); }, 0);
  const ownDay = (cs, d) => cs.reduce((a, c) => { const m = M.I.compOwn.get(b.p + ',' + c); return a + ((m && m.get(d)) || 0); }, 0);
  const custs = b.customers.map((c) => ({ ...c, f: (d) => (b.ownerOnly ? ownDay(c.comps, d) : compDay(c.comps, d)) }));
  const ceva = (d) => (b.ceva && b.p >= 0 ? pv(M, b.p, d, 2) : 0);
  const total = (d) => custs.reduce((a, c) => a + c.f(d), 0) + ceva(d);
  const act = (d) => (b.p >= 0 ? pv(M, b.p, d, 0) + pv(M, b.p, d, 1) : 0);
  return { custs, ceva, total, act };
}
export function custModel(M, L, F) {
  const days = F.days, cols = [col('CUSTOMER / REP'), ...days.map(dayCol), col('TOTALS', 'wk')], rows = [];
  const blank = () => cols.slice(1).map(() => cS(''));
  const line = (label, f, cls, lcls, title) => ({ cls, c: [cS(label, lcls, title), ...days.map((d) => cZ(f(d))), cZ(sumBy(days, f))] });
  for (const b of visibleBlocks(L, F)) {
    const { custs, ceva, total, act } = blockLines(M, b), rep = b.p >= 0 ? M.D.people[b.p] : 'Rep not in the loaded exports';
    rows.push({ cls: 'grp', c: [cS(b.label, '', rep), ...blank()] });
    for (const c of custs) {
      if (F.cust && !F.cust.has(c.id)) continue;
      rows.push(line(b.mailbox ? 'E-mails received' : c.label, c.f, '', c.comps.length ? 'indent' : 'indent muted', c.comps.length ? c.comps.map((x) => M.D.companies[x]).join(', ') : 'No Salesforce company matched'));
    }
    if (b.ceva) rows.push(line('E-MAILS FROM CEVA STATIONS', ceva, '', 'indent'));
    if (b.total) rows.push(line('TOTAL', total, 'sub'));
    if (b.unowned) continue;   // nobody to have actioned them
    rows.push(line('E-MAILS ACTIONED', act, 'key', '', `${rep}: emails actioned + cases closed`));
    const pct = (a, t) => (t ? a / t : null), all = pct(sumBy(days, act), sumBy(days, total));
    rows.push({ cls: 'key pct', c: [cS('% ACTIONED (GOAL 85%)'), ...days.map((d) => { const x = pct(act(d), total(d)); return cP(x, goalCls(x)); }), cP(all, goalCls(all))] });
  }
  return { cls: 'mx mirror', cols, rows, empty: 'No customer blocks for these filters.' };
}

// Per person: actioned, closed, and (for reps with CUST E-MAILS blocks) the block's customer emails and % actioned.
export function personTotals(M, L, F) {
  const blocksOf = new Map();
  for (const b of L.cust) if (b.rowKey) { if (!blocksOf.has(b.rowKey)) blocksOf.set(b.rowKey, []); blocksOf.get(b.rowKey).push(b); }
  const out = [];
  for (const s of L.sections) for (const r of s.rows) {
    if (!F.rows.has(r.key)) continue;
    const t = { sec: s.name, key: r.key, label: r.label, p: r.p, act: 0, clo: 0, recv: null, worked: null, sla: 0, met: 0, unowned: isUnowned(M, r.p) };
    if (r.p >= 0) for (const d of F.days) { t.act += pv(M, r.p, d, 0); t.clo += pv(M, r.p, d, 1); }
    if (r.p >= 0) for (const x of slaOf(M, r.p, F)) { t.sla++; t.met += x[7]; }
    const bs = blocksOf.get(r.key);
    if (bs && r.p >= 0) {
      t.recv = 0;
      for (const b of bs) { const { total } = blockLines(M, b); t.recv += sumBy(F.days, total); }
      t.worked = (t.act + t.clo) * bs.length;   // each block's E-MAILS ACTIONED counts the rep's work once
    }
    out.push(t);
  }
  return out;
}
export function summaryModel(M, L, F) {
  const ts = personTotals(M, L, F), cc = M.has.cc, sl = M.has.sla;
  const cols = [col('Person'), ...(cc ? [col('Cases actioned'), col('Cases closed'), col('Actioned + closed', 'wk'), col('Customer emails'), col('% actioned<small>goal 85%</small>', '', '% actioned')] : []),
    ...(sl ? [col('SLA cases'), col('SLA met %')] : [])];
  const line = (label, xs, lcls, title) => {
    const a = xs.reduce((s, x) => s + x.act, 0), c = xs.reduce((s, x) => s + x.clo, 0);
    const withB = xs.filter((x) => x.recv != null), recv = withB.reduce((s, x) => s + x.recv, 0);
    const own = withB.filter((x) => !x.unowned), orecv = own.reduce((s, x) => s + x.recv, 0), worked = own.reduce((s, x) => s + x.worked, 0);
    const p = own.length && orecv ? worked / orecv : null;
    const n = xs.reduce((s, x) => s + x.sla, 0), met = xs.reduce((s, x) => s + x.met, 0);
    return [cS(label, lcls, title), ...(cc ? [cZ(a), cZ(c), cZ(a + c), withB.length ? cZ(recv) : cS(''), cP(p, goalCls(p))] : []), ...(sl ? [cZ(n), cQ(n ? met / n : null)] : [])];
  };
  const rows = [], secs = [...new Set(ts.map((x) => x.sec))];
  for (const s of secs) {
    const xs = ts.filter((x) => x.sec === s);
    if (secs.length > 1 || L.sections.length > 1) rows.push({ cls: 'grp', c: [cS(s), ...cols.slice(1).map(() => cS(''))] });
    for (const x of xs) rows.push({ c: line(x.label, [x], x.p < 0 ? 'muted indent' : 'indent', x.p >= 0 ? M.D.people[x.p] : 'Not in the loaded exports') });
    if (secs.length > 1) rows.push({ cls: 'sub', c: line(`${s} total`, xs) });
  }
  if (ts.length > 1) rows.push({ cls: 'total', c: line('Total', ts) });
  return { cols, rows, empty: 'Nobody matches these filters.' };
}
// Totals per day for the chart and the headline figures.
export function daily(M, L, F) {
  const ps = [...new Set(L.sections.flatMap((s) => s.rows.filter((r) => F.rows.has(r.key) && r.p >= 0).map((r) => r.p)))];
  const blocks = L.cust.filter((b) => b.rowKey && F.rows.has(b.rowKey)).map((b) => ({ b, ...blockLines(M, b) }));
  // recv: every customer email shown; precv: those with someone to action them (the base of % actioned)
  return F.days.map((d) => {
    let act = 0, clo = 0, recv = 0, precv = 0, worked = 0;
    for (const p of ps) { act += pv(M, p, d, 0); clo += pv(M, p, d, 1); }
    for (const x of blocks) if (x.b.p >= 0) { const t = x.total(d); recv += t; if (!x.b.unowned) { precv += t; worked += x.act(d); } }
    return { d, act, clo, recv, precv, worked };
  });
}
// Customer rows by emails received in the range (for the overview bars).
export function topCustomers(M, L, F, n = 10) {
  const out = [];
  for (const b of visibleBlocks(L, F)) {
    const { custs } = blockLines(M, b);
    for (const c of custs) if ((!F.cust || F.cust.has(c.id)) && c.comps.length) out.push({ label: b.mailbox ? b.label : c.label, rep: b.label, n: sumBy(F.days, c.f) });
  }
  return out.filter((x) => x.n > 0).sort((a, b) => b.n - a.n).slice(0, n);
}

// ---------- SLA (response time) ----------
// SLA cases count by the day they were opened, every day in the range (weekends too), for the people on the
// rows the filters show. F.from / F.to bound the range.
export function slaOf(M, p, F) {
  const a = M.I.slaBy.get(p); if (!a) return [];
  return a.filter((r) => { const d = slaDay(r); return d >= F.from && d <= F.to; });
}
export function slaRows(M, L, F) {
  const ps = new Set(L.sections.flatMap((s) => s.rows.filter((r) => F.rows.has(r.key) && r.p >= 0).map((r) => r.p)));
  return [...ps].flatMap((p) => slaOf(M, p, F));
}
// SLA cases in the date range whose owner isn't on any row (queues, supervisors), so they're not counted.
export function slaOffRows(M, L, F) {
  const on = new Set(L.sections.flatMap((s) => s.rows.map((r) => r.p)).filter((p) => p >= 0));
  return M.D.sla.filter((r) => !on.has(r[1]) && slaDay(r) >= F.from && slaDay(r) <= F.to).length;
}
export function slaStats(rows) {
  const s = { n: 0, met: 0, sum: 0, cnt: 0, m: { n: 0, sum: 0, cnt: 0 }, b: { n: 0, sum: 0, cnt: 0 } };
  for (const r of rows) {
    const g = r[7] ? s.m : s.b; s.n++; g.n++; if (r[7]) s.met++;
    if (r[5] != null) { s.sum += r[5]; s.cnt++; g.sum += r[5]; g.cnt++; }
  }
  return s;
}
export const slaAvg = (o) => (o && o.cnt ? o.sum / o.cnt : null);
const isoWeekOf = (min) => { const x = new Date(Math.floor(min / 1440) * 864e5), day = x.getUTCDay() || 7; x.setUTCDate(x.getUTCDate() + 4 - day); const y0 = Date.UTC(x.getUTCFullYear(), 0, 1); return [x.getUTCFullYear(), Math.ceil(((x - y0) / 864e5 + 1) / 7)]; };
// Ways to group the SLA table: key and label per case, and how the groups are ordered.
export function slaGroups(M, L) {
  const D = M.D, label = new Map();
  for (const s of L.sections) for (const r of s.rows) if (r.p >= 0 && !label.has(r.p)) label.set(r.p, r.label);
  const txt = (i) => D.strs[i] || '(Blank)';
  return {
    owner: { name: 'Case owner', key: (r) => r[1], label: (k) => (k >= 0 ? label.get(k) || D.people[k] : '(No owner)'), title: (k) => (k >= 0 ? D.people[k] : ''), order: 'n' },
    week: { name: 'Week', key: (r) => { const [y, w] = isoWeekOf(r[3]); return y * 100 + (r[13] || w); }, label: (k, rs) => { const ds = rs.map(slaDay); return `Week ${k % 100} (${md(Math.min(...ds))} – ${md(Math.max(...ds))})`; }, order: 'k' },
    day: { name: 'Day opened', key: slaDay, label: (k) => `${DOW[dowOf(k)]} ${md(k)}`, order: 'k' },
    company: { name: 'Company', key: (r) => r[2], label: txt, order: 'n' },
    parent: { name: 'Ultimate parent', key: (r) => r[10], label: txt, order: 'n' },
    segment: { name: 'Customer segment', key: (r) => r[9], label: txt, order: 't' },
    status: { name: 'Status', key: (r) => r[8], label: txt, order: 't' },
  };
}
// The SLA matrix: per group, SLA met and breached (average elapsed time and cases), then the total and SLA met %.
// With a second grouping (by2), each row can be opened (open: Set of row keys, or 'all') to show its rows by that.
// Row keys: "k" for a top row, "k\u0001k2" for a row under it.
const SEP = '\u0001';
function slaGroupBy(G, rows) {
  const m = new Map();
  for (const r of rows) { const k = G.key(r); let a = m.get(k); if (!a) m.set(k, (a = [])); a.push(r); }
  const groups = [...m].map(([k, rs]) => ({ k, rs, s: slaStats(rs), l: G.label(k, rs) }));
  return groups.sort(G.order === 'k' ? (a, b) => a.k - b.k : G.order === 't' ? (a, b) => byText(a.l, b.l) : (a, b) => b.s.n - a.s.n || byText(a.l, b.l));
}
export function slaModel(M, L, rows, by, sel, by2, open) {
  const GS = slaGroups(M, L), G = GS[by], G2 = by2 && by2 !== by ? GS[by2] : null;
  const groups = slaGroupBy(G, rows);
  const cols = [col(G2 ? `${G.name} › ${G2.name}` : G.name, '', G2 ? `${G.name} / ${G2.name}` : G.name), col('Met<small>avg elapsed</small>', 'met', 'Met avg elapsed'), col('Met<small>cases</small>', 'met', 'Met cases'),
    col('Breached<small>avg elapsed</small>', 'br', 'Breached avg elapsed'), col('Breached<small>cases</small>', 'br', 'Breached cases'),
    col('Total<small>avg elapsed</small>', 'wk', 'Total avg elapsed'), col('Total<small>cases</small>', 'wk', 'Total cases'), col('SLA met %', 'wk')];
  const cells = (s) => [cT(slaAvg(s.m)), cN(s.m.n), cT(slaAvg(s.b)), cN(s.b.n, s.b.n ? 'low' : ''), cT(slaAvg(s)), cZ(s.n), cQ(s.n ? s.met / s.n : null)];
  const out = [];
  for (const g of groups) {
    const key = String(g.k), isOpen = !!G2 && (open === 'all' || (open && open.has(key)));
    out.push({ cls: [sel === key ? 'selr' : '', G2 ? 'lv0' : ''].filter(Boolean).join(' '), sel: key, tog: G2 ? (isOpen ? 'open' : 'shut') : null,
      c: [cS(g.l, '', G.title ? G.title(g.k) : g.l), ...cells(g.s)] });
    if (isOpen) for (const h of slaGroupBy(G2, g.rs)) {
      const k2 = key + SEP + h.k;
      out.push({ cls: ['lv1', sel === k2 ? 'selr' : ''].filter(Boolean).join(' '), sel: k2, c: [cS(h.l, 'indent', G2.title ? G2.title(h.k) : h.l), ...cells(h.s)] });
    }
  }
  if (groups.length > 1) out.push({ cls: 'total', c: [cS('Total'), ...cells(slaStats(rows))] });
  return { cls: 'mx sla', cols, rows: out, empty: 'No SLA cases for these filters.' };
}
export function slaPick(M, L, rows, by, sel, by2) {
  if (sel == null) return rows;
  const GS = slaGroups(M, L), [k1, k2] = sel.split(SEP), G2 = by2 && GS[by2];
  return rows.filter((r) => String(GS[by].key(r)) === k1 && (k2 == null || !G2 || String(G2.key(r)) === k2));
}
const mdt = (min) => { if (min == null) return ''; const d = new Date(min * 6e4); let h = d.getUTCHours(); const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12; return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()} ${h}:${String(d.getUTCMinutes()).padStart(2, '0')} ${ap}`; };
// The case list under the matrix, newest first.
export function slaCasesModel(M, rows, limit) {
  const D = M.D, txt = (i) => D.strs[i] || '';
  const list = rows.slice().sort((a, b) => b[3] - a[3]);
  const cols = [col('Case #'), col('Opened'), col('Closed'), col('Case owner'), col('Company'), col('Subject'), col('Status'), col('Segment'), col('Target'), col('Elapsed'), col('SLA')];
  const out = list.slice(0, limit).map((r) => ({ c: [cS(r[11]), cS(mdt(r[3])), cS(mdt(r[4])), cS(r[1] >= 0 ? D.people[r[1]] : ''), cS(txt(r[2])), cS(txt(r[12]), 'subj', txt(r[12])),
    cS(txt(r[8])), cS(txt(r[9])), cT(r[6]), cT(r[5]), cS(r[7] ? 'Met' : 'Breached', r[7] ? 'ok' : 'low')] }));
  return { cls: 'sla-cases', cols, rows: out, total: list.length, empty: 'No SLA cases for these filters.' };
}
