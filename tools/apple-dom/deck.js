// Builds this week's PowerPoint from last week's deck: same CEVA masters, layouts and untouched slides (Projects,
// KPI Metrics, closing). Updates the title slide, the slide 4/5/6 tables, the slide 8 comments and the week labels,
// and puts the Charts Template's own charts on slides 7-9 as editable PowerPoint charts (each carries the filled
// workbook, so "Edit Data" opens it). Pure module: runs in the worker and under Node.
import { unzip, zip, quarterOf, slide8Text } from './model.js';

const dec = new TextDecoder(), enc = new TextEncoder();
const xesc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unx = (t) => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const attr = (a, k) => new RegExp(`\\s${k}="([^"]*)"`).exec(a)?.[1];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const ordinal = (d) => (d % 100 >= 11 && d % 100 <= 13 ? 'th' : ['th', 'st', 'nd', 'rd'][d % 10] || 'th');
const EMU = 12700;   // per point
const GOAL = { ok: '00B050', warn: 'FFFF00', bad: 'FF0000' };   // the decks' cell colours: at goal, >2% below, >5% below
const goalFill = (v) => (v == null ? null : v >= 0.98 ? GOAL.ok : v >= 0.93 ? GOAL.warn : GOAL.bad);
const fmtInt = (n) => Math.round(n).toLocaleString('en-US');
const pct = (v, d) => (v * 100).toFixed(d) + '%';
const unadj = (o) => (o.due ? 1 - o.late / o.due : null), adj = (o) => (o.due ? 1 - o.carrier / o.due : null);

function rels(files, path) {
  const rp = path.replace(/[^/]+$/, (f) => '_rels/' + f + '.rels'), x = files.get(rp), out = [];
  if (!x) return out;
  const base = path.replace(/[^/]+$/, '');
  for (const m of dec.decode(x).matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    let t = attr(m[1], 'Target'); if (!t) continue;
    if (attr(m[1], 'TargetMode') !== 'External') { t = t.startsWith('/') ? t.slice(1) : base + t; const seg = []; for (const s of t.split('/')) { if (s === '..') seg.pop(); else if (s !== '.') seg.push(s); } t = seg.join('/'); }
    out.push({ id: attr(m[1], 'Id'), type: attr(m[1], 'Type'), target: t });
  }
  return out;
}
const textOf = (xml) => unx([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(''));

// Rewrite text paragraph by paragraph: fn(text) -> new text (or the same to leave it). Formatting of the first run is kept.
function mapParagraphs(xml, fn) {
  return xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (all, inner) => {
    const runs = [...inner.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)]; if (!runs.length) return all;
    const text = runs.map((r) => unx(/<a:t>([^<]*)<\/a:t>/.exec(r[1])?.[1] || '')).join('');
    const nt = fn(text, runs.length); if (nt == null || nt === text) return all;
    if (Array.isArray(nt)) {   // one string per run (keeps e.g. a superscript run)
      let i = 0; return '<a:p>' + inner.replace(/<a:r>([\s\S]*?)<\/a:r>/g, (r) => r.replace(/<a:t>[^<]*<\/a:t>/, `<a:t>${xesc(nt[i++] ?? '')}</a:t>`)) + '</a:p>';
    }
    let first = true;
    return '<a:p>' + inner.replace(/<a:r>([\s\S]*?)<\/a:r>/g, (r) => { if (!first) return ''; first = false; return r.replace(/<a:t>[^<]*<\/a:t>/, `<a:t>${xesc(nt)}</a:t>`); }) + '</a:p>';
  });
}

// ---------- tables ----------
function setCellText(tc, text) {
  return tc.replace(/<a:txBody>([\s\S]*?)<\/a:txBody>/, (all, body) => {
    const p = /<a:p>([\s\S]*?)<\/a:p>/.exec(body)?.[1] || '';
    const pPr = /<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/.exec(p)?.[0] || '';
    let rPr = /<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/.exec(p)?.[0];
    if (!rPr) { const e = /<a:endParaRPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:endParaRPr>)/.exec(p); rPr = e ? (e[3] != null ? `<a:rPr${e[1]}>${e[3]}</a:rPr>` : `<a:rPr${e[1]}/>`) : '<a:rPr lang="en-US"/>'; }
    const head = body.slice(0, body.indexOf('<a:p>') >= 0 ? body.indexOf('<a:p>') : body.length);
    return `<a:txBody>${head}<a:p>${pPr}${text === '' ? '' : `<a:r>${rPr}<a:t>${xesc(text)}</a:t></a:r>`}${rPr.replace(/^<a:rPr/, '<a:endParaRPr').replace(/<\/a:rPr>$/, '</a:endParaRPr>')}</a:p></a:txBody>`;
  });
}
function setCellFill(tc, hex) {
  if (!hex) return tc;
  return tc.replace(/<a:tcPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:tcPr>)/, (all, a, close, inner = '') => {
    const lines = []; let rest = inner.replace(/<a:(ln[LRTB]|lnTlToBr|lnBlToTr)\b[\s\S]*?<\/a:\1>|<a:(ln[LRTB]|lnTlToBr|lnBlToTr)\b[^>]*\/>/g, (m) => { lines.push(m); return ''; });
    rest = rest.replace(/<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[\s\S]*?<\/a:\1>|<a:(noFill|grpFill)\s*\/>/g, '');
    return `<a:tcPr${a}>${lines.join('')}<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>${rest}</a:tcPr>`;
  });
}
// rows: [{ cells: [text...], fills: [hex|null...], kind }] ; pick(kind) -> index of the existing row to copy formatting from
function fillTable(xml, rows, pick, maxHeightEmu) {
  return xml.replace(/<a:tbl>([\s\S]*?)<\/a:tbl>/, (all, tbl) => {
    const trs = [...tbl.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((m) => m[0]);
    const head = trs[0], first = tbl.indexOf(trs[0]), last = tbl.lastIndexOf(trs[trs.length - 1]) + trs[trs.length - 1].length;
    let out = rows.map((row) => {
      const src = trs[pick(row.kind, trs.length)] || trs[1] || trs[0];
      let i = 0;
      return src.replace(/<a:extLst>[\s\S]*?<\/a:extLst>(?=<\/a:tr>)/, '').replace(/<a:tc\b[\s\S]*?<\/a:tc>/g, (tc) => {
        const k = i++; let x = setCellText(tc, row.cells[k] ?? '');
        if (row.fills && row.fills[k]) x = setCellFill(x, row.fills[k]);
        return x;
      });
    });
    // shrink rows (and their text) when a long table would run off the slide
    if (maxHeightEmu) {
      const h = (tr) => +(attr(tr, 'h') || 0), total = h(head) + out.reduce((s, tr) => s + h(tr), 0);
      if (total > maxHeightEmu) {
        const k = Math.max(0.55, (maxHeightEmu - h(head)) / (total - h(head)));
        out = out.map((tr) => tr.replace(/<a:tr h="(\d+)"/, (m, v) => `<a:tr h="${Math.round(+v * k)}"`).replace(/\ssz="(\d+)"/g, (m, v) => ` sz="${Math.max(700, Math.round((+v * Math.min(1, k * 1.15)) / 100) * 100)}"`));
      }
    }
    return '<a:tbl>' + tbl.slice(0, first) + head + out.join('') + tbl.slice(last) + '</a:tbl>';
  });
}
const tableRows = (xml) => [...(/<a:tbl>([\s\S]*?)<\/a:tbl>/.exec(xml)?.[1] || '').matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((m) => [...m[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((c) => textOf(c[0]).trim()));

// ---------- charts ----------
const CHART_SLOTS = {   // positions (points) from the FY27 FW1 deck
  ftl: [['Overall(WITH AC)', /^FTL-OTD/i, [28, 54, 422, 224]], ['Overall(NO AC)', /^FTL-OTD/i, [21, 284, 442, 238]], ['Overall(WITH AC)', /FTL Delay Code/i, [536, 54, 350, 224]]],
  ltl: [['Overall(WITH AC)', /^LTL-OTD/i, [57, 48, 401, 276]], ['Overall(NO AC)', /^LTL-OTD/i, [518, 59, 401, 277]], ['@s8', /LTL Delay Code/i, [57, 336, 406, 179]]],
  hawb: [['Overall HAWB(WITH AC)', /^FTL-OTD/i, [28, 51, 439, 233]], ['Overall HAWB(WITH AC)', /^LTL-OTD/i, [520, 48, 354, 244]], ['Overall HAWB(WITH AC)', /FTL Delay Code/i, [28, 315, 290, 173]], ['Overall HAWB(WITH AC)', /LTL Delay Code/i, [418, 287, 513, 215]]],
};
const sameSheet = (a, b) => a.replace(/\s+/g, '').toUpperCase() === b.replace(/\s+/g, '').toUpperCase();
// Pasted chart pictures shrank text along with the chart; a live chart keeps its point sizes, so scale them the same way.
function scaleChartText(x, k) {
  if (!(k > 0) || k >= 0.98) return x;
  x = x.replace(/(<a:(?:defRPr|rPr|endParaRPr)\b[^>]*?\ssz=")(\d+)(")/g, (m, a, v, b) => `${a}${Math.max(500, Math.round((+v * k) / 50) * 50)}${b}`);
  // text with no size set uses the chart's default (10 pt): give the chart a scaled default
  const end = x.lastIndexOf('</c:chart>'), tail = x.slice(end);
  const def = `sz="${Math.max(500, Math.round((1000 * k) / 50) * 50)}"`;
  if (/^<\/c:chart>(?:<c:spPr>[\s\S]*?<\/c:spPr>)?<c:txPr>/.test(tail)) {
    return x.slice(0, end) + tail.replace(/<c:txPr>[\s\S]*?<\/c:txPr>/, (t) => t.replace(/<a:defRPr\b(?![^>]*\ssz=)/, `<a:defRPr ${def}`));
  }
  const at = end + '</c:chart>'.length + (/^<\/c:chart><c:spPr>[\s\S]*?<\/c:spPr>/.exec(tail)?.[0].length - 10 || 0);
  return x.slice(0, at) + `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr ${def}/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>` + x.slice(at);
}
function pptChartXml(x) {
  // an Excel chart part becomes a PowerPoint chart part: link it to its embedded workbook, drop Excel-only shapes
  x = x.replace(/<c:userShapes\b[^>]*\/>/g, '').replace(/<c:externalData\b[\s\S]*?<\/c:externalData>/g, '');
  if (!/xmlns:r=/.test(x.slice(0, 600))) x = x.replace('<c:chartSpace', '<c:chartSpace xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  const ext = '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>';
  const i = x.search(/<c:printSettings\b|<c:userShapes\b/);
  if (i >= 0) return x.slice(0, i) + ext + x.slice(i);
  const tail = /<c:extLst>(?:(?!<c:extLst>)[\s\S])*<\/c:extLst>\s*<\/c:chartSpace>\s*$/.exec(x);
  if (tail && x.lastIndexOf('</c:chart>') < tail.index) return x.slice(0, tail.index) + ext + x.slice(tail.index);
  return x.replace(/<\/c:chartSpace>\s*$/, ext + '</c:chartSpace>');
}

// opts: { sel, fileDate (ms), ppt (pptTables result), s8: 'no' | 'with', charts: [{ sheet, title, xml }], xlsx: Uint8Array }
export async function buildDeck(deckBlob, opts) {
  const files = await unzip(deckBlob), warnings = [];
  const { sel, ppt } = opts, Q = quarterOf(sel.fy, sel.wk), yy = String(sel.fy).slice(2);
  const presRels = rels(files, 'ppt/presentation.xml'), pres = dec.decode(files.get('ppt/presentation.xml'));
  const slides = [...pres.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)].map((m) => presRels.find((r) => r.id === m[1])?.target).filter((p) => p && files.has(p));
  let ct = dec.decode(files.get('[Content_Types].xml'));
  let chartNo = Math.max(0, ...[...files.keys()].map((n) => +(/^ppt\/charts\/chart(\d+)\.xml$/.exec(n)?.[1] || 0)));
  const d = new Date(opts.fileDate), day = d.getUTCDate();
  const done = new Set();

  for (const path of slides) {
    let x = dec.decode(files.get(path)); const all = textOf(x);
    let role = null;
    if (/Operational Business Review/i.test(all)) role = 'title';
    else if (/Delay Code with Delay Owner/i.test(all)) role = 's5';
    else if (/Performance By HAWB/i.test(all)) role = 's6';
    else if (/Performance Details/i.test(all) && /Units Due/i.test(all)) role = 's4';
    else if (/Overall FTL OTD/i.test(all)) role = 'ftl';
    else if (/Overall LTL OTD/i.test(all)) role = 'ltl';
    else if (/Delays by HAWB/i.test(all)) role = 'hawb';
    if (!role || (opts.only && !opts.only.includes(role))) continue;   // only: for testing
    done.add(role);

    // week labels in titles and text
    x = mapParagraphs(x, (t) => t
      .replace(/FY\s?\d{2}\s+FW\s?\d{1,2}/gi, `FY${yy} FW${sel.wk}`)
      .replace(/Q\d\s+Week\s+\d{1,2}\s+FY(\s?\d{2})?/gi, `Q${Q.q} Week ${Q.qweek} FY${yy}`)
      .replace(/(^|[^Q\d]\s)Week\s+\d{1,2}\s+FY\s?\d{2}/gi, (m, pre) => `${pre}Week ${sel.wk} FY${yy}`)
      .replace(/Q\d\s+W\d{1,2}\s+FY\s?\d{2}/gi, `Q${Q.q} W${Q.qweek} FY${yy}`));

    if (role === 'title') {
      // "October 07th, 2026" (the "th, 2026" part is often a superscript run)
      x = mapParagraphs(x, (t, n) => {
        if (!/^\s*[A-Z][a-z]+\s+\d{1,2}\s*(st|nd|rd|th)?,?\s*\d{4}\s*$/.test(t)) return null;
        const md = `${MONTHS[d.getUTCMonth()]} ${String(day).padStart(2, '0')}`, rest = `${ordinal(day)}, ${d.getUTCFullYear()}`;
        return n >= 2 ? [md, rest, ...Array(n - 2).fill('')] : `${md}${rest}`;
      });
    } else if (role === 's4' || role === 's6') {
      const units = role === 's4', rowsData = units ? ppt.s4 : ppt.s6;
      const rows = rowsData.map((o) => {
        const q = o.label === 'QTD', none = o.due == null;
        const a = adj(o), u = unadj(o);
        const cells = units
          ? [o.label, none ? '' : fmtInt(o.due), none ? '' : u == null ? '0' : pct(u, q ? 1 : 2), none ? '' : a == null ? '100%' : pct(a, 2)]
          : [o.label, none ? '' : fmtInt(o.due), none ? '' : u == null ? '0' : pct(u, 1), none ? '' : a == null ? '100' : pct(a, 1)];
        return { cells, fills: [null, null, null, none ? null : goalFill(a == null ? 1 : a)], kind: q ? 'qtd' : 'row' };
      });
      const n0 = tableRows(x).length;
      // stay within last week's table height, so the table never runs into the goal legend below it
      const oldH = [...(/<a:tbl>[\s\S]*?<\/a:tbl>/.exec(x)?.[0] || '').matchAll(/<a:tr h="(\d+)"/g)].reduce((s, m) => s + +m[1], 0);
      x = fillTable(x, rows, (kind, n) => (kind === 'qtd' ? n - 1 : 1), oldH || null);
      if (!n0) warnings.push(`Slide "${all.slice(0, 40)}": no table found`);
    } else if (role === 's5') {
      const old = tableRows(x);
      const groupIdx = old.findIndex((r, i) => i > 0 && !r[1] && r[2]), itemIdx = old.findIndex((r, i) => i > 0 && r[1]);
      const rows = ppt.s5.map((r) => ({ cells: [r.label, r.lane, String(r.hawb), fmtInt(r.units).replace(/,/g, '')], kind: r.head ? 'group' : 'item' }));
      if (!rows.length) rows.push({ cells: ['No carrier delays this week', '', '', ''], kind: 'item' });
      const frame = /<p:graphicFrame>(?:(?!<\/p:graphicFrame>)[\s\S])*?<a:tbl>[\s\S]*?<\/p:graphicFrame>/.exec(x)?.[0] || '';
      const fy = +(/<a:off x="-?\d+" y="(-?\d+)"/.exec(frame)?.[1] || 0);
      x = fillTable(x, rows, (kind) => (kind === 'group' ? (groupIdx > 0 ? groupIdx : 1) : itemIdx > 0 ? itemIdx : 2), 6858000 - fy - 25 * EMU);
    } else {
      // slides 7-9: remove the pasted pictures and old charts, add this week's charts from the template
      // (OLE objects such as the hidden "Object 8" carry their own <p:pic> preview, so graphic frames are set aside first)
      const kept = [];
      x = x.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g, (g) => (/<c:chart\b/.test(g) ? '' : `\u0000GF${kept.push(g) - 1}\u0000`));
      x = x.replace(/<p:pic>[\s\S]*?<\/p:pic>/g, '').replace(/\u0000GF(\d+)\u0000/g, (m, i) => kept[+i]);
      if (role === 'ltl') {   // the comments box: BK/HB split and the late reasons, largest first
        const lines = slide8Text(ppt.s8[opts.s8 === 'with' ? 'withAC' : 'noAC']);
        x = x.replace(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g, (sp) => {
          if (!/For BK|Late booking|Flight Related/i.test(textOf(sp))) return sp;
          return sp.replace(/<p:txBody>([\s\S]*?)<\/p:txBody>/, (a, body) => {
            const p = /<a:p>[\s\S]*?<\/a:p>/.exec(body)?.[0] || '<a:p></a:p>';
            const head = body.slice(0, body.indexOf('<a:p>'));
            return `<p:txBody>${head}${lines.map((l) => mapParagraphs(p, () => l)).join('')}</p:txBody>`;
          });
        });
      }
      const relsPath = path.replace(/[^/]+$/, (f) => '_rels/' + f + '.rels');
      let relXml = dec.decode(files.get(relsPath) || enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'));
      let rid = Math.max(0, ...[...relXml.matchAll(/Id="rId(\d+)"/g)].map((m) => +m[1]));
      let id = Math.max(1, ...[...x.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => +m[1]));
      let frames = '';
      for (const [sheet0, re, box] of CHART_SLOTS[role]) {
        const sheet = sheet0 === '@s8' ? (opts.s8 === 'with' ? 'Overall(WITH AC)' : 'Overall(NO AC)') : sheet0;
        const ch = opts.charts.find((c) => sameSheet(c.sheet, sheet) && re.test(c.title.trim()));
        if (!ch) { warnings.push(`No "${re.source}" chart on the "${sheet}" tab of the template`); continue; }
        chartNo++; rid++; id++;
        const cp = `ppt/charts/chart${chartNo}.xml`, ep = `ppt/embeddings/Microsoft_Excel_Worksheet${chartNo}.xlsx`;
        const [, , bw, bh] = box.map((v) => v * EMU), k = ch.size ? Math.max(0.45, Math.min(1, bw / ch.size.w, bh / ch.size.h)) : 1;
        files.set(cp, enc.encode(pptChartXml(scaleChartText(ch.xml, k))));
        files.set(`ppt/charts/_rels/chart${chartNo}.xml.rels`, enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet${chartNo}.xlsx"/></Relationships>`));
        files.set(ep, opts.xlsx);
        ct = ct.replace('</Types>', `<Override PartName="/${cp}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
        relXml = relXml.replace('</Relationships>', `<Relationship Id="rId${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNo}.xml"/></Relationships>`);
        const [l, t, w, h] = box.map((v) => Math.round(v * EMU));
        frames += `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${xesc(ch.sheet + ' ' + ch.title.trim())}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${l}" y="${t}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
      }
      x = x.replace('</p:spTree>', frames + '</p:spTree>');
      files.set(relsPath, enc.encode(relXml));
    }
    files.set(path, enc.encode(x));
  }
  if (!/Extension="xlsx"/i.test(ct)) ct = ct.replace(/(<Types\b[^>]*>)/, '$1<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>');
  files.set('[Content_Types].xml', enc.encode(ct));
  for (const [r, label] of [['title', 'title slide'], ['s4', 'OTD by origin table'], ['s5', 'carrier delay table'], ['s6', 'HAWB table'], ['ftl', 'FTL charts slide'], ['ltl', 'LTL charts slide'], ['hawb', 'HAWB charts slide']])
    if (!done.has(r)) warnings.push(`Last week's deck has no ${label}; it was left out`);
  const blob = await zip(files);
  return { blob: new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }), warnings };
}

export const deckName = (sel) => `Apple-CEVA FW${sel.wk} FY${String(sel.fy).slice(2)} Dom Weekly OTP Report.pptx`;
export async function isDeck(blob) {
  try { const files = await unzip(blob); return files.has('ppt/presentation.xml'); } catch (e) { return false; }
}
