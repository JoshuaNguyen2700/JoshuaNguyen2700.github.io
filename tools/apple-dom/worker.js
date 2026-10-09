// Background worker for the apple-dom tool: reads the Excel files and writes the filled template off the main thread.
import { readAmr, weekChoices, fillTemplate, pptTables, isChartsTemplate, looksLikeAmrName, outName, quarterOf, fileDate } from './model.js';
import { buildDeck, deckName, isDeck } from './deck.js';

let AMR = null;    // { recs, name, weeks }
let LAST = null;   // last filled template: { sel, charts, xlsx, ppt }, used to build the PowerPoint

self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.type === 'classify') {
      // Which dropped file is the APPLE_AMR export, the Charts Template, or last week's PowerPoint?
      const kinds = [];
      for (const f of m.files) kinds.push(/\.pptx$/i.test(f.name) ? ((await isDeck(f)) ? 'deck' : 'other') : looksLikeAmrName(f.name) ? 'amr' : (await isChartsTemplate(f)) ? 'template' : 'amr');
      self.postMessage({ type: 'classified', kinds });
    } else if (m.type === 'amr') {
      let read = 0;
      const { recs, name } = await readAmr(m.file, m.file.name, (n) => { read += n; self.postMessage({ type: 'progress', done: read, total: m.file.size }); });
      if (!recs.length) throw new Error('has no rows with an Oem');
      const weeks = weekChoices(recs, name);
      AMR = { recs, name, weeks };
      self.postMessage({ type: 'amr', name, rows: recs.length, weeks });
    } else if (m.type === 'run') {
      if (!AMR) throw new Error('Load the APPLE_AMR file first.');
      let res;
      if (m.template) res = await fillTemplate(m.template, AMR.recs, m.sel, m.templateName || m.template.name || '');
      else res = { blob: null, tabs: [], warnings: [], quarter: quarterOf(m.sel.fy, m.sel.wk), qtd: null };   // no Charts Template yet: PPT tables only
      const ppt = pptTables(AMR.recs, m.sel, res.qtd);
      LAST = res.blob ? { sel: m.sel, charts: res.charts, xlsx: new Uint8Array(await res.blob.arrayBuffer()), ppt } : null;
      self.postMessage({ type: 'result', sel: m.sel, blob: res.blob, fileName: res.blob ? outName(m.sel) : null, tabs: res.tabs, warnings: res.warnings, quarter: res.quarter, template: res.template || null, ppt });
    } else if (m.type === 'deck') {
      if (!LAST || LAST.sel.wk !== m.sel.wk || LAST.sel.fy !== m.sel.fy) throw new Error('Fill the template for this week first.');
      const out = await buildDeck(m.deck, { sel: LAST.sel, fileDate: fileDate(AMR.name), ppt: LAST.ppt, s8: m.s8, charts: LAST.charts, xlsx: LAST.xlsx });
      self.postMessage({ type: 'deck', blob: out.blob, fileName: deckName(LAST.sel), warnings: out.warnings });
    }
  } catch (err) {
    self.postMessage({ type: 'error', step: m.type, message: err && err.message ? err.message : String(err) });
  }
};
