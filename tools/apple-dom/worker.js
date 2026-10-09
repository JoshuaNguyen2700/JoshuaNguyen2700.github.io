// Background worker for the apple-dom tool: reads the Excel files and writes the filled template off the main thread.
import { readAmr, weekChoices, fillTemplate, pptTables, isChartsTemplate, looksLikeAmrName, outName, quarterOf } from './model.js';

let AMR = null;   // { recs, name, weeks }

self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.type === 'classify') {
      // Which dropped file is the APPLE_AMR export and which is the Charts Template?
      const kinds = [];
      for (const f of m.files) kinds.push(looksLikeAmrName(f.name) ? 'amr' : (await isChartsTemplate(f)) ? 'template' : 'amr');
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
      self.postMessage({ type: 'result', sel: m.sel, blob: res.blob, fileName: res.blob ? outName(m.sel) : null, tabs: res.tabs, warnings: res.warnings, quarter: res.quarter, template: res.template || null, ppt });
    }
  } catch (err) {
    self.postMessage({ type: 'error', step: m.type, message: err && err.message ? err.message : String(err) });
  }
};
