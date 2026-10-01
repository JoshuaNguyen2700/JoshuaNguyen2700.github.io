// Background worker for the otp tool: reads the Excel files off the main thread so the page stays responsive.
import { buildFromFiles } from './parse.js';

self.onmessage = async (e) => {
  try {
    const data = await buildFromFiles(e.data.files, (p) => self.postMessage({ type: 'progress', ...p }));
    self.postMessage({ type: 'done', data });
  } catch (err) {
    self.postMessage({ type: 'error', message: err && err.message ? err.message : String(err) });
  }
};
