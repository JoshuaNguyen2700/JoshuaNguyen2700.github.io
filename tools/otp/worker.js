// Background worker for the otp tool: reads one Excel file at a time off the main thread, so the page stays
// responsive. The page runs several of these at once and combines what they send back.
import { readTask } from './parse.js';

self.onmessage = async (e) => {
  const { id, task } = e.data, progress = (p) => self.postMessage({ type: 'progress', ...p });
  try {
    self.postMessage({ type: 'done', id, result: await readTask(task, progress) });
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    progress({ file: task.name, stage: 'Skipped: ' + message, pct: 100, skipped: true });
    self.postMessage({ type: 'done', id, result: { name: task.name, error: message } });
  }
};
