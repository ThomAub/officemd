// Web Worker: runs WASM conversion off the main thread so the UI never freezes.
import init, { convert_to_markdown, detect_format } from '../pkg/officemd_wasm.js';

let ready = false;

async function setup() {
  await init();
  ready = true;
  postMessage({ type: 'ready' });
}

setup();

onmessage = async (e) => {
  if (!ready) {
    postMessage({ type: 'error', error: 'WASM module is still loading' });
    return;
  }

  const { bytes, id } = e.data;
  try {
    const start = performance.now();
    const markdown = convert_to_markdown(bytes);
    const elapsed = performance.now() - start;

    let fmt;
    try { fmt = detect_format(bytes); } catch { fmt = '?'; }

    postMessage({ type: 'result', id, markdown, fmt, elapsed });
  } catch (err) {
    postMessage({ type: 'error', id, error: err.message || String(err) });
  }
};
