// Web Worker: runs WASM conversion off the main thread so the UI never freezes.
import init, { convert_to_markdown } from '../pkg/officemd_wasm.js';

let ready = false;

async function setup() {
  await init();
  ready = true;
  postMessage({ type: 'ready' });
}

setup();

onmessage = (e) => {
  if (!ready) {
    postMessage({ type: 'error', error: 'WASM module is still loading' });
    return;
  }

  const { bytes, id } = e.data;
  try {
    const start = performance.now();
    // convert_to_markdown returns [format, markdown] in a single pass —
    // no need to call detect_format separately (avoids re-parsing the ZIP).
    const [fmt, markdown] = convert_to_markdown(bytes);
    const elapsed = performance.now() - start;

    postMessage({ type: 'result', id, markdown, fmt, elapsed });
  } catch (err) {
    postMessage({ type: 'error', id, error: err.message || String(err) });
  }
};
