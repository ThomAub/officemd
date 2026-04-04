// Web Worker: runs WASM conversion off the main thread so the UI never freezes.
import init, { convert_to_markdown, convert_to_markdown_with_format } from '../pkg/officemd_wasm.js';

let ready = false;

async function setup() {
  await init();
  ready = true;
  postMessage({ type: 'ready' });
}

setup();

/** Extract file extension (without dot), lowercased. */
function extFromName(name) {
  const dot = name.lastIndexOf('.');
  return dot >= 0 ? name.slice(dot + 1).toLowerCase() : '';
}

onmessage = (e) => {
  if (!ready) {
    postMessage({ type: 'error', error: 'WASM module is still loading' });
    return;
  }

  const { bytes, id, fileName } = e.data;
  try {
    const start = performance.now();
    const ext = extFromName(fileName || '');

    let fmt, markdown;
    if (ext === 'csv') {
      // CSV cannot be auto-detected from bytes — pass the format explicitly.
      markdown = convert_to_markdown_with_format(bytes, 'csv');
      fmt = 'csv';
    } else {
      // convert_to_markdown returns [format, markdown] in a single pass —
      // no need to call detect_format separately (avoids re-parsing the ZIP).
      [fmt, markdown] = convert_to_markdown(bytes);
    }
    const elapsed = performance.now() - start;

    postMessage({ type: 'result', id, markdown, fmt, elapsed });
  } catch (err) {
    postMessage({ type: 'error', id, error: err.message || String(err) });
  }
};
