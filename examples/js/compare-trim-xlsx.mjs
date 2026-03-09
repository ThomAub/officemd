/**
 * Compare XLSX markdown output with compact (trim) vs human (no-trim).
 *
 * Usage:
 *   node examples/js/compare-trim-xlsx.mjs <path.xlsx>
 *   bun run examples/js/compare-trim-xlsx.mjs <path.xlsx>
 */
import { readFileSync } from "node:fs";
import { basename, resolve } from "node:path";

import { loadNativeBinding } from "./load-addon.mjs";

const addon = loadNativeBinding();

const filePath = process.argv[2];
if (!filePath) {
  console.error("Usage: compare-trim-xlsx.mjs <path.xlsx>");
  process.exit(1);
}

const absPath = resolve(filePath);
const content = readFileSync(absPath);

const mdCompact = addon.markdownFromBytes(
  content,
  ".xlsx",
  false,
  true,
  true,
  "compact",
);
const mdHuman = addon.markdownFromBytes(
  content,
  ".xlsx",
  false,
  true,
  true,
  "human",
);

console.log(`=== ${basename(absPath)} ===\n`);

console.log(
  `--- compact/LlmCompact (trim_empty=true, ${mdCompact.length} chars) ---`,
);
console.log(mdCompact);

console.log(`--- human (trim_empty=false, ${mdHuman.length} chars) ---`);
console.log(mdHuman);

const saved = mdHuman.length - mdCompact.length;
const pct = mdHuman.length > 0 ? ((saved / mdHuman.length) * 100).toFixed(1) : "0.0";
console.log(`--- Savings: ${saved} chars (${pct}% reduction) ---`);
