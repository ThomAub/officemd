import { readFileSync } from "node:fs";
import { extname, resolve } from "node:path";

import { loadNativeBinding } from "./load-addon.mjs";

const addon = loadNativeBinding();

const args = Bun.argv.slice(2);
const includeDocumentProperties = args.includes("--include-document-properties");
const filePath = args.find((arg) => !arg.startsWith("--"));
if (!filePath) {
  console.error(
    "Usage: bun run examples/js/bun-officemd.mjs <file.docx|file.xlsx|file.csv|file.pptx> [--include-document-properties]",
  );
  process.exit(1);
}

const absolutePath = resolve(filePath);
const bytes = readFileSync(absolutePath);
const format = extname(absolutePath).toLowerCase();
const explicitFormat = format.length > 0 ? format : undefined;

try {
  console.log("format:", addon.detectFormat(bytes));
} catch {
  if (explicitFormat === ".csv") {
    console.log("format:", ".csv (explicit)");
  } else {
    throw new Error("failed to detect format");
  }
}
const irJson = addon.extractIrJson(bytes, explicitFormat);
console.log("ir_json_length:", irJson.length);

const markdown = addon.markdownFromBytes(bytes, explicitFormat, includeDocumentProperties);
console.log("markdown_preview:");
console.log(markdown.slice(0, 400));

if (format === ".xlsx") {
  const sheetNames = addon.extractSheetNames(bytes);
  console.log("sheet_names:", JSON.stringify(sheetNames));

  const tablesIrJson = addon.extractTablesIrJson(bytes);
  console.log("tables_ir_json_length:", tablesIrJson.length);
}
