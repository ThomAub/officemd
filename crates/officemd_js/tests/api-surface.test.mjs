import assert from "node:assert/strict";
import { existsSync, readdirSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..", "..", "..");
const require = createRequire(import.meta.url);
const packageDir = resolve(repoRoot, "crates", "officemd_js");

function nativeBindingPath() {
  const entries = readdirSync(packageDir).filter((name) => name.endsWith(".node"));
  assert.ok(entries.length > 0, "expected a built native binding in crates/officemd_js");

  const preferred = [
    `office-md.${process.platform}-${process.arch}.node`,
    `office-md.${process.platform}-${process.arch}-gnu.node`,
    `office-md.${process.platform}-${process.arch}-musl.node`,
    `office-md.${process.platform}-universal.node`,
  ];
  for (const candidate of preferred) {
    const fullPath = resolve(packageDir, candidate);
    if (existsSync(fullPath)) {
      return fullPath;
    }
  }

  return resolve(packageDir, entries[0]);
}

const expectedExports = [
  "applyOoxmlPatchJson",
  "createDocumentFromMarkdown",
  "detectFormat",
  "doclingFromBytes",
  "extractCsvTablesIrJson",
  "extractIrJson",
  "extractSheetNames",
  "extractTablesIrJson",
  "inspectPdfFontsJson",
  "inspectPdfJson",
  "markdownFromBytes",
  "markdownFromBytesBatch",
];

test("native binding exports the full binding surface", () => {
  const mod = require(nativeBindingPath());
  for (const name of expectedExports) {
    assert.equal(typeof mod[name], "function", `${name} should be exported`);
  }
});

test("createDocumentFromMarkdown creates OOXML bytes", () => {
  const mod = require(nativeBindingPath());
  const bytes = mod.createDocumentFromMarkdown("## Section: body\n\nHello\n", "docx");
  assert.ok(Buffer.isBuffer(bytes));
  assert.equal(bytes.subarray(0, 2).toString("utf8"), "PK");
});

test("detectFormat returns a string at runtime", () => {
  const mod = require(nativeBindingPath());
  const bytes = mod.createDocumentFromMarkdown("## Section: body\n\nHello\n", "docx");
  assert.equal(typeof mod.detectFormat(bytes), "string");
});

test("applyOoxmlPatchJson returns edited OOXML bytes", () => {
  const mod = require(nativeBindingPath());
  const bytes = mod.createDocumentFromMarkdown("## Section: body\n\nHello\n", "docx");
  const patch = JSON.stringify({
    edits: [{ part: "word/document.xml", from: "Hello", to: "Hello from JS" }],
  });
  const patched = mod.applyOoxmlPatchJson(bytes, patch);
  assert.ok(Buffer.isBuffer(patched));
  const markdown = mod.markdownFromBytes(patched, "docx", true);
  assert.match(markdown, /Hello from JS/);
});
