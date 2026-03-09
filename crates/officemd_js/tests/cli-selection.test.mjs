import { spawnSync } from "node:child_process";
import assert from "node:assert/strict";
import test from "node:test";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..", "..", "..");
const cliPath = resolve(repoRoot, "crates", "officemd_js", "cli.js");
const fixturesDir = resolve(repoRoot, "examples", "data");

const pdfFixture = resolve(fixturesDir, "OpenXML_WhitePaper.pdf");
const xlsxFixture = resolve(fixturesDir, "showcase.xlsx");
const docxFixture = resolve(fixturesDir, "showcase.docx");

function runCli(args) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    encoding: "utf8",
    maxBuffer: 20 * 1024 * 1024,
  });
}

function pageHeadings(stdout) {
  return stdout.match(/^## Page:\s*\d+[ \t]*$/gm) ?? [];
}

function sheetHeadings(stdout) {
  return stdout.match(/^## Sheet:\s*[^\n\r]*$/gm) ?? [];
}

test("markdown --pages filters PDF pages with single/list/range selectors", () => {
  const full = runCli(["markdown", pdfFixture]);
  assert.equal(full.status, 0, full.stderr);
  assert.ok(pageHeadings(full.stdout).length >= 4, "expected multi-page PDF fixture");

  const selected = runCli(["markdown", pdfFixture, "--pages", "1,3-4"]);
  assert.equal(selected.status, 0, selected.stderr);
  assert.deepEqual(pageHeadings(selected.stdout), ["## Page: 1", "## Page: 3", "## Page: 4"]);
});

test("markdown/render spreadsheet selectors support --sheets, --pages alias, and merged selection", () => {
  const byName = runCli(["markdown", xlsxFixture, "--sheets", "Summary"]);
  assert.equal(byName.status, 0, byName.stderr);
  assert.deepEqual(sheetHeadings(byName.stdout), ["## Sheet: Summary"]);

  const byIndex = runCli(["render", xlsxFixture, "--sheets", "1"]);
  assert.equal(byIndex.status, 0, byIndex.stderr);
  assert.match(byIndex.stdout, /Sheet:\s*Sales/);
  assert.doesNotMatch(byIndex.stdout, /Sheet:\s*Summary/);

  const byPagesAlias = runCli(["markdown", xlsxFixture, "--pages", "2"]);
  assert.equal(byPagesAlias.status, 0, byPagesAlias.stderr);
  assert.deepEqual(sheetHeadings(byPagesAlias.stdout), ["## Sheet: Summary"]);

  const merged = runCli(["markdown", xlsxFixture, "--sheets", "Sales", "--pages", "2"]);
  assert.equal(merged.status, 0, merged.stderr);
  assert.deepEqual(sheetHeadings(merged.stdout), ["## Sheet: Sales", "## Sheet: Summary"]);
});

test("selectors return clear errors for invalid selectors, no matches, and unsupported formats", () => {
  const invalidPages = runCli(["markdown", pdfFixture, "--pages", "3-1"]);
  assert.notEqual(invalidPages.status, 0);
  assert.match(invalidPages.stderr, /invalid pages range/);

  const missingSheets = runCli(["markdown", xlsxFixture, "--sheets", "MissingSheet"]);
  assert.notEqual(missingSheets.status, 0);
  assert.match(missingSheets.stderr, /matched no sheets/);

  const unsupported = runCli(["markdown", docxFixture, "--pages", "1"]);
  assert.notEqual(unsupported.status, 0);
  assert.match(unsupported.stderr, /--pages is supported for PDF\/PPTX and XLSX\/CSV/);

  const sheetsOnPdf = runCli(["markdown", pdfFixture, "--sheets", "1"]);
  assert.notEqual(sheetsOnPdf.status, 0);
  assert.match(sheetsOnPdf.stderr, /--sheets is only supported for XLSX and CSV/);
});

test("index.js is importable as ESM and exposes named/default bindings", async () => {
  const entrypointPath = resolve(repoRoot, "crates", "officemd_js", "index.js");
  const mod = await import(entrypointPath);
  assert.equal(typeof mod.markdownFromBytes, "function");
  assert.equal(typeof mod.detectFormat, "function");
  assert.equal(typeof mod.default?.markdownFromBytes, "function");
});
