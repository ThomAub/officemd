import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { expect, test } from "bun:test";

import { loadNativeBinding } from "./load-addon.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..", "..");
const dataDir = resolve(repoRoot, "examples", "data");

const addon = loadNativeBinding();

function readFixture(filename) {
  return readFileSync(resolve(dataDir, filename));
}

test("detects showcase formats", () => {
  expect(addon.detectFormat(readFixture("showcase.docx"))).toBe(".docx");
  expect(addon.detectFormat(readFixture("showcase.xlsx"))).toBe(".xlsx");
  expect(addon.detectFormat(readFixture("showcase.pptx"))).toBe(".pptx");
});

test("extracts sheet names from showcase.xlsx", () => {
  const names = addon.extractSheetNames(readFixture("showcase.xlsx"));
  expect(names).toContain("Sales");
  expect(names).toContain("Summary");
});

test("extracts pptx IR and markdown from showcase.pptx", () => {
  const bytes = readFixture("showcase.pptx");
  const ir = JSON.parse(addon.extractIrJson(bytes, ".pptx"));
  expect(ir.kind).toBe("Pptx");

  const markdown = addon.markdownFromBytes(bytes, ".pptx");
  expect(markdown).toContain("Quarterly Review");
});

test("extracts csv IR and markdown from showcase.csv with explicit format", () => {
  const bytes = readFixture("showcase.csv");
  const ir = JSON.parse(addon.extractIrJson(bytes, ".csv"));
  expect(ir.kind).toBe("Xlsx");

  const markdown = addon.markdownFromBytes(bytes, ".csv");
  expect(markdown).toContain("## Sheet: Sheet1");
  expect(markdown).toContain("| Product | BaseAmount |");
});
