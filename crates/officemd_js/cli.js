#!/usr/bin/env node
/**
 * CLI entry point for office-md.
 *
 * Subcommands:
 *   render <file>           - extract markdown, render to terminal
 *   diff <file_a> <file_b>  - show colored diff of two files' markdown
 *   markdown <file>         - extract markdown, print plain to stdout
 *
 * Selection flags:
 *   --pages <selector>       - PDF/PPTX pages/slides (e.g. 1,3-5)
 *   --sheets <selector>      - XLSX/CSV sheet names/indices (e.g. Summary,1-2)
 */

import { existsSync, readFileSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { renderMarkdownToTerminal } from "./lib/render-ansi.js";
import { renderTerminalDiff } from "./lib/diff-terminal.js";
import { generateHtmlDiff, openInBrowser } from "./lib/diff-html.js";

const require = createRequire(import.meta.url);
const here = dirname(fileURLToPath(import.meta.url));

const PDF_PAGE_HEADING_RE = /^## Page:\s*(\d+)\s*$/;
const PPTX_SLIDE_HEADING_RE = /^## Slide\s+(\d+)(?:\s+-\s+.*)?\s*$/;
const SHEET_HEADING_RE = /^## Sheet:\s*(.+)\s*$/;

// ANSI color helpers for stderr output
const BOLD_RED = "\x1b[1;31m";
const BOLD_YELLOW = "\x1b[1;33m";
const BOLD_BLUE = "\x1b[1;34m";
const BOLD = "\x1b[1m";
const RESET = "\x1b[0m";

const SUPPORTED_FORMATS = ".docx, .xlsx, .csv, .pptx, .pdf";

function formatError(err, filePath) {
  const msg = err?.message || String(err);

  if (err?.code === "ENOENT" || msg.includes("ENOENT")) {
    return `File not found: ${filePath || msg}`;
  }

  if (msg.includes("ZIP error") || msg.includes("EOCD") || msg.includes("invalid Zip archive")) {
    return (
      "This file appears to be encrypted or not a valid OOXML document. " +
      "If the file is password-protected, remove the password and try again."
    );
  }

  if (msg.includes("Could not detect format")) {
    return `Could not detect document format. Supported formats: ${SUPPORTED_FORMATS}`;
  }

  return msg;
}

function warnScannedPdf(content, force, inspectPdfJson) {
  if (!inspectPdfJson) return;

  let diag;
  try {
    diag = JSON.parse(inspectPdfJson(content));
  } catch {
    return;
  }

  const classification = diag.classification;
  if (classification !== "Scanned" && classification !== "ImageBased") return;

  const confidence = diag.confidence || 0;
  const pageCount = diag.page_count || 0;
  const pagesNeedingOcr = diag.pages_needing_ocr || [];

  if (force) {
    process.stderr.write(
      `${BOLD_BLUE}Info:${RESET} PDF classified as ${BOLD}${classification}${RESET}` +
      ` (confidence: ${(confidence * 100).toFixed(0)}%, ${pageCount} page(s)). ` +
      `Forced extraction attempted - output may be empty or incomplete.\n`
    );
  } else {
    const ocrSummary = pagesNeedingOcr.length > 0
      ? `pages needing OCR: ${pagesNeedingOcr.join(", ")}`
      : `${pageCount} page(s)`;
    process.stderr.write(
      `${BOLD_YELLOW}Warning:${RESET} PDF classified as ${BOLD}${classification}${RESET}` +
      ` (confidence: ${(confidence * 100).toFixed(0)}%, ${ocrSummary}). ` +
      `No text could be extracted - this document likely needs OCR.\n` +
      `Hint: use ${BOLD}--force${RESET} to attempt extraction anyway.\n`
    );
  }
}

function isMusl() {
  if (!process.report || typeof process.report.getReport !== "function") {
    return false;
  }
  const { glibcVersionRuntime } = process.report.getReport().header;
  return !glibcVersionRuntime;
}

function getNativeCandidates() {
  const platform = process.platform;
  const arch = process.arch;

  if (platform === "darwin") {
    if (arch === "arm64") {
      return [
        { local: "office-md.darwin-universal.node", pkg: "office-md-darwin-universal" },
        { local: "office-md.darwin-arm64.node", pkg: "office-md-darwin-arm64" },
      ];
    }
    if (arch === "x64") {
      return [
        { local: "office-md.darwin-universal.node", pkg: "office-md-darwin-universal" },
        { local: "office-md.darwin-x64.node", pkg: "office-md-darwin-x64" },
      ];
    }
  }

  if (platform === "linux") {
    const musl = isMusl();
    if (arch === "x64") {
      return musl
        ? [{ local: "office-md.linux-x64-musl.node", pkg: "office-md-linux-x64-musl" }]
        : [{ local: "office-md.linux-x64-gnu.node", pkg: "office-md-linux-x64-gnu" }];
    }
    if (arch === "arm64") {
      return musl
        ? [{ local: "office-md.linux-arm64-musl.node", pkg: "office-md-linux-arm64-musl" }]
        : [{ local: "office-md.linux-arm64-gnu.node", pkg: "office-md-linux-arm64-gnu" }];
    }
    if (arch === "arm") {
      return musl
        ? [{ local: "office-md.linux-arm-musleabihf.node", pkg: "office-md-linux-arm-musleabihf" }]
        : [{ local: "office-md.linux-arm-gnueabihf.node", pkg: "office-md-linux-arm-gnueabihf" }];
    }
    if (arch === "riscv64") {
      return musl
        ? [{ local: "office-md.linux-riscv64-musl.node", pkg: "office-md-linux-riscv64-musl" }]
        : [{ local: "office-md.linux-riscv64-gnu.node", pkg: "office-md-linux-riscv64-gnu" }];
    }
    if (arch === "s390x") {
      return [{ local: "office-md.linux-s390x-gnu.node", pkg: "office-md-linux-s390x-gnu" }];
    }
  }

  if (platform === "win32") {
    if (arch === "x64") {
      return [{ local: "office-md.win32-x64-msvc.node", pkg: "office-md-win32-x64-msvc" }];
    }
    if (arch === "ia32") {
      return [{ local: "office-md.win32-ia32-msvc.node", pkg: "office-md-win32-ia32-msvc" }];
    }
    if (arch === "arm64") {
      return [{ local: "office-md.win32-arm64-msvc.node", pkg: "office-md-win32-arm64-msvc" }];
    }
  }

  if (platform === "freebsd" && arch === "x64") {
    return [{ local: "office-md.freebsd-x64.node", pkg: "office-md-freebsd-x64" }];
  }

  if (platform === "android") {
    if (arch === "arm64") {
      return [{ local: "office-md.android-arm64.node", pkg: "office-md-android-arm64" }];
    }
    if (arch === "arm") {
      return [{ local: "office-md.android-arm-eabi.node", pkg: "office-md-android-arm-eabi" }];
    }
  }

  return [];
}

function loadNativeBinding() {
  const candidates = getNativeCandidates();
  const loadErrors = [];

  for (const candidate of candidates) {
    if (candidate.local) {
      const localPath = join(here, candidate.local);
      if (existsSync(localPath)) {
        try {
          return require(localPath);
        } catch (err) {
          loadErrors.push(err);
        }
      }
    }
    if (candidate.pkg) {
      try {
        return require(candidate.pkg);
      } catch (err) {
        loadErrors.push(err);
      }
    }
  }

  const details = loadErrors.map((err) => (err?.message ? err.message : String(err))).join("\n");
  throw new Error(
    `Failed to load office-md native binding for ${process.platform}/${process.arch}` +
      (details ? `\n${details}` : "")
  );
}

const nativeBinding = loadNativeBinding();
const { markdownFromBytes, inspectPdfJson } = nativeBinding;

function parseOptions(args) {
  const flags = {
    includeDocumentProperties: false,
    useFirstRowAsHeader: true,
    includeHeadersFooters: true,
    markdownStyle: "compact",
    force: false,
  };
  const rest = [];

  let i = 0;
  while (i < args.length) {
    const arg = args[i];
    switch (arg) {
      case "--format":
        i++;
        if (!args[i]) {
          throw new Error("Error: --format requires a value");
        }
        flags.format = args[i];
        break;
      case "--include-document-properties":
        flags.includeDocumentProperties = true;
        break;
      case "--use-first-row-as-header":
        flags.useFirstRowAsHeader = true;
        break;
      case "--no-first-row-header":
        flags.useFirstRowAsHeader = false;
        break;
      case "--include-headers-footers":
        flags.includeHeadersFooters = true;
        break;
      case "--no-headers-footers":
        flags.includeHeadersFooters = false;
        break;
      case "--markdown-style":
        i++;
        if (!args[i]) {
          throw new Error("Error: --markdown-style requires a value");
        }
        flags.markdownStyle = args[i];
        break;
      case "--pages":
        i++;
        if (!args[i]) {
          throw new Error("Error: --pages requires a selector value");
        }
        flags.pages = args[i];
        break;
      case "--sheets":
        i++;
        if (!args[i]) {
          throw new Error("Error: --sheets requires a selector value");
        }
        flags.sheets = args[i];
        break;
      case "--force":
        flags.force = true;
        break;
      case "--html":
        rest.push(arg);
        break;
      case "--output":
        rest.push(arg);
        i++;
        rest.push(args[i]);
        break;
      default:
        rest.push(arg);
        break;
    }
    i++;
  }

  return { flags, rest };
}

function normalizeFormat(format) {
  if (!format) {
    return undefined;
  }
  const lower = format.toLowerCase();
  return lower.startsWith(".") ? lower : `.${lower}`;
}

function resolveFormat(filePath, explicitFormat) {
  const requestedFormat = normalizeFormat(explicitFormat);
  if (requestedFormat) {
    return requestedFormat;
  }
  return inferFormatFromPath(filePath);
}

function parseNumberSelector(raw, label) {
  const values = new Set();
  const tokens = raw.split(",");

  for (const tokenRaw of tokens) {
    const token = tokenRaw.trim();
    if (!token) {
      throw new Error(`Error: invalid ${label} selector "${raw}"`);
    }

    const rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const start = Number.parseInt(rangeMatch[1], 10);
      const end = Number.parseInt(rangeMatch[2], 10);
      if (start < 1 || end < 1 || start > end) {
        throw new Error(`Error: invalid ${label} range "${token}"`);
      }
      for (let i = start; i <= end; i++) {
        values.add(i);
      }
      continue;
    }

    if (/^\d+$/.test(token)) {
      const value = Number.parseInt(token, 10);
      if (value < 1) {
        throw new Error(`Error: invalid ${label} value "${token}"`);
      }
      values.add(value);
      continue;
    }

    throw new Error(
      `Error: invalid ${label} selector token "${token}" (expected integer or range like 3-5)`
    );
  }

  if (values.size === 0) {
    throw new Error(`Error: invalid ${label} selector "${raw}"`);
  }
  return values;
}

function parseSheetSelector(raw) {
  const indices = new Set();
  const names = new Set();
  const tokens = raw.split(",");

  for (const tokenRaw of tokens) {
    const token = tokenRaw.trim();
    if (!token) {
      throw new Error(`Error: invalid sheets selector "${raw}"`);
    }

    const rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const start = Number.parseInt(rangeMatch[1], 10);
      const end = Number.parseInt(rangeMatch[2], 10);
      if (start < 1 || end < 1 || start > end) {
        throw new Error(`Error: invalid sheets range "${token}"`);
      }
      for (let i = start; i <= end; i++) {
        indices.add(i);
      }
      continue;
    }

    if (/^\d+$/.test(token)) {
      const index = Number.parseInt(token, 10);
      if (index < 1) {
        throw new Error(`Error: invalid sheet index "${token}"`);
      }
      indices.add(index);
      continue;
    }

    names.add(token);
  }

  if (indices.size === 0 && names.size === 0) {
    throw new Error(`Error: invalid sheets selector "${raw}"`);
  }

  return { indices, names };
}

function buildSelection(flags) {
  const selection = {};
  if (flags.pages) {
    selection.pages = {
      raw: flags.pages,
    };
  }
  if (flags.sheets) {
    selection.sheets = {
      raw: flags.sheets,
    };
  }
  return selection;
}

function splitLines(markdown) {
  return markdown.split(/\r?\n/);
}

function collectSections(lines, headingPattern) {
  const headings = [];
  for (let i = 0; i < lines.length; i++) {
    const match = lines[i].match(headingPattern);
    if (match) {
      headings.push({ lineIndex: i, match });
    }
  }

  return headings.map((heading, i) => {
    const nextLineIndex = i + 1 < headings.length ? headings[i + 1].lineIndex : lines.length;
    return {
      heading,
      startLine: heading.lineIndex,
      endLine: nextLineIndex,
      content: lines.slice(heading.lineIndex, nextLineIndex).join("\n"),
    };
  });
}

function composeFilteredMarkdown(markdown, lines, sections, preludeEndLine = 0) {
  const parts = [];
  if (sections.length === 0) {
    return "";
  }

  const prelude = lines.slice(0, preludeEndLine).join("\n");
  if (prelude.trim().length > 0) {
    parts.push(prelude);
  }

  for (const section of sections) {
    parts.push(section.content);
  }

  let result = parts.join("\n");
  if (markdown.endsWith("\n")) {
    result += "\n";
  }
  return result;
}

function filterPageLikeSections(markdown, pages, headingPattern, label) {
  const lines = splitLines(markdown);
  const allSections = collectSections(lines, headingPattern).map((section) => ({
    ...section,
    number: Number.parseInt(section.heading.match[1], 10),
  }));

  if (allSections.length === 0) {
    throw new Error(`Error: could not find ${label} headings in markdown output`);
  }

  const selected = allSections.filter((section) => pages.numbers.has(section.number));
  if (selected.length === 0) {
    throw new Error(`Error: --pages selector "${pages.raw}" matched no ${label}`);
  }

  return composeFilteredMarkdown(markdown, lines, selected, allSections[0].startLine);
}

function filterSheetSections(markdown, sheets) {
  const lines = splitLines(markdown);
  const allSections = collectSections(lines, SHEET_HEADING_RE).map((section, index) => ({
    ...section,
    index: index + 1,
    name: section.heading.match[1].trim(),
  }));

  if (allSections.length === 0) {
    throw new Error("Error: could not find sheet headings in markdown output");
  }

  const selected = allSections.filter(
    (section) => sheets.parsed.indices.has(section.index) || sheets.parsed.names.has(section.name)
  );

  if (selected.length === 0) {
    throw new Error(`Error: --sheets selector "${sheets.raw}" matched no sheets`);
  }

  return composeFilteredMarkdown(markdown, lines, selected, allSections[0].startLine);
}

function applySelection(markdown, format, selection) {
  const normalizedFormat = normalizeFormat(format);
  const hasPageSelection = Boolean(selection.pages);
  const hasSheetSelection = Boolean(selection.sheets);

  if (!hasPageSelection && !hasSheetSelection) {
    return markdown;
  }

  if (!normalizedFormat) {
    throw new Error(
      "Error: cannot determine input format for selectors; pass --format when using --pages or --sheets"
    );
  }

  if (normalizedFormat === ".pdf" || normalizedFormat === ".pptx") {
    if (hasSheetSelection) {
      throw new Error("Error: --sheets is only supported for XLSX and CSV inputs");
    }
    if (!hasPageSelection) {
      return markdown;
    }
    const parsedPages = {
      raw: selection.pages.raw,
      numbers: parseNumberSelector(selection.pages.raw, "pages"),
    };
    if (normalizedFormat === ".pdf") {
      return filterPageLikeSections(markdown, parsedPages, PDF_PAGE_HEADING_RE, "pages");
    }
    return filterPageLikeSections(markdown, parsedPages, PPTX_SLIDE_HEADING_RE, "slides");
  }

  if (normalizedFormat === ".xlsx" || normalizedFormat === ".csv") {
    const selectorRaw = [selection.sheets?.raw, selection.pages?.raw].filter(Boolean).join(",");
    const parsedSheets = {
      raw: selectorRaw,
      parsed: parseSheetSelector(selectorRaw),
    };
    return filterSheetSections(markdown, parsedSheets);
  }

  if (hasSheetSelection) {
    throw new Error("Error: --sheets is only supported for XLSX and CSV inputs");
  }
  throw new Error("Error: --pages is supported for PDF/PPTX and XLSX/CSV inputs");
}

function extractMarkdown(filePath, opts, selection = {}) {
  const absPath = resolve(filePath);

  // Pre-flight checks
  if (!existsSync(absPath)) {
    throw Object.assign(new Error(`File not found: ${absPath}`), { code: "ENOENT" });
  }
  const stat = statSync(absPath);
  if (stat.size === 0) {
    throw new Error(`File is empty: ${absPath}`);
  }

  const content = readFileSync(absPath);
  const inferredFormat = resolveFormat(absPath, opts.format);
  const markdown = markdownFromBytes(
    content,
    inferredFormat,
    opts.includeDocumentProperties,
    opts.useFirstRowAsHeader,
    opts.includeHeadersFooters,
    opts.markdownStyle,
    opts.force || false,
  );

  // Post-hoc scanned PDF warning
  if (inferredFormat === ".pdf" && markdown.trim().length < 50) {
    warnScannedPdf(content, opts.force, inspectPdfJson);
  }

  return applySelection(markdown, inferredFormat, selection);
}

function inferFormatFromPath(path) {
  const extension = extname(path).toLowerCase();
  switch (extension) {
    case ".docx":
    case ".xlsx":
    case ".csv":
    case ".pptx":
    case ".pdf":
      return extension;
    default:
      return undefined;
  }
}

function usage(exitCode = 1) {
  console.log(`Usage: office-md <command> [options] <file(s)>

Commands:
  markdown <file>            Extract markdown, print to stdout
  render <file>              Extract markdown, render to terminal with ANSI
  diff <file_a> <file_b>     Diff markdown output of two files

Options:
  --format <docx|xlsx|csv|pptx|pdf>    Force document format
  --include-document-properties    Include document properties
  --use-first-row-as-header        Use first row as table header (default)
  --no-first-row-header            Do not use first row as table header
  --include-headers-footers        Include headers and footers (default)
  --no-headers-footers             Do not include headers and footers
  --markdown-style <compact|human> Markdown profile (default: compact)
  --pages <selector>                Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. 1,3-5)
  --sheets <selector>               Select XLSX/CSV sheets by name/index (e.g. Summary,1-2)
  --force                           Force extraction even for scanned/image-based PDFs

Diff-specific options:
  --html                           Generate HTML diff and open in browser
  --output <path>                  Output path for HTML diff`);
  process.exit(exitCode);
}

async function main() {
  const rawArgs = process.argv.slice(2);

  if (rawArgs.length === 0) {
    usage();
  }

  if (rawArgs[0] === "--help" || rawArgs[0] === "-h") {
    usage(0);
  }

  const command = rawArgs[0];
  const { flags, rest } = parseOptions(rawArgs.slice(1));
  const selection = buildSelection(flags);

  switch (command) {
    case "markdown": {
      const file = rest[0];
      if (!file) {
        console.error("Error: missing <file> argument");
        usage();
      }
      const md = extractMarkdown(file, flags, selection);
      process.stdout.write(md);
      break;
    }

    case "render": {
      const file = rest[0];
      if (!file) {
        console.error("Error: missing <file> argument");
        usage();
      }
      const md = extractMarkdown(file, flags, selection);
      const rendered = renderMarkdownToTerminal(md);
      process.stdout.write(rendered.endsWith("\n") ? rendered : `${rendered}\n`);
      break;
    }

    case "diff": {
      const htmlIndex = rest.indexOf("--html");
      const useHtml = htmlIndex !== -1;
      const filteredRest = rest.filter((_, i) => i !== htmlIndex);

      let outputPath;
      const outputIndex = filteredRest.indexOf("--output");
      let files = filteredRest;
      if (outputIndex !== -1) {
        outputPath = filteredRest[outputIndex + 1];
        files = filteredRest.filter((_, i) => i !== outputIndex && i !== outputIndex + 1);
      }

      const fileA = files[0];
      const fileB = files[1];
      if (!fileA || !fileB) {
        console.error("Error: diff requires two file arguments");
        usage();
      }

      const mdA = extractMarkdown(fileA, flags, selection);
      const mdB = extractMarkdown(fileB, flags, selection);

      if (useHtml) {
        const htmlPath = await generateHtmlDiff(mdA, mdB, outputPath);
        console.log(`HTML diff written to: ${htmlPath}`);
        await openInBrowser(htmlPath);
      } else {
        const diff = renderTerminalDiff(mdA, mdB);
        process.stdout.write(diff);
      }
      break;
    }

    default:
      console.error(`Unknown command: ${command}`);
      usage();
  }
}

main().catch((err) => {
  const message = formatError(err);
  process.stderr.write(`${BOLD_RED}Error:${RESET} ${message}\n`);
  process.exit(1);
});
