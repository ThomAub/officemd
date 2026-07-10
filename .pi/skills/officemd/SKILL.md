---
name: officemd
description: Safely inspect, edit, render, verify, and compare DOCX, XLSX, PPTX, CSV, and PDF artifacts with the OfficeMD CLI. Use for bounded document understanding, auditable Office edits, visual evidence, formula checks, or artifact diffs.
---

# OfficeMD artifact workflow

Use OfficeMD as the primary document interface. Keep source files immutable and deliver a new native artifact.

## Establish capability

1. Resolve the executable with `command -v officemd`. If it is unavailable, stop and report that OfficeMD is required.
2. Run `officemd probe <file> --output-format json --pretty` for every input artifact.
3. Read the reported format, SHA-256 fingerprint, mutable operations, render backend, verification checks, and risks. Do not assume that extraction, mutation, or rendering is supported.

## Inspect narrowly

- Start with `officemd inspect <file> --output-format json --pretty` to discover sheets, pages, or slides.
- For XLSX, inspect only relevant ranges with an `xlsx_range` agent query. Include values, formulas, and number formats before deciding edits.
- Prefer one query covering each relevant used range. Do not issue overlapping range queries when the first result already contains the required cells.
- Filter inspection JSON to non-empty findings when a range contains many blank cells. Keep model context bounded.
- For PDF, inspect likely pages with a `pdf_pages` agent query. If the probe reports `pdf_may_require_ocr` and bounded extraction is empty, do not infer content from background knowledge. Render relevant pages and use image understanding if available, otherwise report that OCR is required.
- Render formatted templates and image-heavy sources before editing. Semantic extraction does not expose every visual object.
- Do not re-read successfully extracted values through raw OOXML, `openpyxl`, or another document library. Use a fallback only when OfficeMD explicitly cannot expose required source content, and state the limitation.

Read `{baseDir}/references/cli.md` for query and patch JSON shapes.

## Plan and apply edits

1. Determine exact stable locators. For XLSX use sheet names and absolute A1 addresses.
2. Create one typed patch plan containing every intended operation. Use formulas for derived spreadsheet values.
3. Use the source SHA-256 from `probe` as `--expected-source-sha256`.
4. Apply to a new path with `officemd apply`. Never overwrite the source and never pre-create the output.
5. Treat rejected, failed, unsupported, or ambiguous operations as failures. Do not silently switch to broad text replacement.

OfficeMD preserves an existing XLSX cell's style when setting its value or formula. Prefer filling a supplied template over rebuilding it.
Preserve sheet names, print settings, and template structure unless the user asks to change them. Leave unavailable optional fields blank instead of writing placeholders such as `N/A` or `Not in source`.

## Verify and render

1. Inspect the exact edited range and reconcile it with the source data and requested calculations.
2. Run `officemd verify <output> --checks structure,formula-references,visual-render --render-output-dir <dir> --output-format json --pretty` for XLSX. Select only applicable checks for other formats.
3. Read the rendered PNG evidence. A successful render proves only that rendering ran; visually check clipping, missing columns, broken layout, and blank output.
4. Use `officemd diff-artifact <source> <output> --semantic --rendered --output-dir <dir> --output-format json --pretty` when the expected change needs an auditable before/after comparison.
5. Deliver the native artifact, verification report, and evidence paths. State any non-runnable check or renderer limitation.

## Safety rules

- Never mutate in place.
- Never guess a cell, page, slide, or shape locator.
- Never claim formulas were recalculated merely because references are valid. OfficeMD reports formula evaluation as unavailable when it cannot prove cached results.
- Never treat rendered PNGs as deliverables.
- Never conceal `pdf_may_require_ocr`, missing renderer, or unsupported mutation diagnostics.
- Never install a second document-processing stack when the reported OfficeMD capabilities cover the task.
