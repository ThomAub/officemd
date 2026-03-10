# office-md

Fast Office document extraction for LLMs and agents. Native Node/Bun bindings for converting DOCX, XLSX, CSV, PPTX, and PDF into clean markdown, structured JSON IR, and Docling output.

## Install

```bash
npm install office-md
# or
bun add office-md
```

For a one-shot run without installing:

```bash
npx office-md markdown report.docx
bunx office-md markdown report.docx
```

## CLI

```bash
office-md markdown report.docx
office-md markdown budget.xlsx --sheets "Summary,Q1"
office-md render report.docx
office-md diff old.docx new.docx
office-md diff old.docx new.docx --html
```

## SDK

```js
import { readFileSync } from "node:fs";
import { markdownFromBytes, extractIrJson, doclingFromBytes } from "office-md";

const content = readFileSync("report.docx");

// Markdown
console.log(markdownFromBytes(content, "docx"));

// Structured JSON IR
console.log(extractIrJson(content, "docx"));

// Docling JSON
console.log(doclingFromBytes(content, "docx"));
```

### API

- `detectFormat(content)` - detect document format from bytes
- `extractIrJson(content, format?)` - extract intermediate representation as JSON
- `markdownFromBytes(content, format?, options...)` - render as markdown
- `markdownFromBytesBatch(contents, format?, workers?, options...)` - parallel markdown rendering
- `extractSheetNames(content)` - list XLSX sheet names
- `extractTablesIrJson(content, options...)` - extract XLSX table data as JSON
- `extractCsvTablesIrJson(content, options...)` - extract CSV table data as JSON
- `inspectPdfJson(content)` - PDF diagnostics as JSON
- `inspectPdfFontsJson(content)` - PDF font information as JSON
- `doclingFromBytes(content, format?)` - convert to Docling JSON

## Supported Formats

| Format     | Extension | Markdown | JSON IR | Docling |
| ---------- | --------- | -------- | ------- | ------- |
| Word       | .docx     | yes      | yes     | yes     |
| Excel      | .xlsx     | yes      | yes     | yes     |
| CSV        | .csv      | yes      | yes     | -       |
| PowerPoint | .pptx     | yes      | yes     | yes     |
| PDF        | .pdf      | yes      | yes     | -       |

## License

MIT
