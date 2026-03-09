# OfficeMD

Fast Office document extraction for LLMs and agents. Converts DOCX, XLSX, CSV, PPTX, and PDF into clean markdown, structured JSON IR, and Docling output.

- Native Rust core - fast, no runtime dependencies
- Three output modes: markdown, structured JSON IR, Docling JSON
- CLI and SDK for Python, Node/Bun, and Rust
- Sheet, slide, and page selection
- Document property extraction

## Quick Start

No install needed - run directly:

```bash
uvx officemd markdown report.docx
npx office-md markdown report.docx
bunx office-md markdown report.docx
```

## Install

### Python

```bash
uv tool install officemd
```

Or add as a dependency:

```bash
uv add officemd
```

### Node / Bun

```bash
npm install office-md
# or
bun add office-md
```

### Rust

```bash
cargo install officemd_cli
```

## CLI

All three surfaces expose a CLI named `officemd` (Python, Rust) or `office-md` (Node/Bun).

```bash
officemd markdown report.docx
officemd markdown budget.xlsx --sheets "Summary,Q1"
officemd markdown deck.pptx --pages 1-3
officemd render report.docx
officemd diff old.docx new.docx
```

The Rust CLI has additional subcommands:

```bash
officemd stream report.docx                    # stream to stdout (supports stdin via -)
officemd convert report.docx --output out.md   # write to file
officemd inspect report.pdf --output-format json --pretty
```

### Common options

| Flag | Description |
|------|-------------|
| `--format` | Force document format (docx, xlsx, csv, pptx, pdf) |
| `--pages` | Select pages/slides/sheets by index (e.g. "1,3-5") |
| `--sheets` | Select sheets by name or index (e.g. "Sales,1-2") |
| `--include-document-properties` | Include document metadata in output |
| `--markdown-style` | Output style: `compact` (default) or `human` |

## SDK

### Python

```python
from pathlib import Path
from officemd import extract_ir_json, markdown_from_bytes, docling_from_bytes

content = Path("report.docx").read_bytes()

print(markdown_from_bytes(content, format="docx"))
print(extract_ir_json(content, format="docx"))
print(docling_from_bytes(content, format="docx"))
```

### JavaScript

```js
import { readFileSync } from "node:fs";
import { markdownFromBytes, extractIrJson, doclingFromBytes } from "office-md";

const content = readFileSync("report.docx");

console.log(markdownFromBytes(content, "docx"));
console.log(extractIrJson(content, "docx"));
console.log(doclingFromBytes(content, "docx"));
```

### Rust

OfficeMD is a workspace of focused crates. Use them directly:

```toml
[dependencies]
officemd_docx = "0.1"
officemd_markdown = "0.1"
officemd_core = "0.1"
```

## Supported Formats

| Format | Extension | Markdown | JSON IR | Docling |
|--------|-----------|----------|---------|---------|
| Word | .docx | yes | yes | yes |
| Excel | .xlsx | yes | yes | yes |
| CSV | .csv | yes | yes | - |
| PowerPoint | .pptx | yes | yes | yes |
| PDF | .pdf | yes | yes | - |

## Development

```bash
cargo test --workspace
cargo clippy --workspace --all-targets -- -D warnings
```

For JS and Python tests, see [examples/README.md](examples/README.md).

## Acknowledgements

PDF extraction vendors [pdf-inspector](https://github.com/firecrawl/pdf-inspector) by Firecrawl (MIT).

## License

Apache 2.0 - see [LICENSE](LICENSE).
