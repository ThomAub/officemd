# officemd

Fast Office document extraction for LLMs and agents. Converts DOCX, XLSX, CSV, PPTX, and PDF into clean markdown, structured JSON IR, and Docling output.

## Install

```bash
uv add officemd
# or
pip install officemd
```

For the CLI without adding to a project:

```bash
uvx officemd markdown report.docx
```

## CLI

```bash
officemd markdown report.docx
officemd markdown budget.xlsx --sheets "Summary,Q1"
officemd render report.docx
officemd diff old.docx new.docx
```

## SDK

```python
from pathlib import Path
from officemd import extract_ir_json, markdown_from_bytes, docling_from_bytes

content = Path("report.docx").read_bytes()

# Markdown
print(markdown_from_bytes(content, format="docx"))

# Structured JSON IR
print(extract_ir_json(content, format="docx"))

# Docling JSON
print(docling_from_bytes(content, format="docx"))
```

## Supported Formats

| Format | Extension | Markdown | JSON IR | Docling |
|--------|-----------|----------|---------|---------|
| Word | .docx | yes | yes | yes |
| Excel | .xlsx | yes | yes | yes |
| CSV | .csv | yes | yes | - |
| PowerPoint | .pptx | yes | yes | yes |
| PDF | .pdf | yes | yes | - |

## License

MIT
