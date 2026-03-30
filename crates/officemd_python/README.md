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

## Typed OOXML patching with reports

```python
import officemd
from pathlib import Path

content = Path("report.docx").read_bytes()
patch = officemd.DocxPatch(
    scoped_replacements=[
        officemd.ScopedDocxReplace(
            officemd.DocxTextScope.ALL_TEXT,
            officemd.TextReplace("word", "term"),
        )
    ]
)

single = officemd.patch_docx_with_report(content, patch)
print(single.report.replacements_applied)

batch = officemd.patch_docx_batch_with_report([content, content], patch, workers=4)
for item in batch:
    print(item.report.parts_scanned, item.report.parts_modified, item.report.replacements_applied)
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
