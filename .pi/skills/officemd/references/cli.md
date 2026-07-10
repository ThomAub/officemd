# OfficeMD agent CLI reference

## Bounded XLSX inspection

Write a query file such as:

```json
{
  "kind": "xlsx_range",
  "sheet": "Sheet1",
  "range": "A1:I20",
  "include": {
    "values": true,
    "formulas": true,
    "number_formats": true
  }
}
```

Run:

```bash
officemd inspect input.xlsx \
  --agent-query inspect-range.json \
  --output-format json \
  --pretty
```

Other query shapes:

```json
{"kind":"xlsx_sheets"}
{"kind":"pdf_pages","pages":[1,2],"include":{"markdown":true}}
{"kind":"pptx_slides","start":1,"end":3}
{"kind":"document_summary"}
```

## Typed XLSX patch

`value` is an untagged JSON string, number, boolean, or `null`. Formula strings may start with `=`; OfficeMD normalizes the stored formula.

```json
{
  "version": "v1",
  "request_id": "manifest-2025-06-25",
  "operations": [
    {
      "op": "set_xlsx_cell_value",
      "target": {
        "sheet": "Sheet1",
        "address": "A3"
      },
      "value": "A-1001"
    },
    {
      "op": "set_xlsx_cell_value",
      "target": {
        "sheet": "Sheet1",
        "address": "G3"
      },
      "value": 5.95
    },
    {
      "op": "set_xlsx_cell_formula",
      "target": {
        "sheet": "Sheet1",
        "address": "I3"
      },
      "formula": "=H3-G3"
    }
  ]
}
```

Apply the plan:

```bash
officemd apply template.xlsx \
  --patch patch.json \
  --output completed.xlsx \
  --expected-source-sha256 SOURCE_SHA256 \
  --output-format json \
  --pretty
```

The output path must differ from the input and must not exist. A successful apply writes `<output>.officemd.json` as provenance and idempotency evidence.

## Verification and evidence

```bash
officemd verify completed.xlsx \
  --checks structure,formula-references,visual-render \
  --render-output-dir evidence/verify \
  --output-format json \
  --pretty

officemd render-artifact completed.xlsx \
  --output-dir evidence/render \
  --output-format json \
  --pretty

officemd diff-artifact template.xlsx completed.xlsx \
  --semantic \
  --rendered \
  --output-dir evidence/diff \
  --output-format json \
  --pretty
```

For PDF rendering, OfficeMD uses Poppler directly. For DOCX, XLSX, and PPTX, it uses LibreOffice to produce PDF and Poppler to produce PNG evidence.
