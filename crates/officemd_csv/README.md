# officemd_csv

CSV extraction helpers aligned with the `officemd_xlsx` table-to-markdown flow.

## Features
- Single-sheet IR extraction from CSV bytes.
- Table-focused IR (`extract_tables_ir*`) with synthetic `ColN` headers.
- Formula footnotes from cells starting with `=`.
- Markdown rendering through `officemd_markdown`.
- Optional document properties block (`source_format=csv`, delimiter metadata).

## Rust Usage

```rust
use officemd_csv::{extract_tables_ir, markdown_from_bytes};
use officemd_csv::table_ir::{extract_tables_ir_with_options, CsvExtractOptions};

let content = b"item,amount\nwidget,42\n";

let doc = extract_tables_ir(content)?;
let markdown = markdown_from_bytes(content)?;

let semicolon_doc = extract_tables_ir_with_options(
    b"item;amount\nwidget;42\n",
    CsvExtractOptions {
        delimiter: b';',
        ..Default::default()
    },
)?;
```

## Tests

```bash
cargo test -p officemd_csv
```
