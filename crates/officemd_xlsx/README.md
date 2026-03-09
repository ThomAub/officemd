# officemd_xlsx

XLSX extraction helpers built on `officemd_core` with optional Markdown rendering.

## Features
- Sheet name extraction and minimal IR JSON.
- Table-focused IR extraction (in-house XML reader).
- Markdown rendering from table IR.
- Document properties are opt-in (`include_document_properties=false` by default).

## Rust Usage

```rust
use officemd_xlsx::extract_ir::extract_sheet_names;
use officemd_xlsx::{inspect_sheet_summaries, SheetFilter, XlsxExtractOptions, extract_tables_ir, extract_tables_ir_with_options};

let names = extract_sheet_names(xlsx_bytes)?;
let doc = extract_tables_ir(xlsx_bytes)?;
let summary = inspect_sheet_summaries(xlsx_bytes, None)?;

let mut only_data = SheetFilter::default();
only_data.names.insert("Data".to_string());
let filtered = extract_tables_ir_with_options(
    xlsx_bytes,
    XlsxExtractOptions {
        style_aware_values: false,
        streaming_rows: false,
        sheet_filter: Some(only_data),
        include_document_properties: false,
        trim_empty: false,
    },
)?;
```

## Cargo Example

```bash
cargo run -p officemd_examples --bin extract_ir_xlsx -- path/to/file.xlsx
```

## Python usage

```python
from officemd import (
    extract_sheet_names,
    extract_tables_ir_json,
    markdown_from_bytes,
)

with open("file.xlsx", "rb") as f:
    data = f.read()

print(extract_sheet_names(data))
print(extract_tables_ir_json(data, streaming_rows=True))
print(extract_tables_ir_json(data, include_document_properties=True))
print(markdown_from_bytes(data, format="xlsx"))
print(markdown_from_bytes(data, format="xlsx", include_document_properties=True))
```

Streaming example script:

```bash
uv run --project crates/officemd_python python examples/python/xlsx_streaming_ir.py path/to/file.xlsx
```

## Tests

```bash
# Rust tests
cargo test -p officemd_xlsx

# Python bindings (from dedicated adapter crate)
cd crates/officemd_python && uv run maturin develop --release
```

## Fixture note
- `tests/fixtures/sample.xlsx` is the sample fixture path for unit tests (fixture-based tests skip when missing).
- Real-world XLSX files can be placed in `tests/data/*.xlsx` (crate local) or `../../tests/data/*.xlsx` (repo root).
