# officemd_docx

DOCX extraction helpers built on `officemd_core`. This crate streams WordprocessingML parts to build the shared IR.

## Features
- Extracts body, headers, footers, footnotes, endnotes into `DocSection` blocks.
- Resolves hyperlinks via `.rels` and a best-effort field-code fallback.
- Parses comments and inserts inline anchors plus per-section footnotes.
- Collects document properties from `docProps/*`.

## Rust Usage

```rust
use officemd_docx::extract_ir;

let doc = extract_ir(docx_bytes)?;
```

## Cargo Example

```bash
cargo run -p officemd_examples --bin extract_ir_docx -- path/to/document.docx
```

## Python usage

```python
from officemd import extract_ir_json, markdown_from_bytes

doc_json = extract_ir_json(docx_bytes)
markdown = markdown_from_bytes(docx_bytes, format="docx")
markdown_with_props = markdown_from_bytes(
    docx_bytes,
    format="docx",
    include_document_properties=True,
)
```

## Tests

```bash
# Rust tests
cargo test -p officemd_docx

# Python bindings (from dedicated adapter crate)
cd crates/officemd_python && uv run maturin develop --release
```

Fixture note:
- `tests/fixtures/basic.docx` is the crate-local sample fixture path (fixture-based tests skip when missing).
- Real-world fixtures can be placed in `tests/data/*.docx` (crate local) or `../../tests/data/*.docx` (repo root).
