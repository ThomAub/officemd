# officemd_pptx

PPTX extraction helpers built on `officemd_core`. The crate emits shared OOXML IR for slides, notes,
comments, tables, and hyperlinks.

## Features
- Slide order via `ppt/presentation.xml` + rels
- Slide text and titles (placeholder-aware)
- DrawingML tables to IR tables with synthetic headers
- Slide notes extraction
- Slide comments with author mapping
- Hyperlink resolution via slide rels

## Build

```bash
cargo test -p officemd_pptx
```

## Rust usage

```rust
use std::collections::HashSet;
use officemd_pptx::{extract_ir_with_options, PptxExtractOptions};

let mut only_one = HashSet::new();
only_one.insert(1);
let doc = extract_ir_with_options(
    pptx_bytes,
    PptxExtractOptions {
        slide_numbers: Some(only_one),
    },
)?;
```

## Python usage

```python
from officemd import extract_ir_json, markdown_from_bytes

with open("deck.pptx", "rb") as f:
    data = f.read()

print(extract_ir_json(data, format="pptx"))
print(markdown_from_bytes(data, format="pptx"))
print(markdown_from_bytes(data, format="pptx", include_document_properties=True))
```

## Cargo example

```bash
cargo run -p officemd_examples --bin extract_ir_pptx -- path/to/deck.pptx
```

## Tests

```bash
cd crates/officemd_python && uv run maturin develop --release
uv run --project crates/officemd_python pytest crates/officemd_pptx/tests -q
```

## Fixture note
- `tests/fixtures/sample.pptx` is the sample fixture path for unit tests (fixture-based tests skip when missing).
- Real-world PPTX files can be placed in `tests/data/*.pptx` (crate local) or `../../tests/data/*.pptx` (repo root).
