# Examples

Shared fixtures and runnable examples for Rust, JavaScript, and Python.

## Generate Fixtures

```bash
uv run --with python-docx --with openpyxl --with python-pptx python examples/generate_data.py
```

Primary fixtures:

- `examples/data/showcase.docx`
- `examples/data/showcase.xlsx`
- `examples/data/showcase.csv`
- `examples/data/showcase.pptx`
- `examples/data/OpenXML_WhitePaper.pdf`

## Rust

```bash
cargo run -p officemd_examples --bin extract_ir_docx -- examples/data/showcase.docx
cargo run -p officemd_examples --bin extract_ir_xlsx -- examples/data/showcase.xlsx
cargo run -p officemd_examples --bin extract_ir_csv -- examples/data/showcase.csv
cargo run -p officemd_examples --bin extract_ir_pptx -- examples/data/showcase.pptx
cargo run -p officemd_examples --bin edit_showcase_words
cargo run -p officemd_examples --bin edit_showcase_words_batch
```

## JavaScript

```bash
cd crates/officemd_js
npm install
npm run build
node ../../examples/js/node-officemd.mjs ../../examples/data/showcase.docx
bun run ../../examples/js/bun-officemd.mjs ../../examples/data/showcase.docx
```

## Python

```bash
cd crates/officemd_python
uv sync --dev
uv run maturin develop --release
uv run python ../../examples/python/extract_ir.py ../../examples/data/showcase.docx
uv run python ../../examples/python/xlsx_streaming_ir.py ../../examples/data/showcase.xlsx --style-aware-values
uv run python ../../examples/python/pdf_inspect_fonts.py ../../examples/data/OpenXML_WhitePaper.pdf --limit 10
```

## Smoke Tests

```bash
cd crates/officemd_python
uv run pytest ../../examples/python/tests/test_showcase.py -q
```
