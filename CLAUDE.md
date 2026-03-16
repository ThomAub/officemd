# officemd

Monorepo for extracting OOXML/PDF documents to Markdown with a shared IR.

## Crate layout

| Crate | Purpose |
|---|---|
| `officemd_core` | Shared IR types, OPC package reader |
| `officemd_markdown` | Markdown renderer + bidirectional parser |
| `officemd_docx` | DOCX extractor |
| `officemd_xlsx` | XLSX extractor |
| `officemd_csv` | CSV extractor |
| `officemd_pptx` | PPTX extractor |
| `officemd_pdf` | PDF extractor (pdf-inspector vendored) |
| `officemd_docling` | Docling format converter |
| `officemd_cli` | Rust CLI binary |
| `officemd_js` | Node/Bun napi bindings + JS CLI |
| `officemd_python` | Python (PyO3/maturin) bindings |

## Build and test

```sh
cargo test --workspace                    # all Rust tests
cargo clippy --workspace --exclude officemd_pdf  # lint (pdf_inspector is vendored)

# Python
cd crates/officemd_python
uv run maturin develop --release
uv run pytest ../../crates/tests/snapshots/ -q

# JS
cd crates/officemd_js
npm install && npm run build
node --test tests/cli-selection.test.mjs
```

## Releasing

All packages share a single version. Use `bump.py` to bump, commit, and tag:

```sh
uv run bump.py patch            # 0.1.4 -> 0.1.5
uv run bump.py minor            # 0.1.4 -> 0.2.0
uv run bump.py major            # 0.1.4 -> 1.0.0
uv run bump.py patch --dry-run  # preview without writing
```

This updates versions in all 22 files (12 Cargo.toml, 9 package.json, 1 pyproject.toml),
runs `cargo check` to update Cargo.lock, commits, and creates the git tag.

After running, push with:

```sh
git push origin main && git push origin v<new-version>
```

### Version files managed by bump.py

- `Cargo.toml` (workspace)
- `crates/officemd_*/Cargo.toml` (12 crates, including inter-crate dep versions)
- `crates/officemd_js/package.json` + `crates/officemd_js/npm/*/package.json` (9 files)
- `crates/officemd_python/pyproject.toml`

### Snapshot regeneration

After format changes to the markdown renderer, snapshots need updating:

```sh
# Rust (insta)
INSTA_UPDATE=always cargo test -p officemd_snapshot_tests

# Python (pytest-regressions)
cd crates/officemd_python
uv run maturin develop --release
uv run pytest ../../crates/tests/snapshots/ --force-regen -q
```

## Conventions

- `officemd_pdf/src/pdf_inspector/` is vendored code - excluded from clippy
- Pre-commit hooks via prek: `cargo fmt`, `cargo clippy`, `oxlint`, `ruff`
- Snapshot tests live in `crates/tests/rust_snapshots/` (Rust) and `crates/tests/snapshots/` (Python)
- Fixtures in `examples/data/`
