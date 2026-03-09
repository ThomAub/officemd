# officemd_core

Shared IR, package reading, and relationship parsing for OfficeMD.

This crate is the internal foundation for the format extractors, markdown renderer, Docling conversion layer, and language bindings.

## Responsibilities

- Shared document IR types
- OOXML package and relationship helpers
- Common building blocks used by DOCX, XLSX, CSV, PPTX, and PDF crates

## Build

```bash
cargo test -p officemd_core
```
