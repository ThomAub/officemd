# Architecture

OfficeMD is a standalone Rust-first workspace for extracting Office documents into markdown and structured representations for LLM and agent workflows.

## Layers

```text
officemd_core
  Shared IR, package reading, relationships, low-level helpers

officemd_docx / officemd_xlsx / officemd_csv / officemd_pptx / officemd_pdf
  Format-specific extraction into the shared IR

officemd_markdown / officemd_docling
  Output layers: markdown and Docling JSON

officemd_cli / officemd_js / officemd_python
  User-facing surfaces
```

## Data Flow

```text
document bytes
  -> format detection / package reader
  -> format extractor
  -> shared IR
  -> markdown output
     or JSON IR
     or Docling JSON
```

## Public Surfaces

- Rust CLI: `officemd`
- Rust crates: `officemd_*`
- Node/Bun package: `office-md`
- Python package: `officemd`

## Markdown Profiles

- `compact`: token-lean markdown for LLM context windows
- `human`: richer markdown for inspection and review

## PDF Note

PDF support remains part of the workspace because it is useful in document pipelines, but the primary product story is Office document extraction. The PDF crate currently relies on Git-based dependencies; keep that in mind when moving from dry-run packaging to live registry publication.
