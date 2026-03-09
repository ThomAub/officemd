# officemd_cli

`officemd` is the Rust CLI for OfficeMD.

## Install

```bash
cargo install --path crates/officemd_cli
```

## Examples

```bash
officemd stream examples/data/showcase.docx
officemd stream examples/data/showcase.xlsx --markdown-style compact
officemd inspect examples/data/OpenXML_WhitePaper.pdf --output-format json --pretty
officemd convert examples/data/showcase.pptx --output deck.md
```

## Focus

- Office document extraction for LLM and agent workflows
- Compact markdown and JSON IR from a single CLI
- Docling-compatible output via the bindings layer
