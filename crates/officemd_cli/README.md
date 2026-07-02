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
officemd plan examples/data/showcase.xlsx
officemd stream examples/data/OpenXML_WhitePaper.pdf --pages 1-3
officemd convert examples/data/showcase.pptx --output deck.md
```

## Focus

- Office document extraction for LLM and agent workflows
- Compact markdown and JSON IR from a single CLI
- Agent-oriented parse plans that expose selectors and exact follow-up commands
- Docling-compatible output via the bindings layer
