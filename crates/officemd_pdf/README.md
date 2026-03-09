# officemd_pdf

PDF extraction and diagnostics for OfficeMD.

PDF support is included because it is useful in agent-oriented document pipelines, but the main OfficeMD story remains Office documents first.

## Notes

- PDF text extraction is powered by a vendored copy of [pdf-inspector](https://github.com/firecrawl/pdf-inspector) (MIT, by Firecrawl)
- All dependencies are crates.io-compatible; this crate is publishable
