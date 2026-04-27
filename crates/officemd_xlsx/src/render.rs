//! Glue to render XLSX bytes to Markdown via the shared renderer.

use crate::error::XlsxError;
use crate::table_ir::{
    XlsxExtractOptions, XlsxIncludeOptions, XlsxTrimOptions, extract_tables_ir_with_options,
};
use officemd_markdown::RenderOptions;

/// Render XLSX bytes to Markdown using table IR (single-table per sheet for now).
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn markdown_from_bytes(content: &[u8]) -> Result<String, XlsxError> {
    markdown_from_bytes_with_options(content, RenderOptions::default())
}

/// Render XLSX bytes to Markdown with rendering options.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn markdown_from_bytes_with_options(
    content: &[u8],
    options: RenderOptions,
) -> Result<String, XlsxError> {
    let doc = extract_tables_ir_with_options(
        content,
        &XlsxExtractOptions {
            include: XlsxIncludeOptions {
                document_properties: options.include.document_properties,
            },
            trim: XlsxTrimOptions {
                empty_edges: matches!(
                    options.markdown_profile,
                    officemd_markdown::MarkdownProfile::LlmCompact
                ),
            },
            ..Default::default()
        },
    )?;
    Ok(officemd_markdown::render_document_with_options(
        &doc, options,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    fn build_xlsx(parts: Vec<(&str, &str)>) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut writer = ZipWriter::new(std::io::Cursor::new(&mut buffer));
        let options: FileOptions<'_, ()> = FileOptions::default();

        for (path, contents) in parts {
            writer.start_file(path, options).expect("start file");
            writer
                .write_all(contents.as_bytes())
                .expect("write contents");
        }

        writer.finish().expect("finish zip");
        buffer
    }

    fn minimal_xlsx_with_doc_props() -> Vec<u8> {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1"><c r="A1"><v>1</v></c></row>
    </sheetData>
</worksheet>"#;

        let core = r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Quarterly Results</dc:title>
</cp:coreProperties>"#;

        build_xlsx(vec![
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", sheet),
            ("docProps/core.xml", core),
        ])
    }

    #[test]
    fn markdown_includes_document_properties_when_requested() {
        let bytes = minimal_xlsx_with_doc_props();
        let markdown = markdown_from_bytes_with_options(
            &bytes,
            RenderOptions {
                include: officemd_markdown::RenderIncludeOptions {
                    document_properties: true,
                    ..Default::default()
                },
                markdown_profile: officemd_markdown::MarkdownProfile::Human,
                ..Default::default()
            },
        )
        .expect("render markdown");

        assert!(markdown.contains("### Document Properties"));
        assert!(markdown.contains("Quarterly Results"));
    }

    #[test]
    fn markdown_omits_document_properties_by_default() {
        let bytes = minimal_xlsx_with_doc_props();
        let markdown = markdown_from_bytes(&bytes).expect("render markdown");

        assert!(!markdown.contains("### Document Properties"));
        assert!(!markdown.contains("Quarterly Results"));
    }
}
