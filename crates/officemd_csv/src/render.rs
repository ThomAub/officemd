//! Glue to render CSV bytes to Markdown via the shared renderer.

use officemd_markdown::RenderOptions;

use crate::error::CsvError;
use crate::table_ir::{CsvExtractOptions, extract_tables_ir_with_options};

/// Render CSV bytes to Markdown.
///
/// # Errors
///
/// Returns `CsvError` if CSV parsing fails.
pub fn markdown_from_bytes(content: &[u8]) -> Result<String, CsvError> {
    markdown_from_bytes_with_options(content, RenderOptions::default())
}

/// Render CSV bytes to Markdown with rendering options.
///
/// # Errors
///
/// Returns `CsvError` if CSV parsing fails.
pub fn markdown_from_bytes_with_options(
    content: &[u8],
    options: RenderOptions,
) -> Result<String, CsvError> {
    let doc = extract_tables_ir_with_options(
        content,
        CsvExtractOptions {
            include_document_properties: options.include_document_properties,
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

    #[test]
    fn markdown_includes_document_properties_when_requested() {
        let bytes = b"name,value\nwidget,42\n";
        let markdown = markdown_from_bytes_with_options(
            bytes,
            RenderOptions {
                include_document_properties: true,
                ..Default::default()
            },
        )
        .expect("render markdown");

        assert!(markdown.contains("properties:"));
        assert!(markdown.contains("source_format=csv"));
    }

    #[test]
    fn markdown_omits_document_properties_by_default() {
        let bytes = b"name,value\nwidget,42\n";
        let markdown = markdown_from_bytes(bytes).expect("render markdown");

        assert!(!markdown.contains("### Document Properties"));
    }
}
