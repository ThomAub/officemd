//! Glue to render DOCX bytes to Markdown via the shared renderer.

use crate::error::DocxError;
use crate::extract::extract_ir;
use officemd_markdown::RenderOptions;

/// Render DOCX bytes to Markdown using the shared renderer.
///
/// # Errors
///
/// Returns [`DocxError`] if the DOCX content cannot be extracted.
pub fn markdown_from_bytes(content: &[u8]) -> Result<String, DocxError> {
    markdown_from_bytes_with_options(content, RenderOptions::default())
}

/// Render DOCX bytes to Markdown with rendering options.
///
/// # Errors
///
/// Returns [`DocxError`] if the DOCX content cannot be extracted.
pub fn markdown_from_bytes_with_options(
    content: &[u8],
    options: RenderOptions,
) -> Result<String, DocxError> {
    let doc = extract_ir(content)?;
    Ok(officemd_markdown::render_document_with_options(
        &doc, options,
    ))
}
