//! Glue to render PPTX bytes to Markdown via the shared renderer.

use crate::error::PptxError;
use crate::extract::extract_ir;
use officemd_markdown::RenderOptions;

/// Render PPTX bytes to Markdown using the shared renderer.
///
/// # Errors
///
/// Returns `PptxError` if the PPTX content cannot be parsed.
pub fn markdown_from_bytes(content: &[u8]) -> Result<String, PptxError> {
    markdown_from_bytes_with_options(content, RenderOptions::default())
}

/// Render PPTX bytes to Markdown with rendering options.
///
/// # Errors
///
/// Returns `PptxError` if the PPTX content cannot be parsed.
pub fn markdown_from_bytes_with_options(
    content: &[u8],
    options: RenderOptions,
) -> Result<String, PptxError> {
    let doc = extract_ir(content)?;
    Ok(officemd_markdown::render_document_with_options(
        &doc, options,
    ))
}
