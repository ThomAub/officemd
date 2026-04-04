use wasm_bindgen::prelude::*;

use officemd_core::format::{detect_format_from_bytes, parse_format, DocumentFormat};

/// Convert a document to markdown, auto-detecting the format from file content.
///
/// Accepts raw file bytes (e.g. from a `Uint8Array` in JavaScript) and returns
/// the extracted markdown string.
///
/// Supported formats: DOCX, XLSX, PPTX, PDF, CSV (CSV requires explicit format).
#[wasm_bindgen]
pub fn convert_to_markdown(content: &[u8]) -> Result<String, JsError> {
    let format =
        detect_format_from_bytes(content).map_err(|e| JsError::new(&e))?;
    convert_with_format(content, format)
}

/// Convert a document to markdown with an explicit format string.
///
/// Format must be one of: "docx", "xlsx", "csv", "pptx", "pdf".
#[wasm_bindgen]
pub fn convert_to_markdown_with_format(
    content: &[u8],
    format: &str,
) -> Result<String, JsError> {
    let fmt = parse_format(format)
        .ok_or_else(|| JsError::new("format must be one of: docx, xlsx, csv, pptx, pdf"))?;
    convert_with_format(content, fmt)
}

/// Convert a document to its JSON IR representation, auto-detecting format.
#[wasm_bindgen]
pub fn convert_to_json(content: &[u8]) -> Result<String, JsError> {
    let format =
        detect_format_from_bytes(content).map_err(|e| JsError::new(&e))?;
    convert_to_json_with_format(content, &format.to_string())
}

/// Convert a document to its JSON IR representation with an explicit format.
#[wasm_bindgen]
pub fn convert_to_json_with_format(
    content: &[u8],
    format: &str,
) -> Result<String, JsError> {
    let fmt = parse_format(format)
        .ok_or_else(|| JsError::new("format must be one of: docx, xlsx, csv, pptx, pdf"))?;
    let ir = extract_ir(content, fmt)?;
    serde_json::to_string(&ir).map_err(|e| JsError::new(&e.to_string()))
}

/// Detect the format of a document from its bytes.
///
/// Returns a string like "docx", "xlsx", "pptx", or "pdf".
#[wasm_bindgen]
pub fn detect_format(content: &[u8]) -> Result<String, JsError> {
    let format =
        detect_format_from_bytes(content).map_err(|e| JsError::new(&e))?;
    Ok(format.to_string())
}

fn convert_with_format(content: &[u8], format: DocumentFormat) -> Result<String, JsError> {
    match format {
        DocumentFormat::Docx => officemd_docx::markdown_from_bytes(content)
            .map_err(|e| JsError::new(&e.to_string())),
        DocumentFormat::Xlsx => officemd_xlsx::markdown_from_bytes(content)
            .map_err(|e| JsError::new(&e.to_string())),
        DocumentFormat::Pptx => officemd_pptx::markdown_from_bytes(content)
            .map_err(|e| JsError::new(&e.to_string())),
        DocumentFormat::Pdf => {
            officemd_pdf::markdown_from_bytes_with_options(content, Default::default())
                .map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Csv => officemd_csv::markdown_from_bytes(content)
            .map_err(|e| JsError::new(&e.to_string())),
    }
}

fn extract_ir(
    content: &[u8],
    format: DocumentFormat,
) -> Result<officemd_core::ir::OoxmlDocument, JsError> {
    match format {
        DocumentFormat::Docx => {
            officemd_docx::extract_ir(content).map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Xlsx => {
            officemd_xlsx::extract_tables_ir(content).map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Pptx => {
            officemd_pptx::extract_ir(content).map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Pdf => {
            officemd_pdf::extract_ir(content).map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Csv => {
            officemd_csv::extract_tables_ir(content).map_err(|e| JsError::new(&e.to_string()))
        }
    }
}
