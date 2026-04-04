use wasm_bindgen::prelude::*;

use officemd_core::format::{detect_format_from_bytes, parse_format, DocumentFormat};

/// One-time setup: install a panic hook so Rust panics show up in the
/// browser console instead of the unhelpful "unreachable" default.
#[wasm_bindgen(start)]
pub fn init_panic_hook() {
    console_error_panic_hook::set_once();
}

/// Convert a document to markdown, auto-detecting the format from file content.
///
/// Returns a two-element JS array: `[format_string, markdown_string]` so the
/// caller gets the detected format without re-parsing the file.
#[wasm_bindgen]
pub fn convert_to_markdown(content: &[u8]) -> Result<Box<[JsValue]>, JsError> {
    let format = detect_format_from_bytes(content).map_err(|e| JsError::new(&e))?;
    let md = convert_with_format(content, format)?;
    Ok(Box::new([JsValue::from_str(&format.to_string()), JsValue::from_str(&md)]))
}

/// Convert a document to markdown with an explicit format string.
///
/// Skips format detection entirely — fastest path when format is known.
#[wasm_bindgen]
pub fn convert_to_markdown_with_format(content: &[u8], format: &str) -> Result<String, JsError> {
    let fmt = parse_format(format)
        .ok_or_else(|| JsError::new("format must be one of: docx, xlsx, csv, pptx, pdf"))?;
    convert_with_format(content, fmt)
}

/// Convert a document to its JSON IR representation, auto-detecting format.
#[wasm_bindgen]
pub fn convert_to_json(content: &[u8]) -> Result<String, JsError> {
    let format = detect_format_from_bytes(content).map_err(|e| JsError::new(&e))?;
    let ir = extract_ir(content, format)?;
    serde_json::to_string(&ir).map_err(|e| JsError::new(&e.to_string()))
}

/// Detect the format of a document from its bytes.
///
/// Returns a string like "docx", "xlsx", "pptx", or "pdf".
#[wasm_bindgen]
pub fn detect_format(content: &[u8]) -> Result<String, JsError> {
    let format = detect_format_from_bytes(content).map_err(|e| JsError::new(&e))?;
    Ok(format.to_string())
}

fn convert_with_format(content: &[u8], format: DocumentFormat) -> Result<String, JsError> {
    match format {
        DocumentFormat::Docx => {
            officemd_docx::markdown_from_bytes(content).map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Xlsx => {
            officemd_xlsx::markdown_from_bytes(content).map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Pptx => {
            officemd_pptx::markdown_from_bytes(content).map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Pdf => {
            officemd_pdf::markdown_from_bytes_with_options(content, Default::default())
                .map_err(|e| JsError::new(&e.to_string()))
        }
        DocumentFormat::Csv => {
            officemd_csv::markdown_from_bytes(content).map_err(|e| JsError::new(&e.to_string()))
        }
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
