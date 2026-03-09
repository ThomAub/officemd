//! Lightweight IR extraction entrypoints for CSV.

use officemd_core::ir::{DocumentKind, OoxmlDocument, Sheet};

use crate::error::CsvError;

/// Extract logical sheet names for CSV (single sheet document).
///
/// # Errors
///
/// Returns `CsvError` if CSV parsing fails.
pub fn extract_sheet_names(_content: &[u8]) -> Result<Vec<String>, CsvError> {
    Ok(vec!["Sheet1".to_string()])
}

/// Extract minimal IR JSON for CSV.
///
/// # Errors
///
/// Returns `CsvError` on parse failure or JSON serialization error.
pub fn extract_ir_json(content: &[u8]) -> Result<String, CsvError> {
    let doc = extract_ir(content)?;
    serde_json::to_string(&doc).map_err(|e| CsvError::Json(e.to_string()))
}

/// Build a minimal `OoxmlDocument` for CSV.
///
/// # Errors
///
/// Returns `CsvError` if CSV parsing fails.
pub fn extract_ir(content: &[u8]) -> Result<OoxmlDocument, CsvError> {
    let sheet_names = extract_sheet_names(content)?;
    let sheets: Vec<Sheet> = sheet_names
        .into_iter()
        .map(|name| Sheet {
            name,
            ..Default::default()
        })
        .collect();

    Ok(OoxmlDocument {
        kind: DocumentKind::Xlsx,
        sheets,
        ..Default::default()
    })
}
