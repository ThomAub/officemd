//! Build a table-centric IR for CSV.

use std::collections::HashMap;
use std::fmt::Write as _;

use officemd_core::ir::{
    DocumentKind, DocumentProperties, FormulaNote, Inline, OoxmlDocument, Paragraph, Sheet, Table,
    TableCell,
};

use crate::error::CsvError;

/// CSV extraction options.
#[derive(Debug, Clone)]
pub struct CsvExtractOptions {
    pub delimiter: u8,
    pub flexible: bool,
    pub include_document_properties: bool,
    pub sheet_name: String,
}

impl Default for CsvExtractOptions {
    fn default() -> Self {
        Self {
            delimiter: b',',
            flexible: true,
            include_document_properties: false,
            sheet_name: "Sheet1".to_string(),
        }
    }
}

use officemd_core::ir::synthetic_col_headers;

/// Extract table-centric IR with default options.
///
/// # Errors
///
/// Returns `CsvError` if CSV parsing fails.
pub fn extract_tables_ir(content: &[u8]) -> Result<OoxmlDocument, CsvError> {
    extract_tables_ir_with_options(content, CsvExtractOptions::default())
}

/// Extract table-centric IR with explicit options.
///
/// # Errors
///
/// Returns `CsvError` if CSV parsing fails.
pub fn extract_tables_ir_with_options(
    content: &[u8],
    options: CsvExtractOptions,
) -> Result<OoxmlDocument, CsvError> {
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .delimiter(options.delimiter)
        .flexible(options.flexible)
        .from_reader(content);

    let mut max_cols = 0usize;
    let mut table_rows: Vec<Vec<TableCell>> = Vec::new();
    let mut formulas = Vec::new();

    for (row_idx, record) in reader.records().enumerate() {
        let record = record?;
        max_cols = max_cols.max(record.len());

        let mut out_row = Vec::with_capacity(record.len().max(1));
        for (col_idx, value) in record.iter().enumerate() {
            let value = value.to_string();
            if let Some(formula) = value.strip_prefix('=') {
                let trimmed = formula.trim();
                if !trimmed.is_empty() {
                    let col_name = col_to_name(col_idx + 1);
                    let mut cell_ref = String::with_capacity(col_name.len() + 20);
                    cell_ref.push_str(&col_name);
                    let _ = write!(cell_ref, "{}", row_idx + 1);
                    formulas.push(FormulaNote {
                        cell_ref,
                        formula: trimmed.to_string(),
                    });
                }
            }

            out_row.push(TableCell {
                content: vec![Paragraph {
                    inlines: vec![Inline::Text(value)],
                }],
            });
        }
        table_rows.push(out_row);
    }

    let cols = max_cols.max(1);
    let headers = synthetic_col_headers(cols);
    for row in &mut table_rows {
        while row.len() < cols {
            row.push(TableCell {
                content: vec![Paragraph {
                    inlines: vec![Inline::Text(String::new())],
                }],
            });
        }
    }

    if table_rows.is_empty() {
        table_rows.push(vec![TableCell {
            content: vec![Paragraph {
                inlines: vec![Inline::Text(String::new())],
            }],
        }]);
    }

    let caption = Some(format!(
        "Table 1 (rows 1–{}, cols A–{})",
        table_rows.len().max(1),
        col_to_name(cols)
    ));

    let table = Table {
        caption,
        headers,
        rows: table_rows,
        synthetic_headers: true,
    };

    let properties = options
        .include_document_properties
        .then(|| build_properties(options.delimiter));

    Ok(OoxmlDocument {
        kind: DocumentKind::Xlsx,
        properties,
        sheets: vec![Sheet {
            name: options.sheet_name,
            tables: vec![table],
            formulas,
            hyperlinks: Vec::new(),
        }],
        ..Default::default()
    })
}

/// Extract table-centric IR JSON with default options.
///
/// # Errors
///
/// Returns `CsvError` on parse failure or JSON serialization error.
pub fn extract_tables_ir_json(content: &[u8]) -> Result<String, CsvError> {
    let doc = extract_tables_ir(content)?;
    serde_json::to_string(&doc).map_err(|e| CsvError::Json(e.to_string()))
}

/// Extract table-centric IR JSON with explicit options.
///
/// # Errors
///
/// Returns `CsvError` on parse failure or JSON serialization error.
pub fn extract_tables_ir_json_with_options(
    content: &[u8],
    delimiter: u8,
    include_document_properties: bool,
) -> Result<String, CsvError> {
    let doc = extract_tables_ir_with_options(
        content,
        CsvExtractOptions {
            delimiter,
            include_document_properties,
            ..Default::default()
        },
    )?;
    serde_json::to_string(&doc).map_err(|e| CsvError::Json(e.to_string()))
}

fn build_properties(delimiter: u8) -> DocumentProperties {
    let mut custom = HashMap::new();
    custom.insert("source_format".to_string(), "csv".to_string());
    custom.insert("delimiter".to_string(), char::from(delimiter).to_string());

    DocumentProperties {
        core: HashMap::new(),
        app: HashMap::new(),
        custom,
    }
}

/// Convert column number (1-based) to Excel column name (A, B, ..., AA).
fn col_to_name(mut n: usize) -> String {
    let mut reversed = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        // Safety: rem is always 0..25 from `% 26`, fits in u8.
        #[allow(clippy::cast_possible_truncation)]
        reversed.push((b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    reversed.into_iter().rev().collect()
}
