//! Native Node/Bun bindings for OfficeMD extraction and rendering.

// N-API `#[napi]` function signatures must take owned values, not references.
#![allow(clippy::needless_pass_by_value)]

use napi::bindgen_prelude::{Buffer, Error, Result, Status};
use napi_derive::napi;
use officemd_core::opc::OpcPackage;
use rayon::ThreadPoolBuilder;
use rayon::prelude::*;

const ERR_INVALID_FORMAT: &str = "ERR_INVALID_FORMAT";
const ERR_INVALID_PACKAGE: &str = "ERR_INVALID_OOXML_PACKAGE";
const ERR_UNSUPPORTED_FORMAT: &str = "ERR_UNSUPPORTED_FORMAT";
const ERR_DOCX_EXTRACT: &str = "ERR_DOCX_EXTRACT";
const ERR_XLSX_EXTRACT: &str = "ERR_XLSX_EXTRACT";
const ERR_CSV_EXTRACT: &str = "ERR_CSV_EXTRACT";
const ERR_PPTX_EXTRACT: &str = "ERR_PPTX_EXTRACT";
const ERR_PDF_EXTRACT: &str = "ERR_PDF_EXTRACT";
const ERR_PDF_INSPECT: &str = "ERR_PDF_INSPECT";
const ERR_PDF_FONTS_INSPECT: &str = "ERR_PDF_FONTS_INSPECT";
const ERR_JSON_SERIALIZE: &str = "ERR_JSON_SERIALIZE";
const ERR_DOCLING_CONVERT: &str = "ERR_DOCLING_CONVERT";
const ERR_PARALLELISM: &str = "ERR_PARALLELISM";

#[derive(Debug)]
enum DocumentFormat {
    Docx,
    Xlsx,
    Csv,
    Pptx,
    Pdf,
}

impl DocumentFormat {
    fn extension(&self) -> &'static str {
        match self {
            Self::Docx => ".docx",
            Self::Xlsx => ".xlsx",
            Self::Csv => ".csv",
            Self::Pptx => ".pptx",
            Self::Pdf => ".pdf",
        }
    }
}

fn invalid_arg_error(code: &str, message: impl AsRef<str>) -> Error {
    Error::new(Status::InvalidArg, format!("{code}: {}", message.as_ref()))
}

fn internal_error(code: &str, err: impl std::fmt::Display) -> Error {
    Error::new(Status::GenericFailure, format!("{code}: {err}"))
}

fn resolve_worker_count(workers: Option<u32>) -> usize {
    workers
        .and_then(|v| usize::try_from(v).ok())
        .filter(|v| *v > 0)
        .unwrap_or_else(|| std::thread::available_parallelism().map_or(1, usize::from))
}

fn parse_format(value: &str) -> Option<DocumentFormat> {
    match value.trim().to_ascii_lowercase().as_str() {
        ".docx" | "docx" => Some(DocumentFormat::Docx),
        ".xlsx" | "xlsx" => Some(DocumentFormat::Xlsx),
        ".csv" | "csv" => Some(DocumentFormat::Csv),
        ".pptx" | "pptx" => Some(DocumentFormat::Pptx),
        ".pdf" | "pdf" => Some(DocumentFormat::Pdf),
        _ => None,
    }
}

fn detect_format_from_bytes(content: &[u8]) -> Result<DocumentFormat> {
    if officemd_pdf::looks_like_pdf_header(content) {
        return Ok(DocumentFormat::Pdf);
    }

    let mut package =
        OpcPackage::from_bytes(content).map_err(|err| internal_error(ERR_INVALID_PACKAGE, err))?;

    if package.has_part("word/document.xml") {
        return Ok(DocumentFormat::Docx);
    }
    if package.has_part("xl/workbook.xml") {
        return Ok(DocumentFormat::Xlsx);
    }
    if package.has_part("ppt/presentation.xml") {
        return Ok(DocumentFormat::Pptx);
    }

    Err(internal_error(
        ERR_UNSUPPORTED_FORMAT,
        "Could not detect format from file content (supported: .docx, .xlsx, .csv, .pptx, .pdf; csv requires explicit format)",
    ))
}

fn resolve_format(content: &[u8], format: Option<&str>) -> Result<DocumentFormat> {
    match format {
        Some(value) => parse_format(value).ok_or_else(|| {
            invalid_arg_error(
                ERR_INVALID_FORMAT,
                "format must be one of: .docx, .xlsx, .csv, .pptx, .pdf",
            )
        }),
        None => detect_format_from_bytes(content),
    }
}

fn parse_markdown_style(style: Option<&str>) -> Result<officemd_markdown::MarkdownProfile> {
    match style
        .unwrap_or("compact")
        .trim()
        .to_ascii_lowercase()
        .as_str()
    {
        "compact" | "llm_compact" | "llm-compact" => {
            Ok(officemd_markdown::MarkdownProfile::LlmCompact)
        }
        "human" => Ok(officemd_markdown::MarkdownProfile::Human),
        _ => Err(invalid_arg_error(
            ERR_INVALID_FORMAT,
            "markdown_style must be one of: compact, human",
        )),
    }
}

fn detect_format_impl(content: &[u8]) -> Result<String> {
    let format = detect_format_from_bytes(content)?;
    Ok(format.extension().to_string())
}

fn extract_ir_json_impl(content: &[u8], format: Option<&str>) -> Result<String> {
    let format = resolve_format(content, format)?;
    match format {
        DocumentFormat::Docx => officemd_docx::extract_ir_json(content)
            .map_err(|err| internal_error(ERR_DOCX_EXTRACT, err)),
        DocumentFormat::Xlsx => {
            let doc = officemd_xlsx::extract_ir::extract_ir(content)
                .map_err(|err| internal_error(ERR_XLSX_EXTRACT, err))?;
            serde_json::to_string(&doc).map_err(|err| internal_error(ERR_JSON_SERIALIZE, err))
        }
        DocumentFormat::Csv => officemd_csv::extract_ir::extract_ir_json(content)
            .map_err(|err| internal_error(ERR_CSV_EXTRACT, err)),
        DocumentFormat::Pptx => officemd_pptx::extract_ir_json(content)
            .map_err(|err| internal_error(ERR_PPTX_EXTRACT, err)),
        DocumentFormat::Pdf => officemd_pdf::extract_ir_json(content)
            .map_err(|err| internal_error(ERR_PDF_EXTRACT, err)),
    }
}

fn markdown_from_bytes_impl(
    content: &[u8],
    format: Option<&str>,
    options: officemd_markdown::RenderOptions,
    force_extract: bool,
) -> Result<String> {
    let format = resolve_format(content, format)?;
    match format {
        DocumentFormat::Docx => {
            let doc = officemd_docx::extract_ir(content)
                .map_err(|err| internal_error(ERR_DOCX_EXTRACT, err))?;
            Ok(officemd_markdown::render_document_with_options(
                &doc, options,
            ))
        }
        DocumentFormat::Xlsx => officemd_xlsx::markdown_from_bytes_with_options(content, options)
            .map_err(|err| internal_error(ERR_XLSX_EXTRACT, err)),
        DocumentFormat::Csv => officemd_csv::markdown_from_bytes_with_options(content, options)
            .map_err(|err| internal_error(ERR_CSV_EXTRACT, err)),
        DocumentFormat::Pptx => {
            let doc = officemd_pptx::extract_ir(content)
                .map_err(|err| internal_error(ERR_PPTX_EXTRACT, err))?;
            Ok(officemd_markdown::render_document_with_options(
                &doc, options,
            ))
        }
        DocumentFormat::Pdf => {
            officemd_pdf::markdown_from_bytes_force(content, options, force_extract)
                .map_err(|err| internal_error(ERR_PDF_EXTRACT, err))
        }
    }
}

fn markdown_from_bytes_batch_impl(
    contents: Vec<Vec<u8>>,
    format: Option<String>,
    options: officemd_markdown::RenderOptions,
    workers: Option<u32>,
) -> Result<Vec<String>> {
    if contents.is_empty() {
        return Ok(Vec::new());
    }

    let worker_count = resolve_worker_count(workers);
    if worker_count <= 1 || contents.len() <= 1 {
        return contents
            .into_iter()
            .map(|content| markdown_from_bytes_impl(&content, format.as_deref(), options, false))
            .collect();
    }

    let pool = ThreadPoolBuilder::new()
        .num_threads(worker_count)
        .build()
        .map_err(|err| internal_error(ERR_PARALLELISM, err))?;

    pool.install(|| {
        contents
            .into_par_iter()
            .map(|content| markdown_from_bytes_impl(&content, format.as_deref(), options, false))
            .collect()
    })
}

fn extract_sheet_names_impl(content: &[u8]) -> Result<Vec<String>> {
    officemd_xlsx::extract_ir::extract_sheet_names(content)
        .map_err(|err| internal_error(ERR_XLSX_EXTRACT, err))
}

fn extract_tables_ir_json_impl(
    content: &[u8],
    style_aware_values: bool,
    streaming_rows: bool,
    include_document_properties: bool,
) -> Result<String> {
    officemd_xlsx::table_ir::extract_tables_ir_json_with_options(
        content,
        style_aware_values,
        streaming_rows,
        include_document_properties,
    )
    .map_err(|err| internal_error(ERR_XLSX_EXTRACT, err))
}

fn extract_csv_tables_ir_json_impl(
    content: &[u8],
    delimiter: Option<String>,
    include_document_properties: Option<bool>,
) -> Result<String> {
    let delimiter = delimiter.unwrap_or_else(|| ",".to_string());
    if delimiter.len() != 1 {
        return Err(invalid_arg_error(
            ERR_INVALID_FORMAT,
            "delimiter must be exactly one character",
        ));
    }

    let delimiter_byte = delimiter.as_bytes()[0];
    officemd_csv::table_ir::extract_tables_ir_json_with_options(
        content,
        delimiter_byte,
        include_document_properties.unwrap_or(false),
    )
    .map_err(|err| internal_error(ERR_CSV_EXTRACT, err))
}

fn inspect_pdf_json_impl(content: &[u8]) -> Result<String> {
    let diagnostics =
        officemd_pdf::inspect_pdf(content).map_err(|err| internal_error(ERR_PDF_INSPECT, err))?;
    serde_json::to_string(&diagnostics).map_err(|err| internal_error(ERR_JSON_SERIALIZE, err))
}

fn inspect_pdf_fonts_json_impl(content: &[u8]) -> Result<String> {
    officemd_pdf::inspect_pdf_fonts_json(content)
        .map_err(|err| internal_error(ERR_PDF_FONTS_INSPECT, err))
}

fn docling_from_bytes_impl(content: &[u8], format: Option<&str>) -> Result<String> {
    let format = resolve_format(content, format)?;
    let doc = match format {
        DocumentFormat::Docx => officemd_docx::extract_ir(content)
            .map_err(|err| internal_error(ERR_DOCX_EXTRACT, err))?,
        DocumentFormat::Xlsx => officemd_xlsx::extract_ir::extract_ir(content)
            .map_err(|err| internal_error(ERR_XLSX_EXTRACT, err))?,
        DocumentFormat::Csv => officemd_csv::extract_ir::extract_ir(content)
            .map_err(|err| internal_error(ERR_CSV_EXTRACT, err))?,
        DocumentFormat::Pptx => officemd_pptx::extract_ir(content)
            .map_err(|err| internal_error(ERR_PPTX_EXTRACT, err))?,
        DocumentFormat::Pdf => {
            officemd_pdf::extract_ir(content).map_err(|err| internal_error(ERR_PDF_EXTRACT, err))?
        }
    };
    officemd_docling::convert_document_json(&doc)
        .map_err(|err| internal_error(ERR_DOCLING_CONVERT, err))
}

/// Detect the document format from raw bytes.
///
/// # Errors
///
/// Returns an error if the format cannot be determined from the content.
#[napi]
pub fn detect_format(content: Buffer) -> Result<String> {
    detect_format_impl(content.as_ref())
}

/// Extract the intermediate representation as JSON.
///
/// # Errors
///
/// Returns an error if format detection or extraction fails.
#[napi]
pub fn extract_ir_json(content: Buffer, format: Option<String>) -> Result<String> {
    extract_ir_json_impl(content.as_ref(), format.as_deref())
}

/// Render document bytes as Markdown.
///
/// # Errors
///
/// Returns an error if format detection, extraction, or rendering fails.
#[napi]
#[allow(clippy::too_many_arguments)]
pub fn markdown_from_bytes(
    content: Buffer,
    format: Option<String>,
    include_document_properties: Option<bool>,
    use_first_row_as_header: Option<bool>,
    include_headers_footers: Option<bool>,
    include_formulas: Option<bool>,
    markdown_style: Option<String>,
    force_extract: Option<bool>,
) -> Result<String> {
    let markdown_profile = parse_markdown_style(markdown_style.as_deref())?;
    let options = officemd_markdown::RenderOptions {
        include_document_properties: include_document_properties.unwrap_or(false),
        use_first_row_as_header: use_first_row_as_header.unwrap_or(true),
        include_headers_footers: include_headers_footers.unwrap_or(true),
        include_formulas: include_formulas.unwrap_or(true),
        markdown_profile,
    };
    markdown_from_bytes_impl(
        content.as_ref(),
        format.as_deref(),
        options,
        force_extract.unwrap_or(false),
    )
}

/// Render multiple documents as Markdown with Rust-side parallel workers.
///
/// # Errors
///
/// Returns an error if any item fails format detection, extraction, or rendering.
#[napi]
#[allow(clippy::too_many_arguments)]
pub fn markdown_from_bytes_batch(
    contents: Vec<Buffer>,
    format: Option<String>,
    workers: Option<u32>,
    include_document_properties: Option<bool>,
    use_first_row_as_header: Option<bool>,
    include_headers_footers: Option<bool>,
    include_formulas: Option<bool>,
    markdown_style: Option<String>,
) -> Result<Vec<String>> {
    let markdown_profile = parse_markdown_style(markdown_style.as_deref())?;
    let options = officemd_markdown::RenderOptions {
        include_document_properties: include_document_properties.unwrap_or(false),
        use_first_row_as_header: use_first_row_as_header.unwrap_or(true),
        include_headers_footers: include_headers_footers.unwrap_or(true),
        include_formulas: include_formulas.unwrap_or(true),
        markdown_profile,
    };
    let payloads = contents.into_iter().map(|buf| buf.to_vec()).collect();
    markdown_from_bytes_batch_impl(payloads, format, options, workers)
}

/// Extract sheet names from an XLSX workbook.
///
/// # Errors
///
/// Returns an error if the content is not a valid XLSX workbook.
#[napi]
pub fn extract_sheet_names(content: Buffer) -> Result<Vec<String>> {
    extract_sheet_names_impl(content.as_ref())
}

/// Extract XLSX table data as a JSON string.
///
/// # Errors
///
/// Returns an error if XLSX extraction or JSON serialization fails.
#[napi]
pub fn extract_tables_ir_json(
    content: Buffer,
    style_aware_values: Option<bool>,
    streaming_rows: Option<bool>,
    include_document_properties: Option<bool>,
) -> Result<String> {
    extract_tables_ir_json_impl(
        content.as_ref(),
        style_aware_values.unwrap_or(false),
        streaming_rows.unwrap_or(false),
        include_document_properties.unwrap_or(false),
    )
}

/// Extract CSV table data as a JSON string.
///
/// # Errors
///
/// Returns an error if the delimiter is invalid or CSV extraction fails.
#[napi]
pub fn extract_csv_tables_ir_json(
    content: Buffer,
    delimiter: Option<String>,
    include_document_properties: Option<bool>,
) -> Result<String> {
    extract_csv_tables_ir_json_impl(content.as_ref(), delimiter, include_document_properties)
}

/// Inspect a PDF and return diagnostics as JSON.
///
/// # Errors
///
/// Returns an error if the content is not a valid PDF or inspection fails.
#[napi]
pub fn inspect_pdf_json(content: Buffer) -> Result<String> {
    inspect_pdf_json_impl(content.as_ref())
}

/// Inspect PDF font information and return as JSON.
///
/// # Errors
///
/// Returns an error if the content is not a valid PDF or font inspection fails.
#[napi]
pub fn inspect_pdf_fonts_json(content: Buffer) -> Result<String> {
    inspect_pdf_fonts_json_impl(content.as_ref())
}

/// Convert document bytes to Docling JSON format.
///
/// # Errors
///
/// Returns an error if format detection, extraction, or Docling conversion fails.
#[napi]
pub fn docling_from_bytes(content: Buffer, format: Option<String>) -> Result<String> {
    docling_from_bytes_impl(content.as_ref(), format.as_deref())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    fn build_zip(parts: Vec<(&str, &str)>) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut writer = ZipWriter::new(std::io::Cursor::new(&mut buffer));
        let options: FileOptions<'_, ()> = FileOptions::default();

        for (path, contents) in parts {
            writer.start_file(path, options).unwrap();
            writer.write_all(contents.as_bytes()).unwrap();
        }

        writer.finish().unwrap();
        buffer
    }

    fn minimal_xlsx() -> Vec<u8> {
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
    <row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>
  </sheetData>
</worksheet>"#;
        build_zip(vec![
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", sheet),
        ])
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
        build_zip(vec![
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", sheet),
            ("docProps/core.xml", core),
        ])
    }

    #[test]
    fn detects_docx_format() {
        let bytes = build_zip(vec![("word/document.xml", "<w:document/>")]);
        let detected = detect_format_impl(&bytes).unwrap();
        assert_eq!(detected, ".docx");
    }

    #[test]
    fn detects_xlsx_format() {
        let bytes = build_zip(vec![("xl/workbook.xml", "<workbook/>")]);
        let detected = detect_format_impl(&bytes).unwrap();
        assert_eq!(detected, ".xlsx");
    }

    #[test]
    fn detects_pptx_format() {
        let bytes = build_zip(vec![("ppt/presentation.xml", "<p:presentation/>")]);
        let detected = detect_format_impl(&bytes).unwrap();
        assert_eq!(detected, ".pptx");
    }

    #[test]
    fn detects_pdf_format() {
        let detected = detect_format_impl(b"%PDF-1.7\n").unwrap();
        assert_eq!(detected, ".pdf");
    }

    #[test]
    fn rejects_invalid_explicit_format() {
        let err = resolve_format(b"test", Some(".txt")).unwrap_err();
        let message = err.to_string();
        assert!(message.contains(ERR_INVALID_FORMAT));
    }

    #[test]
    fn reports_invalid_package_error_code() {
        let err = detect_format_impl(b"not a zip").unwrap_err();
        let message = err.to_string();
        assert!(message.contains(ERR_INVALID_PACKAGE));
    }

    #[test]
    fn parse_format_accepts_case_insensitive_values() {
        assert!(matches!(parse_format("DOCX"), Some(DocumentFormat::Docx)));
        assert!(matches!(parse_format(".xlsx"), Some(DocumentFormat::Xlsx)));
        assert!(matches!(parse_format("csv"), Some(DocumentFormat::Csv)));
        assert!(matches!(parse_format("Pptx"), Some(DocumentFormat::Pptx)));
        assert!(matches!(parse_format("pdf"), Some(DocumentFormat::Pdf)));
    }

    #[test]
    fn markdown_from_csv_with_explicit_format() {
        let bytes = b"name,value\nwidget,42\n";
        let markdown = markdown_from_bytes_impl(
            bytes,
            Some(".csv"),
            officemd_markdown::RenderOptions::default(),
            false,
        )
        .expect("markdown");
        assert!(markdown.contains("## Sheet: Sheet1"));
        assert!(markdown.contains("| name | value |"));
    }

    #[test]
    fn markdown_batch_runs_for_multiple_csv_documents() {
        let docs = vec![b"name,value\na,1\n".to_vec(), b"name,value\nb,2\n".to_vec()];
        let out = markdown_from_bytes_batch_impl(
            docs,
            Some(".csv".to_string()),
            officemd_markdown::RenderOptions::default(),
            Some(2),
        )
        .expect("batch markdown");
        assert_eq!(out.len(), 2);
        assert!(out[0].contains("| name | value |"));
        assert!(out[1].contains("| name | value |"));
    }

    #[test]
    fn extracts_sheet_names_from_minimal_xlsx() {
        let bytes = minimal_xlsx();
        let sheet_names = extract_sheet_names_impl(&bytes).unwrap();
        assert_eq!(sheet_names, vec!["Sheet1".to_string()]);
    }

    #[test]
    fn extracts_tables_ir_json_from_minimal_xlsx() {
        let bytes = minimal_xlsx();
        let payload = extract_tables_ir_json_impl(&bytes, false, false, false).unwrap();
        let value: serde_json::Value = serde_json::from_str(&payload).unwrap();
        assert_eq!(value["kind"], "Xlsx");
        assert_eq!(value["sheets"][0]["name"], "Sheet1");
        assert!(value["properties"].is_null());
    }

    #[test]
    fn extracts_tables_ir_json_with_document_properties_when_requested() {
        let bytes = minimal_xlsx_with_doc_props();
        let payload = extract_tables_ir_json_impl(&bytes, false, false, true).unwrap();
        let value: serde_json::Value = serde_json::from_str(&payload).unwrap();
        assert!(value["properties"].is_object());
        let core = value["properties"]["core"].as_object().expect("core props");
        assert!(!core.is_empty());
    }

    #[test]
    fn markdown_from_xlsx_includes_document_properties_when_requested() {
        let bytes = minimal_xlsx_with_doc_props();
        let markdown = markdown_from_bytes_impl(
            &bytes,
            Some(".xlsx"),
            officemd_markdown::RenderOptions {
                include_document_properties: true,
                markdown_profile: officemd_markdown::MarkdownProfile::Human,
                ..Default::default()
            },
            false,
        )
        .expect("markdown");
        assert!(markdown.contains("### Document Properties"));
        assert!(markdown.contains("Quarterly Results"));
    }

    #[test]
    fn markdown_from_xlsx_omits_document_properties_by_default() {
        let bytes = minimal_xlsx_with_doc_props();
        let markdown = markdown_from_bytes_impl(
            &bytes,
            Some(".xlsx"),
            officemd_markdown::RenderOptions::default(),
            false,
        )
        .expect("markdown");
        assert!(!markdown.contains("### Document Properties"));
        assert!(!markdown.contains("Quarterly Results"));
    }

    #[test]
    fn inspect_pdf_json_rejects_non_pdf() {
        let err = inspect_pdf_json_impl(b"<html>not pdf</html>").unwrap_err();
        let message = err.to_string();
        assert!(message.contains(ERR_PDF_INSPECT));
    }

    #[test]
    fn inspect_pdf_fonts_json_rejects_non_pdf() {
        let err = inspect_pdf_fonts_json_impl(b"<html>not pdf</html>").unwrap_err();
        let message = err.to_string();
        assert!(message.contains(ERR_PDF_FONTS_INSPECT));
    }
}
