//! Native Python bindings for OfficeMD extraction and rendering.

// PyO3 `#[pyfunction]` signatures must take owned values, not references.
#![allow(clippy::needless_pass_by_value, clippy::too_many_arguments)]

use officemd_core::opc::OpcPackage;
use officemd_markdown::{MarkdownProfile, RenderOptions};
use pyo3::prelude::*;
use rayon::ThreadPoolBuilder;
use rayon::prelude::*;

#[derive(Debug, Clone, Copy)]
enum DocumentFormat {
    Docx,
    Xlsx,
    Csv,
    Pptx,
    Pdf,
}

impl DocumentFormat {
    fn extension(self) -> &'static str {
        match self {
            Self::Docx => ".docx",
            Self::Xlsx => ".xlsx",
            Self::Csv => ".csv",
            Self::Pptx => ".pptx",
            Self::Pdf => ".pdf",
        }
    }
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

fn detect_format_from_bytes(content: &[u8]) -> Result<DocumentFormat, String> {
    if officemd_pdf::looks_like_pdf_header(content) {
        return Ok(DocumentFormat::Pdf);
    }

    let mut package = OpcPackage::from_bytes(content).map_err(|e| e.to_string())?;

    if package.has_part("word/document.xml") {
        return Ok(DocumentFormat::Docx);
    }
    if package.has_part("xl/workbook.xml") {
        return Ok(DocumentFormat::Xlsx);
    }
    if package.has_part("ppt/presentation.xml") {
        return Ok(DocumentFormat::Pptx);
    }

    Err(
        "Could not detect format from file content (supported: .docx, .xlsx, .csv, .pptx, .pdf; csv requires explicit format)"
            .to_string(),
    )
}

fn resolve_format(content: &[u8], format: Option<&str>) -> Result<DocumentFormat, String> {
    match format {
        Some(v) => parse_format(v)
            .ok_or_else(|| "format must be one of: .docx, .xlsx, .csv, .pptx, .pdf".to_string()),
        None => detect_format_from_bytes(content),
    }
}

fn parse_markdown_style(style: Option<&str>) -> Result<MarkdownProfile, String> {
    match style
        .unwrap_or("compact")
        .trim()
        .to_ascii_lowercase()
        .as_str()
    {
        "compact" | "llm_compact" | "llm-compact" => Ok(MarkdownProfile::LlmCompact),
        "human" => Ok(MarkdownProfile::Human),
        _ => Err("markdown_style must be one of: compact, human".to_string()),
    }
}

fn to_py_err(err: impl std::fmt::Display) -> PyErr {
    pyo3::exceptions::PyValueError::new_err(err.to_string())
}

fn resolve_worker_count(workers: Option<usize>) -> usize {
    workers
        .filter(|v| *v > 0)
        .unwrap_or_else(|| std::thread::available_parallelism().map_or(1, usize::from))
}

fn markdown_from_owned_bytes_impl(
    content: &[u8],
    format: Option<&str>,
    options: RenderOptions,
    force_extract: bool,
) -> Result<String, String> {
    let format = resolve_format(content, format)?;
    match format {
        DocumentFormat::Docx => officemd_docx::markdown_from_bytes_with_options(content, options)
            .map_err(|e| e.to_string()),
        DocumentFormat::Xlsx => officemd_xlsx::markdown_from_bytes_with_options(content, options)
            .map_err(|e| e.to_string()),
        DocumentFormat::Csv => officemd_csv::markdown_from_bytes_with_options(content, options)
            .map_err(|e| e.to_string()),
        DocumentFormat::Pptx => officemd_pptx::markdown_from_bytes_with_options(content, options)
            .map_err(|e| e.to_string()),
        DocumentFormat::Pdf => {
            officemd_pdf::markdown_from_bytes_force(content, options, force_extract)
                .map_err(|e| e.to_string())
        }
    }
}

#[pyfunction(signature = (content, format=None))]
fn extract_ir_json(py: Python<'_>, content: &[u8], format: Option<String>) -> PyResult<String> {
    let owned_content = content.to_vec();
    let format = resolve_format(&owned_content, format.as_deref()).map_err(to_py_err)?;
    py.detach(move || match format {
        DocumentFormat::Docx => {
            officemd_docx::extract_ir_json(&owned_content).map_err(|e| e.to_string())
        }
        DocumentFormat::Xlsx => {
            officemd_xlsx::extract_ir::extract_ir_json(&owned_content).map_err(|e| e.to_string())
        }
        DocumentFormat::Csv => {
            officemd_csv::extract_ir::extract_ir_json(&owned_content).map_err(|e| e.to_string())
        }
        DocumentFormat::Pptx => {
            officemd_pptx::extract_ir_json(&owned_content).map_err(|e| e.to_string())
        }
        DocumentFormat::Pdf => {
            officemd_pdf::extract_ir_json(&owned_content).map_err(|e| e.to_string())
        }
    })
    .map_err(to_py_err)
}

#[pyfunction(signature = (
    content,
    format=None,
    include_document_properties=false,
    use_first_row_as_header=true,
    include_headers_footers=true,
    include_formulas=true,
    markdown_style=None,
    force_extract=false,
))]
fn markdown_from_bytes(
    py: Python<'_>,
    content: &[u8],
    format: Option<String>,
    include_document_properties: bool,
    use_first_row_as_header: bool,
    include_headers_footers: bool,
    include_formulas: bool,
    markdown_style: Option<String>,
    force_extract: bool,
) -> PyResult<String> {
    let owned_content = content.to_vec();
    let markdown_profile = parse_markdown_style(markdown_style.as_deref()).map_err(to_py_err)?;
    let options = RenderOptions {
        include_document_properties,
        use_first_row_as_header,
        include_headers_footers,
        include_formulas,
        markdown_profile,
    };
    py.detach(move || {
        markdown_from_owned_bytes_impl(&owned_content, format.as_deref(), options, force_extract)
    })
    .map_err(to_py_err)
}

#[pyfunction(signature = (
    contents,
    format=None,
    workers=None,
    include_document_properties=false,
    use_first_row_as_header=true,
    include_headers_footers=true,
    include_formulas=true,
    markdown_style=None,
))]
fn markdown_from_bytes_batch(
    py: Python<'_>,
    contents: Vec<Vec<u8>>,
    format: Option<String>,
    workers: Option<usize>,
    include_document_properties: bool,
    use_first_row_as_header: bool,
    include_headers_footers: bool,
    include_formulas: bool,
    markdown_style: Option<String>,
) -> PyResult<Vec<String>> {
    let markdown_profile = parse_markdown_style(markdown_style.as_deref()).map_err(to_py_err)?;
    let options = RenderOptions {
        include_document_properties,
        use_first_row_as_header,
        include_headers_footers,
        include_formulas,
        markdown_profile,
    };
    let worker_count = resolve_worker_count(workers);
    py.detach(move || {
        if worker_count <= 1 || contents.len() <= 1 {
            return contents
                .into_iter()
                .map(|content| {
                    markdown_from_owned_bytes_impl(&content, format.as_deref(), options, false)
                })
                .collect::<Result<Vec<_>, _>>();
        }

        let pool = ThreadPoolBuilder::new()
            .num_threads(worker_count)
            .build()
            .map_err(|e| e.to_string())?;
        pool.install(|| {
            contents
                .into_par_iter()
                .map(|content| {
                    markdown_from_owned_bytes_impl(&content, format.as_deref(), options, false)
                })
                .collect::<Result<Vec<_>, _>>()
        })
    })
    .map_err(to_py_err)
}

#[pyfunction]
fn detect_format(py: Python<'_>, content: &[u8]) -> PyResult<String> {
    let owned_content = content.to_vec();
    let format = py
        .detach(move || detect_format_from_bytes(&owned_content))
        .map_err(to_py_err)?;
    Ok(format.extension().to_string())
}

#[pyfunction]
fn inspect_pdf_json(py: Python<'_>, content: &[u8]) -> PyResult<String> {
    let owned_content = content.to_vec();
    py.detach(move || {
        let diagnostics = officemd_pdf::inspect_pdf(&owned_content).map_err(|e| e.to_string())?;
        serde_json::to_string(&diagnostics).map_err(|e| e.to_string())
    })
    .map_err(to_py_err)
}

#[pyfunction]
fn inspect_pdf_fonts_json(py: Python<'_>, content: &[u8]) -> PyResult<String> {
    let owned_content = content.to_vec();
    py.detach(move || {
        officemd_pdf::inspect_pdf_fonts_json(&owned_content).map_err(|e| e.to_string())
    })
    .map_err(to_py_err)
}

#[pyfunction]
fn extract_sheet_names(py: Python<'_>, content: &[u8]) -> PyResult<Vec<String>> {
    let owned_content = content.to_vec();
    py.detach(move || {
        officemd_xlsx::extract_ir::extract_sheet_names(&owned_content).map_err(|e| e.to_string())
    })
    .map_err(to_py_err)
}

#[pyfunction(signature = (content, format=None))]
fn docling_from_bytes(py: Python<'_>, content: &[u8], format: Option<String>) -> PyResult<String> {
    let owned_content = content.to_vec();
    let format = resolve_format(&owned_content, format.as_deref()).map_err(to_py_err)?;
    py.detach(move || {
        let doc =
            match format {
                DocumentFormat::Docx => {
                    officemd_docx::extract_ir(&owned_content).map_err(|e| e.to_string())?
                }
                DocumentFormat::Xlsx => officemd_xlsx::extract_ir::extract_ir(&owned_content)
                    .map_err(|e| e.to_string())?,
                DocumentFormat::Csv => officemd_csv::extract_ir::extract_ir(&owned_content)
                    .map_err(|e| e.to_string())?,
                DocumentFormat::Pptx => {
                    officemd_pptx::extract_ir(&owned_content).map_err(|e| e.to_string())?
                }
                DocumentFormat::Pdf => {
                    officemd_pdf::extract_ir(&owned_content).map_err(|e| e.to_string())?
                }
            };
        officemd_docling::convert_document_json(&doc).map_err(|e| e.to_string())
    })
    .map_err(to_py_err)
}

#[pyfunction(signature = (
    content,
    style_aware_values=false,
    streaming_rows=false,
    include_document_properties=false
))]
fn extract_tables_ir_json(
    py: Python<'_>,
    content: &[u8],
    style_aware_values: bool,
    streaming_rows: bool,
    include_document_properties: bool,
) -> PyResult<String> {
    let owned_content = content.to_vec();
    py.detach(move || {
        officemd_xlsx::table_ir::extract_tables_ir_json_with_options(
            &owned_content,
            style_aware_values,
            streaming_rows,
            include_document_properties,
        )
        .map_err(|e| e.to_string())
    })
    .map_err(to_py_err)
}

#[pyfunction(signature = (
    content,
    delimiter=',',
    include_document_properties=false
))]
fn extract_csv_tables_ir_json(
    py: Python<'_>,
    content: &[u8],
    delimiter: char,
    include_document_properties: bool,
) -> PyResult<String> {
    let mut utf8_buf = [0u8; 4];
    let encoded = delimiter.encode_utf8(&mut utf8_buf);
    if encoded.len() != 1 {
        return Err(pyo3::exceptions::PyValueError::new_err(
            "delimiter must be a single-byte character",
        ));
    }

    let owned_content = content.to_vec();
    let delimiter_byte = encoded.as_bytes()[0];
    py.detach(move || {
        officemd_csv::table_ir::extract_tables_ir_json_with_options(
            &owned_content,
            delimiter_byte,
            include_document_properties,
        )
        .map_err(|e| e.to_string())
    })
    .map_err(to_py_err)
}

#[pymodule]
fn _officemd(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(detect_format, m)?)?;
    m.add_function(wrap_pyfunction!(inspect_pdf_json, m)?)?;
    m.add_function(wrap_pyfunction!(inspect_pdf_fonts_json, m)?)?;
    m.add_function(wrap_pyfunction!(extract_ir_json, m)?)?;
    m.add_function(wrap_pyfunction!(docling_from_bytes, m)?)?;
    m.add_function(wrap_pyfunction!(markdown_from_bytes, m)?)?;
    m.add_function(wrap_pyfunction!(markdown_from_bytes_batch, m)?)?;
    m.add_function(wrap_pyfunction!(extract_sheet_names, m)?)?;
    m.add_function(wrap_pyfunction!(extract_tables_ir_json, m)?)?;
    m.add_function(wrap_pyfunction!(extract_csv_tables_ir_json, m)?)?;
    Ok(())
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
    fn parse_format_accepts_expected_variants() {
        assert!(matches!(parse_format("docx"), Some(DocumentFormat::Docx)));
        assert!(matches!(parse_format(".XLSX"), Some(DocumentFormat::Xlsx)));
        assert!(matches!(parse_format("csv"), Some(DocumentFormat::Csv)));
        assert!(matches!(parse_format("PPTX"), Some(DocumentFormat::Pptx)));
        assert!(matches!(parse_format("pdf"), Some(DocumentFormat::Pdf)));
    }

    #[test]
    fn detects_formats_from_minimal_packages() {
        let docx = build_zip(vec![("word/document.xml", "<w:document/>")]);
        let xlsx = build_zip(vec![("xl/workbook.xml", "<workbook/>")]);
        let pptx = build_zip(vec![("ppt/presentation.xml", "<p:presentation/>")]);

        assert!(matches!(
            detect_format_from_bytes(&docx),
            Ok(DocumentFormat::Docx)
        ));
        assert!(matches!(
            detect_format_from_bytes(&xlsx),
            Ok(DocumentFormat::Xlsx)
        ));
        assert!(matches!(
            detect_format_from_bytes(&pptx),
            Ok(DocumentFormat::Pptx)
        ));
        assert!(matches!(
            detect_format_from_bytes(b"%PDF-1.7\n"),
            Ok(DocumentFormat::Pdf)
        ));
    }

    #[test]
    fn resolve_format_uses_explicit_value() {
        let not_zip = b"not a zip archive";
        let resolved = resolve_format(not_zip, Some(".docx")).unwrap();
        assert!(matches!(resolved, DocumentFormat::Docx));
    }

    #[test]
    fn resolve_format_rejects_invalid_explicit_value() {
        let err = resolve_format(b"irrelevant", Some(".txt")).unwrap_err();
        assert!(err.contains("format must be one of"));
    }

    #[test]
    fn markdown_from_csv_requires_explicit_format() {
        let bytes = b"name,value\nwidget,42\n";
        Python::initialize();
        let markdown = Python::attach(|py| {
            markdown_from_bytes(
                py,
                bytes,
                Some(".csv".to_string()),
                false,
                true,
                true,
                true,
                Some("compact".to_string()),
                false,
            )
        })
        .expect("markdown");

        assert!(markdown.contains("## Sheet: Sheet1"));
        assert!(markdown.contains("| name | value |"));
    }

    #[test]
    fn markdown_from_bytes_batch_processes_multiple_csv_inputs() {
        let docs = vec![b"name,value\na,1\n".to_vec(), b"name,value\nb,2\n".to_vec()];
        Python::initialize();
        let markdowns = Python::attach(|py| {
            markdown_from_bytes_batch(
                py,
                docs,
                Some(".csv".to_string()),
                Some(2),
                false,
                true,
                true,
                true,
                Some("compact".to_string()),
            )
        })
        .expect("markdown batch");
        assert_eq!(markdowns.len(), 2);
        assert!(markdowns[0].contains("| name | value |"));
        assert!(markdowns[1].contains("| name | value |"));
    }

    #[test]
    fn markdown_from_xlsx_includes_document_properties_when_requested() {
        let bytes = minimal_xlsx_with_doc_props();
        Python::initialize();
        let markdown = Python::attach(|py| {
            markdown_from_bytes(
                py,
                &bytes,
                Some(".xlsx".to_string()),
                true,
                true,
                true,
                true,
                Some("human".to_string()),
                false,
            )
        })
        .expect("markdown");

        assert!(markdown.contains("### Document Properties"));
        assert!(markdown.contains("Quarterly Results"));
    }

    #[test]
    fn extract_tables_ir_json_omits_properties_by_default() {
        let bytes = minimal_xlsx_with_doc_props();
        Python::initialize();
        let payload = Python::attach(|py| extract_tables_ir_json(py, &bytes, false, false, false))
            .expect("extract tables ir json");
        let value: serde_json::Value = serde_json::from_str(&payload).expect("valid json");

        assert!(value["properties"].is_null());
    }

    #[test]
    fn extract_tables_ir_json_includes_properties_when_requested() {
        let bytes = minimal_xlsx_with_doc_props();
        Python::initialize();
        let payload = Python::attach(|py| extract_tables_ir_json(py, &bytes, false, false, true))
            .expect("extract tables ir json");
        let value: serde_json::Value = serde_json::from_str(&payload).expect("valid json");

        assert!(value["properties"].is_object());
        let core = value["properties"]["core"]
            .as_object()
            .expect("core properties map");
        assert!(!core.is_empty());
    }
}
