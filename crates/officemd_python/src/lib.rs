//! Native Python bindings for OfficeMD extraction and rendering.

// PyO3 `#[pyfunction]` signatures must take owned values, not references.
#![allow(clippy::needless_pass_by_value, clippy::too_many_arguments)]

use officemd_core::format::{
    DocumentFormat, detect_format_from_bytes, parse_format, resolve_format, resolve_worker_count,
};
use officemd_core::{
    DocxPatch, PptxPatch, XlsxPatch, apply_ooxml_patch_json as apply_ooxml_patch_json_core,
    patch_docx_batch_json as patch_docx_batch_json_core,
    patch_docx_batch_json_with_report as patch_docx_batch_json_with_report_core,
    patch_docx_json as patch_docx_json_core, patch_docx_with_report as patch_docx_with_report_core,
    patch_pptx_batch_json as patch_pptx_batch_json_core,
    patch_pptx_batch_json_with_report as patch_pptx_batch_json_with_report_core,
    patch_pptx_json as patch_pptx_json_core, patch_pptx_with_report as patch_pptx_with_report_core,
    patch_xlsx_batch_json as patch_xlsx_batch_json_core,
    patch_xlsx_batch_json_with_report as patch_xlsx_batch_json_with_report_core,
    patch_xlsx_json as patch_xlsx_json_core, patch_xlsx_with_report as patch_xlsx_with_report_core,
};
use officemd_markdown::{MarkdownProfile, RenderOptions};
use pyo3::prelude::*;
use rayon::ThreadPoolBuilder;
use rayon::prelude::*;

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

fn apply_ooxml_patch_json_impl(content: &[u8], patch_json: &str) -> Result<Vec<u8>, String> {
    apply_ooxml_patch_json_core(content, patch_json).map_err(|e| e.to_string())
}

fn patch_docx_json_impl(content: &[u8], patch_json: &str) -> Result<Vec<u8>, String> {
    patch_docx_json_core(content, patch_json).map_err(|e| e.to_string())
}

fn patch_pptx_json_impl(content: &[u8], patch_json: &str) -> Result<Vec<u8>, String> {
    patch_pptx_json_core(content, patch_json).map_err(|e| e.to_string())
}

#[pyfunction(signature = (content, patch_json))]
fn apply_ooxml_patch_json(py: Python<'_>, content: &[u8], patch_json: String) -> PyResult<Vec<u8>> {
    let owned_content = content.to_vec();
    py.detach(move || apply_ooxml_patch_json_impl(&owned_content, &patch_json))
        .map_err(to_py_err)
}

#[pyfunction(signature = (content, patch_json))]
fn _patch_docx_json(py: Python<'_>, content: &[u8], patch_json: String) -> PyResult<Vec<u8>> {
    let owned_content = content.to_vec();
    py.detach(move || patch_docx_json_impl(&owned_content, &patch_json))
        .map_err(to_py_err)
}

#[pyfunction(signature = (content, patch_json))]
fn _patch_pptx_json(py: Python<'_>, content: &[u8], patch_json: String) -> PyResult<Vec<u8>> {
    let owned_content = content.to_vec();
    py.detach(move || patch_pptx_json_impl(&owned_content, &patch_json))
        .map_err(to_py_err)
}

fn patch_docx_batch_json_impl(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, String> {
    patch_docx_batch_json_core(contents, patch_json, workers).map_err(|e| e.to_string())
}

fn patch_pptx_batch_json_impl(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, String> {
    patch_pptx_batch_json_core(contents, patch_json, workers).map_err(|e| e.to_string())
}

#[pyfunction(signature = (contents, patch_json, workers=None))]
fn _patch_docx_batch_json(
    py: Python<'_>,
    contents: Vec<Vec<u8>>,
    patch_json: String,
    workers: Option<usize>,
) -> PyResult<Vec<Vec<u8>>> {
    py.detach(move || patch_docx_batch_json_impl(contents, &patch_json, workers))
        .map_err(to_py_err)
}

#[pyfunction(signature = (contents, patch_json, workers=None))]
fn _patch_pptx_batch_json(
    py: Python<'_>,
    contents: Vec<Vec<u8>>,
    patch_json: String,
    workers: Option<usize>,
) -> PyResult<Vec<Vec<u8>>> {
    py.detach(move || patch_pptx_batch_json_impl(contents, &patch_json, workers))
        .map_err(to_py_err)
}

fn patch_docx_batch_json_with_report_impl(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<String, String> {
    let results = patch_docx_batch_json_with_report_core(contents, patch_json, workers)
        .map_err(|e| e.to_string())?;
    serde_json::to_string(&results).map_err(|e| e.to_string())
}

fn patch_pptx_batch_json_with_report_impl(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<String, String> {
    let results = patch_pptx_batch_json_with_report_core(contents, patch_json, workers)
        .map_err(|e| e.to_string())?;
    serde_json::to_string(&results).map_err(|e| e.to_string())
}

fn patch_docx_json_with_report_impl(content: &[u8], patch_json: &str) -> Result<String, String> {
    let patch: DocxPatch = serde_json::from_str(patch_json).map_err(|e| e.to_string())?;
    let result = patch_docx_with_report_core(content, &patch).map_err(|e| e.to_string())?;
    serde_json::to_string(&result).map_err(|e| e.to_string())
}

fn patch_pptx_json_with_report_impl(content: &[u8], patch_json: &str) -> Result<String, String> {
    let patch: PptxPatch = serde_json::from_str(patch_json).map_err(|e| e.to_string())?;
    let result = patch_pptx_with_report_core(content, &patch).map_err(|e| e.to_string())?;
    serde_json::to_string(&result).map_err(|e| e.to_string())
}

#[pyfunction(signature = (content, patch_json))]
fn _patch_docx_json_with_report(
    py: Python<'_>,
    content: &[u8],
    patch_json: String,
) -> PyResult<String> {
    let owned_content = content.to_vec();
    py.detach(move || patch_docx_json_with_report_impl(&owned_content, &patch_json))
        .map_err(to_py_err)
}

#[pyfunction(signature = (content, patch_json))]
fn _patch_pptx_json_with_report(
    py: Python<'_>,
    content: &[u8],
    patch_json: String,
) -> PyResult<String> {
    let owned_content = content.to_vec();
    py.detach(move || patch_pptx_json_with_report_impl(&owned_content, &patch_json))
        .map_err(to_py_err)
}

#[pyfunction(signature = (contents, patch_json, workers=None))]
fn _patch_docx_batch_json_with_report(
    py: Python<'_>,
    contents: Vec<Vec<u8>>,
    patch_json: String,
    workers: Option<usize>,
) -> PyResult<String> {
    py.detach(move || patch_docx_batch_json_with_report_impl(contents, &patch_json, workers))
        .map_err(to_py_err)
}

#[pyfunction(signature = (contents, patch_json, workers=None))]
fn _patch_pptx_batch_json_with_report(
    py: Python<'_>,
    contents: Vec<Vec<u8>>,
    patch_json: String,
    workers: Option<usize>,
) -> PyResult<String> {
    py.detach(move || patch_pptx_batch_json_with_report_impl(contents, &patch_json, workers))
        .map_err(to_py_err)
}

fn patch_xlsx_json_impl(content: &[u8], patch_json: &str) -> Result<Vec<u8>, String> {
    patch_xlsx_json_core(content, patch_json).map_err(|e| e.to_string())
}

fn patch_xlsx_json_with_report_impl(content: &[u8], patch_json: &str) -> Result<String, String> {
    let patch: XlsxPatch = serde_json::from_str(patch_json).map_err(|e| e.to_string())?;
    let result = patch_xlsx_with_report_core(content, &patch).map_err(|e| e.to_string())?;
    serde_json::to_string(&result).map_err(|e| e.to_string())
}

fn patch_xlsx_batch_json_impl(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, String> {
    patch_xlsx_batch_json_core(contents, patch_json, workers).map_err(|e| e.to_string())
}

fn patch_xlsx_batch_json_with_report_impl(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<String, String> {
    let results = patch_xlsx_batch_json_with_report_core(contents, patch_json, workers)
        .map_err(|e| e.to_string())?;
    serde_json::to_string(&results).map_err(|e| e.to_string())
}

#[pyfunction(signature = (content, patch_json))]
fn _patch_xlsx_json(py: Python<'_>, content: &[u8], patch_json: String) -> PyResult<Vec<u8>> {
    let owned_content = content.to_vec();
    py.detach(move || patch_xlsx_json_impl(&owned_content, &patch_json))
        .map_err(to_py_err)
}

#[pyfunction(signature = (content, patch_json))]
fn _patch_xlsx_json_with_report(
    py: Python<'_>,
    content: &[u8],
    patch_json: String,
) -> PyResult<String> {
    let owned_content = content.to_vec();
    py.detach(move || patch_xlsx_json_with_report_impl(&owned_content, &patch_json))
        .map_err(to_py_err)
}

#[pyfunction(signature = (contents, patch_json, workers=None))]
fn _patch_xlsx_batch_json(
    py: Python<'_>,
    contents: Vec<Vec<u8>>,
    patch_json: String,
    workers: Option<usize>,
) -> PyResult<Vec<Vec<u8>>> {
    py.detach(move || patch_xlsx_batch_json_impl(contents, &patch_json, workers))
        .map_err(to_py_err)
}

#[pyfunction(signature = (contents, patch_json, workers=None))]
fn _patch_xlsx_batch_json_with_report(
    py: Python<'_>,
    contents: Vec<Vec<u8>>,
    patch_json: String,
    workers: Option<usize>,
) -> PyResult<String> {
    py.detach(move || patch_xlsx_batch_json_with_report_impl(contents, &patch_json, workers))
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
    m.add_function(wrap_pyfunction!(create_document_from_markdown, m)?)?;
    m.add_function(wrap_pyfunction!(apply_ooxml_patch_json, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_docx_json, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_pptx_json, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_xlsx_json, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_docx_json_with_report, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_pptx_json_with_report, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_xlsx_json_with_report, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_docx_batch_json, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_pptx_batch_json, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_xlsx_batch_json, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_docx_batch_json_with_report, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_pptx_batch_json_with_report, m)?)?;
    m.add_function(wrap_pyfunction!(_patch_xlsx_batch_json_with_report, m)?)?;
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
    use officemd_core::test_helpers::{build_zip, minimal_xlsx_with_doc_props};

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
