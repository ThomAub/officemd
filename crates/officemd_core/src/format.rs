//! Document format detection and resolution utilities.
//!
//! Provides a shared [`DocumentFormat`] enum and helpers used by all binding
//! and CLI crates. Distinct from [`crate::ir::DocumentKind`] which represents
//! the IR document type and does not include CSV.

use std::fmt;

use crate::opc::OpcPackage;

/// Supported document formats for input processing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentFormat {
    Docx,
    Xlsx,
    Csv,
    Pptx,
    Pdf,
}

impl DocumentFormat {
    /// File extension including the leading dot.
    #[must_use]
    pub fn extension(self) -> &'static str {
        match self {
            Self::Docx => ".docx",
            Self::Xlsx => ".xlsx",
            Self::Csv => ".csv",
            Self::Pptx => ".pptx",
            Self::Pdf => ".pdf",
        }
    }
}

impl fmt::Display for DocumentFormat {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Docx => write!(f, "docx"),
            Self::Xlsx => write!(f, "xlsx"),
            Self::Csv => write!(f, "csv"),
            Self::Pptx => write!(f, "pptx"),
            Self::Pdf => write!(f, "pdf"),
        }
    }
}

/// Parse a format string (e.g. `"docx"`, `".xlsx"`) into a [`DocumentFormat`].
///
/// Case-insensitive, accepts with or without leading dot.
#[must_use]
pub fn parse_format(value: &str) -> Option<DocumentFormat> {
    let v = value.trim();
    if v.eq_ignore_ascii_case("docx") || v.eq_ignore_ascii_case(".docx") {
        Some(DocumentFormat::Docx)
    } else if v.eq_ignore_ascii_case("xlsx") || v.eq_ignore_ascii_case(".xlsx") {
        Some(DocumentFormat::Xlsx)
    } else if v.eq_ignore_ascii_case("csv") || v.eq_ignore_ascii_case(".csv") {
        Some(DocumentFormat::Csv)
    } else if v.eq_ignore_ascii_case("pptx") || v.eq_ignore_ascii_case(".pptx") {
        Some(DocumentFormat::Pptx)
    } else if v.eq_ignore_ascii_case("pdf") || v.eq_ignore_ascii_case(".pdf") {
        Some(DocumentFormat::Pdf)
    } else {
        None
    }
}

/// Returns `true` when the byte content starts with a PDF header (`%PDF-`),
/// after stripping any UTF-8 BOM and leading ASCII whitespace.
#[must_use]
pub fn looks_like_pdf_header(content: &[u8]) -> bool {
    if content.is_empty() {
        return false;
    }
    let header = &content[..content.len().min(1024)];
    let trimmed = strip_bom_and_whitespace(header);
    trimmed.starts_with(b"%PDF-")
}

/// Detect the document format by inspecting file content bytes.
///
/// Checks for PDF header first, then probes OPC package parts.
/// CSV cannot be detected from bytes alone and requires an explicit format.
///
/// # Errors
///
/// Returns a descriptive error string if the format cannot be determined.
pub fn detect_format_from_bytes(content: &[u8]) -> Result<DocumentFormat, String> {
    if looks_like_pdf_header(content) {
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

    Err("Could not detect format from file content \
         (supported: .docx, .xlsx, .csv, .pptx, .pdf; csv requires explicit format)"
        .to_string())
}

/// Resolve the document format: use the explicit string if provided, otherwise
/// detect from bytes.
///
/// # Errors
///
/// Returns an error if the explicit format is unrecognized or byte detection fails.
pub fn resolve_format(content: &[u8], format: Option<&str>) -> Result<DocumentFormat, String> {
    match format {
        Some(value) => parse_format(value)
            .ok_or_else(|| "format must be one of: .docx, .xlsx, .csv, .pptx, .pdf".to_string()),
        None => detect_format_from_bytes(content),
    }
}

/// Resolve the number of worker threads from an optional user hint.
///
/// Returns 1 when parallelism cannot be detected or on WASM targets.
#[must_use]
pub fn resolve_worker_count(workers: Option<usize>) -> usize {
    workers.filter(|v| *v > 0).unwrap_or_else(|| {
        #[cfg(not(target_arch = "wasm32"))]
        {
            std::thread::available_parallelism().map_or(1, usize::from)
        }
        #[cfg(target_arch = "wasm32")]
        {
            1
        }
    })
}

fn strip_bom_and_whitespace(bytes: &[u8]) -> &[u8] {
    let b = if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        &bytes[3..]
    } else {
        bytes
    };
    let start = b
        .iter()
        .position(|c| !c.is_ascii_whitespace())
        .unwrap_or(b.len());
    &b[start..]
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_format_accepts_case_insensitive_values() {
        assert!(matches!(parse_format("DOCX"), Some(DocumentFormat::Docx)));
        assert!(matches!(parse_format(".xlsx"), Some(DocumentFormat::Xlsx)));
        assert!(matches!(parse_format("csv"), Some(DocumentFormat::Csv)));
        assert!(matches!(parse_format("Pptx"), Some(DocumentFormat::Pptx)));
        assert!(matches!(parse_format("pdf"), Some(DocumentFormat::Pdf)));
        assert!(parse_format(".txt").is_none());
    }

    #[test]
    fn looks_like_pdf_detects_header() {
        assert!(looks_like_pdf_header(b"%PDF-1.7\n"));
        assert!(looks_like_pdf_header(b"\xEF\xBB\xBF   %PDF-1.4\n"));
        assert!(!looks_like_pdf_header(b"not a pdf"));
        assert!(!looks_like_pdf_header(b""));
    }

    #[test]
    fn detect_format_from_minimal_packages() {
        use crate::test_helpers::build_zip;

        let docx = build_zip(vec![("word/document.xml", "<w:document/>")]);
        assert_eq!(
            detect_format_from_bytes(&docx).unwrap(),
            DocumentFormat::Docx
        );

        let xlsx = build_zip(vec![("xl/workbook.xml", "<workbook/>")]);
        assert_eq!(
            detect_format_from_bytes(&xlsx).unwrap(),
            DocumentFormat::Xlsx
        );

        let pptx = build_zip(vec![("ppt/presentation.xml", "<p:presentation/>")]);
        assert_eq!(
            detect_format_from_bytes(&pptx).unwrap(),
            DocumentFormat::Pptx
        );

        assert_eq!(
            detect_format_from_bytes(b"%PDF-1.7\n").unwrap(),
            DocumentFormat::Pdf
        );
    }

    #[test]
    fn resolve_format_uses_explicit_value() {
        assert_eq!(
            resolve_format(b"irrelevant", Some(".docx")).unwrap(),
            DocumentFormat::Docx,
        );
    }

    #[test]
    fn resolve_format_rejects_invalid_explicit_value() {
        let err = resolve_format(b"irrelevant", Some(".txt")).unwrap_err();
        assert!(err.contains("format must be one of"));
    }

    #[test]
    fn resolve_worker_count_uses_hint_when_positive() {
        assert_eq!(resolve_worker_count(Some(4)), 4);
        assert!(resolve_worker_count(Some(0)) >= 1);
        assert!(resolve_worker_count(None) >= 1);
    }

    #[test]
    fn display_formats() {
        assert_eq!(DocumentFormat::Docx.to_string(), "docx");
        assert_eq!(DocumentFormat::Csv.to_string(), "csv");
    }

    #[test]
    fn extension_includes_dot() {
        assert_eq!(DocumentFormat::Docx.extension(), ".docx");
        assert_eq!(DocumentFormat::Pdf.extension(), ".pdf");
    }
}
