#![allow(clippy::all)]
//! Smart PDF detection and text extraction using lopdf
//!
//! # Quick start
//!
//! ```ignore
//! // Full processing (detect + extract + markdown) with defaults
//! let result = pdf_inspector::process_pdf("document.pdf").unwrap();
//! println!("type: {:?}, pages: {}", result.pdf_type, result.page_count);
//! if let Some(md) = &result.markdown {
//!     println!("{md}");
//! }
//!
//! // Fast metadata-only detection (no text extraction)
//! let info = pdf_inspector::detect_pdf("document.pdf").unwrap();
//! println!("type: {:?}, pages: {}", info.pdf_type, info.page_count);
//!
//! // Custom options via builder
//! use pdf_inspector::{PdfOptions, ProcessMode};
//! let result = pdf_inspector::process_pdf_with_options(
//!     "document.pdf",
//!     PdfOptions::new().mode(ProcessMode::Analyze),
//! ).unwrap();
//! ```

pub mod adobe_korea1;
pub mod detector;
pub mod extractor;
pub mod glyph_names;
pub mod markdown;
pub mod process_mode;
pub mod tables;
pub mod text_utils;
pub mod tounicode;
pub mod types;

pub use detector::{
    DetectionConfig, PdfType, PdfTypeResult, ScanStrategy, detect_pdf_type, detect_pdf_type_mem,
    detect_pdf_type_mem_with_config, detect_pdf_type_with_config,
};
pub use extractor::{extract_text, extract_text_with_positions, extract_text_with_positions_pages};
pub use markdown::{
    MarkdownOptions, to_markdown, to_markdown_from_items, to_markdown_from_items_with_rects,
};
pub use process_mode::ProcessMode;
pub use types::{LayoutComplexity, PdfRect, TextItem};

use lopdf::Document;
use std::collections::HashSet;
use std::path::Path;
use tounicode::FontCMaps;

#[cfg(not(target_arch = "wasm32"))]
type ProcessTimer = std::time::Instant;
#[cfg(target_arch = "wasm32")]
type ProcessTimer = f64;

#[cfg(not(target_arch = "wasm32"))]
fn start_timer() -> ProcessTimer {
    std::time::Instant::now()
}

#[cfg(target_arch = "wasm32")]
fn start_timer() -> ProcessTimer {
    js_sys::Date::now()
}

#[cfg(not(target_arch = "wasm32"))]
fn elapsed_ms(start: ProcessTimer) -> u64 {
    start.elapsed().as_millis() as u64
}

#[cfg(target_arch = "wasm32")]
fn elapsed_ms(start: ProcessTimer) -> u64 {
    (js_sys::Date::now() - start).max(0.0) as u64
}

// =========================================================================
// Result type
// =========================================================================

/// High-level PDF processing result.
#[derive(Debug)]
pub struct PdfProcessResult {
    /// The detected PDF type.
    pub pdf_type: PdfType,
    /// Markdown output (populated in [`ProcessMode::Full`], `None` otherwise).
    pub markdown: Option<String>,
    /// Page count.
    pub page_count: u32,
    /// Processing time in milliseconds.
    pub processing_time_ms: u64,
    /// 1-indexed page numbers that need OCR.
    pub pages_needing_ocr: Vec<u32>,
    /// Title from PDF metadata (if available).
    pub title: Option<String>,
    /// Detection confidence score (0.0–1.0).
    pub confidence: f32,
    /// Layout complexity analysis (tables, multi-column detection).
    pub layout: LayoutComplexity,
    /// `true` when broken font encodings are detected (garbled text,
    /// replacement characters). Clients should fall back to OCR.
    pub has_encoding_issues: bool,
}

// =========================================================================
// Options builder
// =========================================================================

/// Configuration for [`process_pdf_with_options`] and friends.
///
/// Use the builder methods to customise behaviour:
///
/// ```ignore
/// use pdf_inspector::{PdfOptions, ProcessMode};
///
/// let opts = PdfOptions::new()
///     .mode(ProcessMode::Analyze)
///     .pages([1, 3, 5]);
/// ```
#[derive(Debug, Clone)]
pub struct PdfOptions {
    /// How far the pipeline should run (default: [`ProcessMode::Full`]).
    pub mode: ProcessMode,
    /// Detection configuration.
    pub detection: DetectionConfig,
    /// Markdown formatting options (only used in [`ProcessMode::Full`]).
    pub markdown: MarkdownOptions,
    /// Optional set of 1-indexed pages to process.  `None` = all pages.
    pub page_filter: Option<HashSet<u32>>,
    /// When `true`, attempt text extraction even for Scanned/ImageBased PDFs
    /// instead of returning `None` immediately.  Extraction failure is treated
    /// as non-fatal (like Mixed PDFs).
    pub force_extract: bool,
}

impl Default for PdfOptions {
    fn default() -> Self {
        Self {
            mode: ProcessMode::Full,
            detection: DetectionConfig::default(),
            markdown: MarkdownOptions::default(),
            page_filter: None,
            force_extract: false,
        }
    }
}

impl PdfOptions {
    /// Create options with all defaults ([`ProcessMode::Full`]).
    pub fn new() -> Self {
        Self::default()
    }

    /// Shorthand for detect-only options.
    pub fn detect_only() -> Self {
        Self {
            mode: ProcessMode::DetectOnly,
            ..Self::default()
        }
    }

    /// Set the processing mode.
    pub fn mode(mut self, mode: ProcessMode) -> Self {
        self.mode = mode;
        self
    }

    /// Set detection configuration.
    pub fn detection(mut self, config: DetectionConfig) -> Self {
        self.detection = config;
        self
    }

    /// Set markdown formatting options.
    pub fn markdown(mut self, options: MarkdownOptions) -> Self {
        self.markdown = options;
        self
    }

    /// Limit processing to specific 1-indexed pages.
    pub fn pages(mut self, pages: impl IntoIterator<Item = u32>) -> Self {
        self.page_filter = Some(pages.into_iter().collect());
        self
    }

    /// Force text extraction even for Scanned/ImageBased PDFs.
    pub fn force_extract(mut self, force: bool) -> Self {
        self.force_extract = force;
        self
    }
}

// =========================================================================
// Public convenience functions
// =========================================================================

/// Process a PDF file with full extraction (detect → extract → markdown).
///
/// This is the most common entry point.  Equivalent to
/// `process_pdf_with_options(path, PdfOptions::new())`.
pub fn process_pdf<P: AsRef<Path>>(path: P) -> Result<PdfProcessResult, PdfError> {
    process_pdf_with_options(path, PdfOptions::new())
}

/// Fast metadata-only detection — no text extraction or markdown generation.
///
/// Equivalent to `process_pdf_with_options(path, PdfOptions::detect_only())`.
pub fn detect_pdf<P: AsRef<Path>>(path: P) -> Result<PdfProcessResult, PdfError> {
    process_pdf_with_options(path, PdfOptions::detect_only())
}

/// Process a PDF file with custom options.
///
/// The document is loaded **once** and shared between detection and extraction.
pub fn process_pdf_with_options<P: AsRef<Path>>(
    path: P,
    options: PdfOptions,
) -> Result<PdfProcessResult, PdfError> {
    let start = start_timer();
    validate_pdf_file(&path)?;

    // Load the document once — shared by detection AND extraction.
    let (doc, page_count) = load_document_from_path(&path)?;

    process_document(doc, page_count, options, start)
}

/// Process a PDF from a memory buffer with full extraction.
pub fn process_pdf_mem(buffer: &[u8]) -> Result<PdfProcessResult, PdfError> {
    process_pdf_mem_with_options(buffer, PdfOptions::new())
}

/// Fast metadata-only detection from a memory buffer.
pub fn detect_pdf_mem(buffer: &[u8]) -> Result<PdfProcessResult, PdfError> {
    process_pdf_mem_with_options(buffer, PdfOptions::detect_only())
}

/// Process a PDF from a memory buffer with custom options.
///
/// The buffer is parsed **once** and shared between detection and extraction.
pub fn process_pdf_mem_with_options(
    buffer: &[u8],
    options: PdfOptions,
) -> Result<PdfProcessResult, PdfError> {
    let start = start_timer();
    validate_pdf_bytes(buffer)?;

    let (doc, page_count) = load_document_from_mem(buffer)?;

    process_document(doc, page_count, options, start)
}

// =========================================================================
// Deprecated compat shims
// =========================================================================

/// Process a PDF file with custom detection and markdown configuration.
#[deprecated(since = "0.2.0", note = "Use process_pdf_with_options instead")]
pub fn process_pdf_with_config<P: AsRef<Path>>(
    path: P,
    config: DetectionConfig,
    markdown_options: MarkdownOptions,
) -> Result<PdfProcessResult, PdfError> {
    process_pdf_with_options(
        path,
        PdfOptions::new()
            .detection(config)
            .markdown(markdown_options),
    )
}

/// Process a PDF file with custom configuration and optional page filter.
#[deprecated(since = "0.2.0", note = "Use process_pdf_with_options instead")]
pub fn process_pdf_with_config_pages<P: AsRef<Path>>(
    path: P,
    config: DetectionConfig,
    markdown_options: MarkdownOptions,
    page_filter: Option<&HashSet<u32>>,
) -> Result<PdfProcessResult, PdfError> {
    let mut opts = PdfOptions::new()
        .detection(config)
        .markdown(markdown_options);
    opts.page_filter = page_filter.cloned();
    process_pdf_with_options(path, opts)
}

/// Process PDF from memory buffer with custom detection and markdown configuration.
#[deprecated(since = "0.2.0", note = "Use process_pdf_mem_with_options instead")]
pub fn process_pdf_mem_with_config(
    buffer: &[u8],
    config: DetectionConfig,
    markdown_options: MarkdownOptions,
) -> Result<PdfProcessResult, PdfError> {
    process_pdf_mem_with_options(
        buffer,
        PdfOptions::new()
            .detection(config)
            .markdown(markdown_options),
    )
}

// =========================================================================
// Internal: single-load document pipeline
// =========================================================================

/// Load a PDF from disk, returning the parsed document and page count.
///
/// `Document::load_metadata` for page count + `Document::load` for content
/// are combined here, but lopdf loads the full doc in `load()` so we extract
/// page count from it directly to avoid the metadata-only round-trip.
fn load_document_from_path<P: AsRef<Path>>(path: P) -> Result<(Document, u32), PdfError> {
    let doc = match Document::load(&path) {
        Ok(d) => d,
        Err(ref e) if is_encrypted_lopdf_error(e) => Document::load_with_password(&path, "")?,
        Err(e) => return Err(e.into()),
    };
    let page_count = doc.get_pages().len() as u32;
    Ok((doc, page_count))
}

/// Load a PDF from a memory buffer.
fn load_document_from_mem(buffer: &[u8]) -> Result<(Document, u32), PdfError> {
    let doc = match Document::load_mem(buffer) {
        Ok(d) => d,
        Err(ref e) if is_encrypted_lopdf_error(e) => Document::load_mem_with_password(buffer, "")?,
        Err(e) => return Err(e.into()),
    };
    let page_count = doc.get_pages().len() as u32;
    Ok((doc, page_count))
}

/// Core processing pipeline operating on a pre-loaded document.
fn process_document(
    doc: Document,
    page_count: u32,
    options: PdfOptions,
    start: ProcessTimer,
) -> Result<PdfProcessResult, PdfError> {
    // Step 1 — Detection (cheap: scans content streams for text operators)
    let detection = detector::detect_from_document(&doc, page_count, &options.detection)?;
    let pdf_type = detection.pdf_type;
    let pages_needing_ocr = detection.pages_needing_ocr;
    let title = detection.title;
    let confidence = detection.confidence;

    // DetectOnly → return immediately
    if options.mode == ProcessMode::DetectOnly {
        return Ok(PdfProcessResult {
            pdf_type,
            markdown: None,
            page_count,
            processing_time_ms: elapsed_ms(start),
            pages_needing_ocr,
            title,
            confidence,
            layout: LayoutComplexity::default(),
            has_encoding_issues: false,
        });
    }

    // Scanned / ImageBased → nothing to extract (unless force_extract is set)
    if matches!(pdf_type, PdfType::Scanned | PdfType::ImageBased) && !options.force_extract {
        return Ok(PdfProcessResult {
            pdf_type,
            markdown: None,
            page_count,
            processing_time_ms: elapsed_ms(start),
            pages_needing_ocr,
            title,
            confidence,
            layout: LayoutComplexity::default(),
            has_encoding_issues: false,
        });
    }

    // Step 2 — Extraction (reuses the already-loaded document)
    let extracted = {
        let font_cmaps = FontCMaps::from_doc(&doc);
        extractor::extract_positioned_text_from_doc(&doc, &font_cmaps, options.page_filter.as_ref())
    };

    // For Mixed PDFs (or forced Scanned/ImageBased), extraction failure is non-fatal
    let non_fatal = pdf_type == PdfType::Mixed
        || (options.force_extract && matches!(pdf_type, PdfType::Scanned | PdfType::ImageBased));
    let extracted = if non_fatal {
        extracted.ok()
    } else {
        Some(extracted?)
    };

    let (markdown, layout, has_encoding_issues) = match extracted {
        Some((items, rects)) => {
            let layout = compute_layout_complexity(&items, &rects);

            let md = if options.mode == ProcessMode::Analyze {
                None
            } else {
                let mut markdown_options = options.markdown.clone();
                markdown_options.include_page_numbers = true;
                Some(to_markdown_from_items_with_rects(
                    items,
                    markdown_options,
                    &rects,
                ))
            };

            let enc = md.as_ref().is_some_and(|m| detect_encoding_issues(m));
            (md, layout, enc)
        }
        None => (None, LayoutComplexity::default(), false),
    };

    Ok(PdfProcessResult {
        pdf_type,
        markdown,
        page_count,
        processing_time_ms: elapsed_ms(start),
        pages_needing_ocr,
        title,
        confidence,
        layout,
        has_encoding_issues,
    })
}

// =========================================================================
// Internal helpers
// =========================================================================

/// Detect broken font encodings in extracted markdown text.
///
/// Two heuristics:
/// 1. **U+FFFD**: Any replacement character indicates decode failures.
/// 2. **Dollar-as-space**: Pattern like `Word$Word$Word` where `$` is used as a
///    word separator due to broken ToUnicode CMaps. Triggers when either:
///    - More than 50% of `$` are between letters (clear substitution pattern), OR
///    - More than 20 letter-dollar-letter occurrences (even if some `$` are also
///      used as trailing/leading separators, 20+ is far beyond normal financial text).
fn detect_encoding_issues(markdown: &str) -> bool {
    // Heuristic 1: U+FFFD replacement characters
    if markdown.contains('\u{FFFD}') {
        return true;
    }

    // Heuristic 2: dollar-as-space pattern
    let total_dollars = markdown.matches('$').count();
    if total_dollars > 10 {
        let bytes = markdown.as_bytes();
        let mut letter_dollar_letter = 0usize;
        for i in 1..bytes.len().saturating_sub(1) {
            if bytes[i] == b'$'
                && bytes[i - 1].is_ascii_alphabetic()
                && bytes[i + 1].is_ascii_alphabetic()
            {
                letter_dollar_letter += 1;
            }
        }
        if letter_dollar_letter > 20 || letter_dollar_letter * 2 > total_dollars {
            return true;
        }
    }

    false
}

/// Analyse extracted items and rects for layout complexity.
fn compute_layout_complexity(
    items: &[types::TextItem],
    rects: &[types::PdfRect],
) -> LayoutComplexity {
    let mut seen_pages: Vec<u32> = items.iter().map(|i| i.page).collect();
    seen_pages.sort();
    seen_pages.dedup();

    let mut pages_with_tables: Vec<u32> = Vec::new();
    for &page in &seen_pages {
        let (tables, _) = tables::detect_tables_from_rects(items, rects, page);
        if !tables.is_empty() {
            pages_with_tables.push(page);
        }
    }

    let mut pages_with_columns: Vec<u32> = Vec::new();
    for page in seen_pages {
        let cols = extractor::detect_columns(items, page);
        if cols.len() >= 2 {
            pages_with_columns.push(page);
        }
    }

    let is_complex = !pages_with_tables.is_empty() || !pages_with_columns.is_empty();

    LayoutComplexity {
        is_complex,
        pages_with_tables,
        pages_with_columns,
    }
}

#[derive(Debug, thiserror::Error)]
pub enum PdfError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("PDF parsing error: {0}")]
    Parse(String),
    #[error("PDF is encrypted")]
    Encrypted,
    #[error("Invalid PDF structure")]
    InvalidStructure,
    #[error("Not a PDF: {0}")]
    NotAPdf(String),
}

impl From<lopdf::Error> for PdfError {
    fn from(e: lopdf::Error) -> Self {
        match e {
            lopdf::Error::IO(io_err) => PdfError::Io(io_err),
            lopdf::Error::Decryption(_)
            | lopdf::Error::InvalidPassword
            | lopdf::Error::AlreadyEncrypted
            | lopdf::Error::UnsupportedSecurityHandler(_) => PdfError::Encrypted,
            lopdf::Error::Unimplemented(msg) if msg.contains("encrypted") => PdfError::Encrypted,
            lopdf::Error::Parse(ref pe) if pe.to_string().contains("invalid file header") => {
                PdfError::NotAPdf("invalid PDF file header".to_string())
            }
            lopdf::Error::MissingXrefEntry
            | lopdf::Error::Xref(_)
            | lopdf::Error::IndirectObject { .. }
            | lopdf::Error::ObjectIdMismatch
            | lopdf::Error::InvalidObjectStream(_)
            | lopdf::Error::InvalidOffset(_) => PdfError::InvalidStructure,
            other => PdfError::Parse(other.to_string()),
        }
    }
}

/// Check whether a `lopdf::Error` represents an encryption-related failure.
pub(crate) fn is_encrypted_lopdf_error(e: &lopdf::Error) -> bool {
    matches!(
        e,
        lopdf::Error::Decryption(_)
            | lopdf::Error::InvalidPassword
            | lopdf::Error::AlreadyEncrypted
            | lopdf::Error::UnsupportedSecurityHandler(_)
    ) || matches!(e, lopdf::Error::Unimplemented(msg) if msg.contains("encrypted"))
}

// ---------------------------------------------------------------------------
// PDF validation helpers
// ---------------------------------------------------------------------------

/// Strip UTF-8 BOM and leading ASCII whitespace from a byte slice.
fn strip_bom_and_whitespace(bytes: &[u8]) -> &[u8] {
    let b = if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        &bytes[3..]
    } else {
        bytes
    };
    let start = b
        .iter()
        .position(|&c| !c.is_ascii_whitespace())
        .unwrap_or(b.len());
    &b[start..]
}

/// Case-insensitive prefix check on byte slices.
fn starts_with_ci(haystack: &[u8], needle: &[u8]) -> bool {
    if haystack.len() < needle.len() {
        return false;
    }
    haystack[..needle.len()]
        .iter()
        .zip(needle)
        .all(|(a, b)| a.eq_ignore_ascii_case(b))
}

/// Try to identify what kind of file the bytes represent.
fn detect_file_type_hint(bytes: &[u8]) -> String {
    if bytes.is_empty() {
        return "file is empty".to_string();
    }

    let trimmed = strip_bom_and_whitespace(bytes);

    // HTML
    if starts_with_ci(trimmed, b"<!doctype html")
        || starts_with_ci(trimmed, b"<html")
        || starts_with_ci(trimmed, b"<head")
        || starts_with_ci(trimmed, b"<body")
    {
        return "file appears to be HTML".to_string();
    }

    // XML (but not HTML)
    if trimmed.starts_with(b"<?xml") || trimmed.starts_with(b"<") {
        if starts_with_ci(trimmed, b"<?xml") {
            return "file appears to be XML".to_string();
        }
        if trimmed.starts_with(b"<") && !trimmed.starts_with(b"<%") {
            return "file appears to be XML".to_string();
        }
    }

    // JSON
    if trimmed.starts_with(b"{") || trimmed.starts_with(b"[") {
        return "file appears to be JSON".to_string();
    }

    // PNG
    if bytes.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
        return "file appears to be a PNG image".to_string();
    }

    // JPEG
    if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        return "file appears to be a JPEG image".to_string();
    }

    // ZIP / Office documents
    if bytes.starts_with(&[0x50, 0x4B, 0x03, 0x04]) {
        return "file appears to be a ZIP archive (possibly an Office document)".to_string();
    }

    // If it looks like mostly printable ASCII/UTF-8, call it plain text
    let sample = &bytes[..bytes.len().min(512)];
    let printable = sample
        .iter()
        .filter(|&&b| b.is_ascii_graphic() || b.is_ascii_whitespace())
        .count();
    if printable > sample.len() * 3 / 4 {
        return "file appears to be plain text".to_string();
    }

    "file is not a PDF".to_string()
}

/// Validate that a byte buffer looks like a PDF (has `%PDF-` magic).
///
/// Scans the first 1024 bytes, allowing for a UTF-8 BOM and leading whitespace.
pub(crate) fn validate_pdf_bytes(buffer: &[u8]) -> Result<(), PdfError> {
    if buffer.is_empty() {
        return Err(PdfError::NotAPdf(detect_file_type_hint(buffer)));
    }

    let header = &buffer[..buffer.len().min(1024)];
    let trimmed = strip_bom_and_whitespace(header);

    if trimmed.starts_with(b"%PDF-") {
        Ok(())
    } else {
        Err(PdfError::NotAPdf(detect_file_type_hint(buffer)))
    }
}

/// Validate that a file on disk looks like a PDF.
///
/// Reads only the first 1024 bytes and delegates to [`validate_pdf_bytes`].
pub(crate) fn validate_pdf_file<P: AsRef<Path>>(path: P) -> Result<(), PdfError> {
    use std::io::Read;
    let mut file = std::fs::File::open(path)?;
    let mut buf = [0u8; 1024];
    let n = file.read(&mut buf)?;
    validate_pdf_bytes(&buf[..n])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_encoding_issues_fffd() {
        assert!(detect_encoding_issues(
            "Some text with \u{FFFD} replacement"
        ));
    }

    #[test]
    fn test_detect_encoding_issues_dollar_as_space() {
        // Simulates broken CMap: "$Workshop$on$Chest$Wall$Deformities$and$..."
        let garbled = "Last$advanced$Book$Programm$3th$Workshop$on$Chest$Wall$Deformities$and$More";
        assert!(detect_encoding_issues(garbled));
    }

    #[test]
    fn test_detect_encoding_issues_financial_text() {
        // Legitimate dollar signs in financial text should NOT trigger
        let financial = "Revenue was $100M in Q1, up from $90M. Costs: $50M, $30M, $20M, $15M, $12M, $8M, $5M, $3M, $2M, $1M, $500K.";
        assert!(!detect_encoding_issues(financial));
    }

    #[test]
    fn test_detect_encoding_issues_clean_text() {
        assert!(!detect_encoding_issues(
            "Normal markdown text with no issues."
        ));
    }

    #[test]
    fn test_detect_encoding_issues_few_dollars() {
        // Under threshold of 10 total dollars — should not trigger
        let text = "a$b c$d e$f";
        assert!(!detect_encoding_issues(text));
    }
}
