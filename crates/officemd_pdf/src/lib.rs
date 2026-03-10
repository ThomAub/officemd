//! PDF extraction wrapper for officemd IR and markdown rendering.

#[allow(warnings)]
mod pdf_inspector;

use crate::pdf_inspector::{
    MarkdownOptions, PdfOptions, PdfProcessResult, PdfType, process_pdf_mem_with_options,
};
use lopdf::{Dictionary, Document as LopdfDocument};
use officemd_core::ir::{
    DocumentKind, DocumentProperties, OoxmlDocument, PdfClassification, PdfDiagnostics,
    PdfDocument, PdfPage,
};
use officemd_markdown::{RenderOptions, render_document_with_options};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet, HashMap};

#[derive(Debug, thiserror::Error)]
pub enum OoxmlPdfError {
    #[error(transparent)]
    Pdf(#[from] crate::pdf_inspector::PdfError),
    #[error(transparent)]
    Lopdf(#[from] lopdf::Error),
    #[error("JSON serialization error: {0}")]
    Json(#[from] serde_json::Error),
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PdfFontUsage {
    pub font_name: String,
    pub raw_font_names: Vec<String>,
    pub resource_names: Vec<String>,
    pub pages: Vec<usize>,
    pub text_item_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PdfFontInspection {
    pub diagnostics: PdfDiagnostics,
    pub fonts: Vec<PdfFontUsage>,
}

/// Returns true when bytes look like a PDF file header.
#[must_use]
pub fn looks_like_pdf_header(content: &[u8]) -> bool {
    if content.is_empty() {
        return false;
    }

    let header = &content[..content.len().min(1024)];
    let trimmed = strip_bom_and_whitespace(header);
    trimmed.starts_with(b"%PDF-")
}

/// Detects parseability/classification metadata for a PDF.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when the content is not a valid PDF.
pub fn inspect_pdf(content: &[u8]) -> Result<PdfDiagnostics, OoxmlPdfError> {
    let result = process_pdf_mem_with_options(content, default_pdf_options(false))?;
    Ok(map_diagnostics(&result))
}

/// Detects fonts declared/used in a PDF with diagnostics.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when the content is not a valid PDF, or
/// `OoxmlPdfError::Lopdf` when the low-level PDF structure cannot be parsed.
pub fn inspect_pdf_fonts(content: &[u8]) -> Result<PdfFontInspection, OoxmlPdfError> {
    let result = process_pdf_mem_with_options(content, default_pdf_options(false))?;
    let diagnostics = map_diagnostics(&result);

    let doc = load_lopdf_document(content)?;
    let mut page_font_lookup: HashMap<usize, HashMap<String, String>> = HashMap::new();
    let mut fonts: BTreeMap<String, FontAccumulator> = BTreeMap::new();

    for (page_number, page_id) in doc.get_pages() {
        let page_number = usize::try_from(page_number).unwrap_or(usize::MAX);
        let page_fonts = doc.get_page_fonts(page_id).unwrap_or_default();
        let lookup = page_font_lookup.entry(page_number).or_default();

        for (resource_name, font_dict) in page_fonts {
            let resource_name = String::from_utf8_lossy(&resource_name).to_string();
            let raw_font_name =
                extract_base_font_name(font_dict).unwrap_or_else(|| resource_name.clone());
            let normalized_name = normalize_font_name(&raw_font_name);

            lookup.insert(resource_name.clone(), normalized_name.clone());

            let entry = fonts.entry(normalized_name).or_default();
            entry.raw_font_names.insert(raw_font_name);
            entry.resource_names.insert(resource_name);
            entry.pages.insert(page_number);
        }
    }

    if let Ok(items) = crate::pdf_inspector::extractor::extract_text_with_positions_mem(content) {
        for item in items {
            let page_number = usize::try_from(item.page).unwrap_or(usize::MAX);
            let lookup_name = page_font_lookup
                .get(&page_number)
                .and_then(|lookup| lookup.get(&item.font));
            let entry = match lookup_name {
                Some(font_name) => {
                    if let Some(entry) = fonts.get_mut(font_name) {
                        entry
                    } else {
                        fonts.entry(font_name.clone()).or_default()
                    }
                }
                None => {
                    let normalized = normalize_font_name(&item.font);
                    fonts.entry(normalized).or_default()
                }
            };
            entry.resource_names.insert(item.font);
            entry.pages.insert(page_number);
            if !item.text.trim().is_empty() {
                entry.text_item_count += 1;
            }
        }
    }

    let mut fonts = fonts
        .into_iter()
        .map(|(font_name, accumulator)| PdfFontUsage {
            font_name,
            raw_font_names: accumulator.raw_font_names.into_iter().collect(),
            resource_names: accumulator.resource_names.into_iter().collect(),
            pages: accumulator.pages.into_iter().collect(),
            text_item_count: accumulator.text_item_count,
        })
        .collect::<Vec<_>>();

    fonts.sort_by(|a, b| {
        b.text_item_count
            .cmp(&a.text_item_count)
            .then_with(|| a.font_name.cmp(&b.font_name))
    });

    Ok(PdfFontInspection { diagnostics, fonts })
}

/// Detects PDF fonts and returns JSON.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` or `OoxmlPdfError::Lopdf` when font inspection
/// fails, or `OoxmlPdfError::Json` when serialization fails.
pub fn inspect_pdf_fonts_json(content: &[u8]) -> Result<String, OoxmlPdfError> {
    let inspection = inspect_pdf_fonts(content)?;
    Ok(serde_json::to_string(&inspection)?)
}

/// Extract PDF content as the shared officemd IR.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when the content is not a valid PDF or
/// extraction fails.
pub fn extract_ir(content: &[u8]) -> Result<OoxmlDocument, OoxmlPdfError> {
    extract_ir_force(content, false)
}

/// Extract PDF content as the shared officemd IR, optionally forcing
/// extraction on scanned/image-based PDFs.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when the content is not a valid PDF or
/// extraction fails.
pub fn extract_ir_force(content: &[u8], force_extract: bool) -> Result<OoxmlDocument, OoxmlPdfError> {
    let options = default_pdf_options(force_extract);
    let result = process_pdf_mem_with_options(content, options)?;

    let diagnostics = map_diagnostics(&result);
    let pages = result
        .markdown
        .as_deref()
        .map(split_markdown_into_pages)
        .map(|pages| fill_missing_pages(pages, diagnostics.page_count))
        .unwrap_or_default();

    let properties = Some(build_document_properties(&result, &diagnostics));

    Ok(OoxmlDocument {
        kind: DocumentKind::Pdf,
        properties,
        sheets: vec![],
        slides: vec![],
        sections: vec![],
        pdf: Some(PdfDocument { pages, diagnostics }),
    })
}

/// Extract PDF content as IR JSON.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when extraction fails, or
/// `OoxmlPdfError::Json` when serialization fails.
pub fn extract_ir_json(content: &[u8]) -> Result<String, OoxmlPdfError> {
    let doc = extract_ir(content)?;
    Ok(serde_json::to_string(&doc)?)
}

/// Extract PDF content as IR JSON, optionally forcing extraction on
/// scanned/image-based PDFs.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when extraction fails, or
/// `OoxmlPdfError::Json` when serialization fails.
pub fn extract_ir_json_force(content: &[u8], force_extract: bool) -> Result<String, OoxmlPdfError> {
    let doc = extract_ir_force(content, force_extract)?;
    Ok(serde_json::to_string(&doc)?)
}

/// Render PDF bytes directly to markdown with shared render options.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when the content is not a valid PDF or
/// extraction fails.
pub fn markdown_from_bytes_with_options(
    content: &[u8],
    render: RenderOptions,
) -> Result<String, OoxmlPdfError> {
    markdown_from_bytes_force(content, render, false)
}

/// Render PDF bytes directly to markdown, optionally forcing extraction on
/// scanned/image-based PDFs.
///
/// # Errors
///
/// Returns `OoxmlPdfError::Pdf` when the content is not a valid PDF or
/// extraction fails.
pub fn markdown_from_bytes_force(
    content: &[u8],
    render: RenderOptions,
    force_extract: bool,
) -> Result<String, OoxmlPdfError> {
    let doc = extract_ir_force(content, force_extract)?;
    Ok(render_document_with_options(&doc, render))
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

fn default_pdf_options(force_extract: bool) -> PdfOptions {
    PdfOptions::new()
        .markdown(MarkdownOptions {
            include_page_numbers: true,
            ..Default::default()
        })
        .force_extract(force_extract)
}

#[derive(Debug, Default)]
struct FontAccumulator {
    raw_font_names: BTreeSet<String>,
    resource_names: BTreeSet<String>,
    pages: BTreeSet<usize>,
    text_item_count: usize,
}

fn load_lopdf_document(content: &[u8]) -> Result<LopdfDocument, OoxmlPdfError> {
    match LopdfDocument::load_mem(content) {
        Ok(doc) => Ok(doc),
        Err(_) => Ok(LopdfDocument::load_mem_with_password(content, "")?),
    }
}

fn extract_base_font_name(font_dict: &Dictionary) -> Option<String> {
    font_dict
        .get(b"BaseFont")
        .ok()
        .and_then(|value| value.as_name().ok())
        .map(|name| String::from_utf8_lossy(name).to_string())
}

fn normalize_font_name(raw: &str) -> String {
    let trimmed = raw.trim().trim_start_matches('/');
    let mut chars = trimmed.chars();
    let mut prefix = String::new();

    for _ in 0..6 {
        let Some(ch) = chars.next() else {
            return trimmed.to_string();
        };
        prefix.push(ch);
    }

    if chars.next() == Some('+') && prefix.chars().all(|ch| ch.is_ascii_uppercase()) {
        chars.collect()
    } else {
        trimmed.to_string()
    }
}

fn map_diagnostics(result: &PdfProcessResult) -> PdfDiagnostics {
    PdfDiagnostics {
        classification: map_classification(result.pdf_type),
        confidence: result.confidence,
        page_count: usize::try_from(result.page_count).unwrap_or(usize::MAX),
        pages_needing_ocr: result
            .pages_needing_ocr
            .iter()
            .map(|page| usize::try_from(*page).unwrap_or(usize::MAX))
            .collect(),
        has_encoding_issues: result.has_encoding_issues,
    }
}

fn map_classification(value: PdfType) -> PdfClassification {
    match value {
        PdfType::TextBased => PdfClassification::TextBased,
        PdfType::Scanned => PdfClassification::Scanned,
        PdfType::ImageBased => PdfClassification::ImageBased,
        PdfType::Mixed => PdfClassification::Mixed,
    }
}

fn build_document_properties(
    result: &PdfProcessResult,
    diagnostics: &PdfDiagnostics,
) -> DocumentProperties {
    let mut core = HashMap::new();
    if let Some(title) = &result.title {
        core.insert("title".to_string(), title.clone());
    }

    core.insert(
        "classification".to_string(),
        format!("{:?}", diagnostics.classification),
    );
    core.insert(
        "confidence".to_string(),
        format!("{:.4}", diagnostics.confidence),
    );
    core.insert("page_count".to_string(), diagnostics.page_count.to_string());
    core.insert(
        "pages_needing_ocr".to_string(),
        if diagnostics.pages_needing_ocr.is_empty() {
            String::new()
        } else {
            diagnostics
                .pages_needing_ocr
                .iter()
                .map(usize::to_string)
                .collect::<Vec<_>>()
                .join(",")
        },
    );
    core.insert(
        "has_encoding_issues".to_string(),
        diagnostics.has_encoding_issues.to_string(),
    );

    DocumentProperties {
        core,
        app: HashMap::new(),
        custom: HashMap::new(),
    }
}

fn split_markdown_into_pages(markdown: &str) -> Vec<PdfPage> {
    let mut pages = Vec::new();
    let mut current_page = 1usize;
    let mut current = String::new();
    let mut saw_page_marker = false;

    for line in markdown.lines() {
        if let Some(page_num) = parse_page_marker(line) {
            if saw_page_marker || !current.trim().is_empty() {
                flush_page(&mut pages, current_page, &mut current, true);
            }
            current_page = page_num;
            saw_page_marker = true;
            continue;
        }

        if !current.is_empty() {
            current.push('\n');
        }
        current.push_str(line);
    }

    flush_page(&mut pages, current_page, &mut current, saw_page_marker);
    pages
}

fn fill_missing_pages(pages: Vec<PdfPage>, page_count: usize) -> Vec<PdfPage> {
    if pages.is_empty() {
        return pages;
    }

    let has_explicit_page_numbers = pages.iter().any(|page| page.number > 1);
    if !has_explicit_page_numbers || page_count == 0 {
        return pages;
    }

    let mut page_map: BTreeMap<usize, String> = BTreeMap::new();
    for page in pages {
        page_map.entry(page.number).or_insert(page.markdown);
    }

    let mut normalized = Vec::with_capacity(page_count.max(page_map.len()));
    for number in 1..=page_count {
        normalized.push(PdfPage {
            number,
            markdown: page_map.remove(&number).unwrap_or_default(),
        });
    }

    for (number, markdown) in page_map {
        normalized.push(PdfPage { number, markdown });
    }

    normalized
}

fn parse_page_marker(line: &str) -> Option<usize> {
    let trimmed = line.trim();
    if !trimmed.starts_with("<!--") || !trimmed.ends_with("-->") {
        return None;
    }

    let inner = trimmed.strip_prefix("<!--")?.strip_suffix("-->")?.trim();
    let page = inner.strip_prefix("Page ")?;
    page.parse::<usize>().ok()
}

fn flush_page(
    pages: &mut Vec<PdfPage>,
    page_number: usize,
    current: &mut String,
    include_empty: bool,
) {
    let body = current.trim();
    if include_empty || !body.is_empty() {
        pages.push(PdfPage {
            number: page_number,
            markdown: body.to_string(),
        });
    }
    current.clear();
}

#[cfg(test)]
mod tests {
    use super::*;
    use lopdf::{Document, Object, Stream, dictionary};

    static TEXT_FIXTURE: &[u8] = include_bytes!("../../../examples/data/OpenXML_WhitePaper.pdf");
    static ENCODING_HEURISTIC_FIXTURE: &[u8] =
        include_bytes!("../../../examples/data/encoding_heuristic_fixture.pdf");
    static OCR_SCANNED_FIXTURE: &[u8] =
        include_bytes!("../../../examples/data/ocr_graph_scanned.pdf");
    static OCR_OCRED_FIXTURE: &[u8] = include_bytes!("../../../examples/data/ocr_graph_ocred.pdf");
    static OCR_TEXTBASED_FIXTURE: &[u8] =
        include_bytes!("../../../examples/data/ocr_tagged_textbased.pdf");

    fn blank_pdf() -> Vec<u8> {
        let mut doc = Document::with_version("1.5");
        let pages_id = doc.new_object_id();
        let content_id = doc.add_object(Stream::new(dictionary! {}, Vec::new()));
        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "Contents" => content_id,
            "MediaBox" => vec![0.into(), 0.into(), 595.into(), 842.into()],
        });
        let pages = dictionary! {
            "Type" => "Pages",
            "Kids" => vec![page_id.into()],
            "Count" => 1,
        };
        doc.objects.insert(pages_id, Object::Dictionary(pages));
        let catalog_id = doc.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        doc.trailer.set("Root", catalog_id);
        let mut bytes = Vec::new();
        doc.save_to(&mut bytes).expect("save blank pdf");
        bytes
    }

    #[test]
    fn looks_like_pdf_detects_header() {
        assert!(looks_like_pdf_header(b"%PDF-1.7\n"));
        assert!(looks_like_pdf_header(b"\xEF\xBB\xBF   %PDF-1.4\n"));
        assert!(!looks_like_pdf_header(b"not a pdf"));
    }

    #[test]
    fn extract_ir_populates_pdf_for_text_fixture() {
        let doc = extract_ir(TEXT_FIXTURE).expect("extract pdf");
        assert_eq!(doc.kind, DocumentKind::Pdf);
        assert!(doc.pdf.is_some());

        let pdf = doc.pdf.expect("pdf payload");
        assert!(pdf.diagnostics.page_count > 0);
        assert!(!pdf.pages.is_empty());
    }

    #[test]
    fn extract_ir_handles_scanned_like_pdf_without_error() {
        let bytes = blank_pdf();
        let doc = extract_ir(&bytes).expect("extract scanned-like pdf");
        let pdf = doc.pdf.expect("pdf payload");

        assert!(pdf.pages.is_empty());
        assert!(matches!(
            pdf.diagnostics.classification,
            PdfClassification::Scanned | PdfClassification::ImageBased | PdfClassification::Mixed
        ));
        assert!(pdf.diagnostics.page_count >= 1);
    }

    #[test]
    fn split_markdown_preserves_empty_pages_when_markers_are_present() {
        let markdown = "<!-- Page 1 -->\nFirst\n<!-- Page 2 -->\n<!-- Page 3 -->\nThird";
        let pages = split_markdown_into_pages(markdown);

        assert_eq!(pages.len(), 3);
        assert_eq!(pages[0].number, 1);
        assert_eq!(pages[0].markdown, "First");
        assert_eq!(pages[1].number, 2);
        assert_eq!(pages[1].markdown, "");
        assert_eq!(pages[2].number, 3);
        assert_eq!(pages[2].markdown, "Third");
    }

    #[test]
    fn fill_missing_pages_uses_page_count_for_numbered_markdown() {
        let pages = vec![
            PdfPage {
                number: 1,
                markdown: "one".to_string(),
            },
            PdfPage {
                number: 3,
                markdown: "three".to_string(),
            },
        ];

        let normalized = fill_missing_pages(pages, 3);

        assert_eq!(normalized.len(), 3);
        assert_eq!(normalized[0].number, 1);
        assert_eq!(normalized[0].markdown, "one");
        assert_eq!(normalized[1].number, 2);
        assert_eq!(normalized[1].markdown, "");
        assert_eq!(normalized[2].number, 3);
        assert_eq!(normalized[2].markdown, "three");
    }

    #[test]
    fn inspect_pdf_fails_fast_on_non_pdf_input() {
        let err = inspect_pdf(b"<html>not a pdf</html>").expect_err("expected error");
        let msg = err.to_string().to_lowercase();
        assert!(msg.contains("not a pdf") || msg.contains("html"));
    }

    #[test]
    fn inspect_pdf_fonts_reports_fonts_for_text_fixture() {
        let inspection = inspect_pdf_fonts(TEXT_FIXTURE).expect("inspect fonts");
        assert!(inspection.diagnostics.page_count > 0);
        assert!(!inspection.fonts.is_empty());
        assert!(inspection.fonts.iter().any(|font| font.text_item_count > 0));
    }

    #[test]
    fn inspect_pdf_diagnostics_match_extract_ir_diagnostics() {
        let inspect = inspect_pdf(TEXT_FIXTURE).expect("inspect");
        let extracted = extract_ir(TEXT_FIXTURE).expect("extract");
        let pdf = extracted.pdf.expect("pdf payload");

        assert_eq!(inspect.classification, pdf.diagnostics.classification);
        assert_eq!(inspect.page_count, pdf.diagnostics.page_count);
        assert_eq!(inspect.pages_needing_ocr, pdf.diagnostics.pages_needing_ocr);
        assert_eq!(
            inspect.has_encoding_issues,
            pdf.diagnostics.has_encoding_issues
        );
        assert!((inspect.confidence - pdf.diagnostics.confidence).abs() < 1e-6);
    }

    #[test]
    fn inspect_pdf_flags_encoding_issues_for_heuristic_fixture() {
        let diagnostics = inspect_pdf(ENCODING_HEURISTIC_FIXTURE).expect("inspect");
        assert!(diagnostics.has_encoding_issues);

        let extracted = extract_ir(ENCODING_HEURISTIC_FIXTURE).expect("extract");
        let pdf = extracted.pdf.expect("pdf payload");
        assert!(pdf.diagnostics.has_encoding_issues);
    }

    #[test]
    fn inspect_pdf_detects_ocr_gap_for_scanned_vs_ocred_fixture_pair() {
        let scanned = inspect_pdf(OCR_SCANNED_FIXTURE).expect("inspect scanned fixture");
        let ocred = inspect_pdf(OCR_OCRED_FIXTURE).expect("inspect ocred fixture");

        assert!(scanned.page_count >= 1);
        assert!(ocred.page_count >= 1);
        assert!(
            !scanned.pages_needing_ocr.is_empty(),
            "expected scanned fixture to need OCR"
        );
        assert!(
            ocred.pages_needing_ocr.is_empty(),
            "expected OCRed fixture to not need OCR"
        );
        assert_eq!(ocred.classification, PdfClassification::TextBased);
    }

    #[test]
    fn textbased_fixture_has_markdown_and_no_ocr_pages() {
        let diagnostics = inspect_pdf(OCR_TEXTBASED_FIXTURE).expect("inspect textbased fixture");
        assert_eq!(diagnostics.classification, PdfClassification::TextBased);
        assert!(diagnostics.pages_needing_ocr.is_empty());

        let markdown =
            markdown_from_bytes_with_options(OCR_TEXTBASED_FIXTURE, RenderOptions::default())
                .expect("extract markdown");
        assert!(!markdown.trim().is_empty());
        assert!(markdown.contains("## Page: 1"));
    }

    #[test]
    fn inspect_pdf_fonts_json_is_serializable() {
        let payload = inspect_pdf_fonts_json(TEXT_FIXTURE).expect("inspect fonts json");
        let value: serde_json::Value = serde_json::from_str(&payload).expect("valid json");
        assert!(value.get("fonts").is_some());
    }

    #[test]
    fn inspect_pdf_fonts_fails_fast_on_non_pdf_input() {
        let err = inspect_pdf_fonts(b"not a pdf").expect_err("expected error");
        let msg = err.to_string().to_lowercase();
        assert!(msg.contains("not a pdf") || msg.contains("pdf"));
    }

    #[test]
    fn normalize_font_name_removes_subset_prefix() {
        assert_eq!(normalize_font_name("/ABCDEE+Calibri"), "Calibri");
        assert_eq!(
            normalize_font_name("TimesNewRomanPSMT"),
            "TimesNewRomanPSMT"
        );
    }
}
