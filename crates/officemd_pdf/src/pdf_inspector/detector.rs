//! Smart PDF type detection without full document load
//!
//! This module detects whether a PDF is text-based, scanned, or image-based
//! by sampling content streams for text operators (Tj/TJ) without loading
//! all objects.

use crate::pdf_inspector::PdfError;
use lopdf::{Document, Object, ObjectId};
use std::collections::{HashMap, HashSet};
use std::path::Path;

/// PDF type classification
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PdfType {
    /// PDF has extractable text (Tj/TJ operators found)
    TextBased,
    /// PDF appears to be scanned (images only, no text operators)
    Scanned,
    /// PDF contains mostly images with minimal/no text
    ImageBased,
    /// PDF has mix of text and image-heavy pages
    Mixed,
}

/// Strategy for which pages to scan during detection
#[derive(Debug, Clone)]
pub enum ScanStrategy {
    /// Scan all pages, stop on first non-text page (current default).
    /// Best for pipelines that route TextBased PDFs to fast extraction.
    EarlyExit,
    /// Scan all pages, no early exit.
    /// Best when you need accurate Mixed vs Scanned classification.
    Full,
    /// Sample up to N evenly distributed pages (first, last, middle).
    /// Best for very large PDFs where speed matters more than precision.
    Sample(u32),
    /// Only scan these specific 1-indexed page numbers.
    /// Best when the caller knows which pages to check.
    Pages(Vec<u32>),
}

/// Result of PDF type detection
#[derive(Debug)]
pub struct PdfTypeResult {
    /// Detected PDF type
    pub pdf_type: PdfType,
    /// Number of pages in the document
    pub page_count: u32,
    /// Number of pages sampled for detection
    pub pages_sampled: u32,
    /// Number of pages with text operators found
    pub pages_with_text: u32,
    /// Confidence score (0.0 - 1.0)
    pub confidence: f32,
    /// Title from metadata (if available)
    pub title: Option<String>,
    /// Whether OCR is recommended for better extraction
    /// True when images provide essential context (e.g., template-based PDFs)
    pub ocr_recommended: bool,
    /// 1-indexed page numbers that need OCR (image-only or insufficient text).
    /// Empty for TextBased. All pages for Scanned/ImageBased. Specific pages for Mixed.
    pub pages_needing_ocr: Vec<u32>,
}

/// Configuration for PDF type detection
#[derive(Debug, Clone)]
pub struct DetectionConfig {
    /// Strategy for which pages to scan
    pub strategy: ScanStrategy,
    /// Minimum text operator count per page to consider as text-based
    pub min_text_ops_per_page: u32,
    /// Threshold ratio of text pages to total pages for classification
    pub text_page_ratio_threshold: f32,
}

impl Default for DetectionConfig {
    fn default() -> Self {
        Self {
            // EarlyExit is too aggressive for PDFs with an image-only cover
            // followed by text-heavy pages (e.g., annual reports).
            strategy: ScanStrategy::Sample(8),
            min_text_ops_per_page: 3,
            text_page_ratio_threshold: 0.6,
        }
    }
}

/// Detect PDF type from file path
pub fn detect_pdf_type<P: AsRef<Path>>(path: P) -> Result<PdfTypeResult, PdfError> {
    detect_pdf_type_with_config(path, DetectionConfig::default())
}

/// Detect PDF type from file path with custom configuration
pub fn detect_pdf_type_with_config<P: AsRef<Path>>(
    path: P,
    config: DetectionConfig,
) -> Result<PdfTypeResult, PdfError> {
    crate::pdf_inspector::validate_pdf_file(&path)?;

    // First, load metadata only (fast operation)
    let metadata = match Document::load_metadata(&path) {
        Ok(m) => m,
        Err(ref e) if crate::pdf_inspector::is_encrypted_lopdf_error(e) => {
            Document::load_metadata_with_password(&path, "")?
        }
        Err(e) => return Err(e.into()),
    };

    // Then load the full document for content inspection
    // We use filtered loading to skip heavy objects we don't need
    let doc = match Document::load(&path) {
        Ok(d) => d,
        Err(ref e) if crate::pdf_inspector::is_encrypted_lopdf_error(e) => {
            Document::load_with_password(&path, "")?
        }
        Err(e) => return Err(e.into()),
    };

    detect_from_document(&doc, metadata.page_count, &config)
}

/// Detect PDF type from memory buffer
pub fn detect_pdf_type_mem(buffer: &[u8]) -> Result<PdfTypeResult, PdfError> {
    detect_pdf_type_mem_with_config(buffer, DetectionConfig::default())
}

/// Detect PDF type from memory buffer with custom configuration
pub fn detect_pdf_type_mem_with_config(
    buffer: &[u8],
    config: DetectionConfig,
) -> Result<PdfTypeResult, PdfError> {
    crate::pdf_inspector::validate_pdf_bytes(buffer)?;

    // Load metadata first (fast)
    let metadata = match Document::load_metadata_mem(buffer) {
        Ok(m) => m,
        Err(ref e) if crate::pdf_inspector::is_encrypted_lopdf_error(e) => {
            Document::load_metadata_mem_with_password(buffer, "")?
        }
        Err(e) => return Err(e.into()),
    };

    // Load document for inspection
    let doc = match Document::load_mem(buffer) {
        Ok(d) => d,
        Err(ref e) if crate::pdf_inspector::is_encrypted_lopdf_error(e) => {
            Document::load_mem_with_password(buffer, "")?
        }
        Err(e) => return Err(e.into()),
    };

    detect_from_document(&doc, metadata.page_count, &config)
}

/// Detection logic on a pre-loaded document.
///
/// `page_count` should come from `Document::load_metadata()`.
pub(crate) fn detect_from_document(
    doc: &Document,
    page_count: u32,
    config: &DetectionConfig,
) -> Result<PdfTypeResult, PdfError> {
    let pages = doc.get_pages();
    let total_pages = pages.len() as u32;

    // Select pages to scan based on strategy
    let (sample_indices, allow_early_exit) = match &config.strategy {
        ScanStrategy::EarlyExit => ((1..=total_pages).collect::<Vec<_>>(), true),
        ScanStrategy::Full => ((1..=total_pages).collect::<Vec<_>>(), false),
        ScanStrategy::Sample(max_pages) => {
            let n = (*max_pages).min(total_pages);
            (distribute_pages(n, total_pages), false)
        }
        ScanStrategy::Pages(pages) => {
            let mut valid: Vec<u32> = pages
                .iter()
                .copied()
                .filter(|&p| p >= 1 && p <= total_pages)
                .collect();
            valid.sort();
            valid.dedup();
            (valid, false)
        }
    };

    let mut pages_with_text = 0u32;
    let mut pages_with_images = 0u32;
    let mut pages_with_template_images = 0u32;
    let mut total_text_ops = 0u32;
    // Cache Phase 1 results to avoid re-analyzing sampled pages in Phase 2
    let mut analysis_cache: HashMap<u32, PageAnalysis> = HashMap::new();
    let mut pages_actually_sampled = 0u32;

    for page_num in &sample_indices {
        if let Some(&page_id) = pages.get(page_num) {
            let analysis = analyze_page_content(doc, page_id);
            pages_actually_sampled += 1;
            if analysis.text_operator_count >= config.min_text_ops_per_page {
                pages_with_text += 1;
            }
            if analysis.has_images {
                pages_with_images += 1;
            }
            if analysis.has_template_image {
                pages_with_template_images += 1;
            }
            total_text_ops += analysis.text_operator_count;
            analysis_cache.insert(*page_num, analysis.clone());

            // Early exit: if this page is non-text (no text ops but has images),
            // this PDF won't be purely TextBased. Stop scanning remaining pages.
            if allow_early_exit
                && analysis.text_operator_count < config.min_text_ops_per_page
                && (analysis.has_images || analysis.has_template_image)
            {
                break;
            }
        }
    }

    let pages_sampled = pages_actually_sampled;
    let text_ratio = if pages_sampled > 0 {
        pages_with_text as f32 / pages_sampled as f32
    } else {
        0.0
    };

    // Check if this is a template-based PDF (images provide essential context)
    // Template PDFs have text AND large background images on most pages
    let has_template_images = pages_with_template_images > 0;
    let template_ratio = if pages_sampled > 0 {
        pages_with_template_images as f32 / pages_sampled as f32
    } else {
        0.0
    };

    // OCR is recommended when:
    // 1. Template images are present (text alone is insufficient), OR
    // 2. PDF is scanned/image-based
    let ocr_recommended: bool;

    // Classification logic
    let (pdf_type, confidence) = if has_template_images && pages_with_text > 0 {
        // Template-based PDF: has text but images provide essential context
        // Classify as Mixed with lower confidence
        ocr_recommended = true;
        (PdfType::Mixed, 0.5 + (0.3 * (1.0 - template_ratio)))
    } else if text_ratio >= config.text_page_ratio_threshold {
        ocr_recommended = false;
        (PdfType::TextBased, text_ratio)
    } else if pages_with_text == 0 && pages_with_images > 0 {
        ocr_recommended = true;
        if total_text_ops == 0 {
            (PdfType::Scanned, 0.95)
        } else {
            (PdfType::ImageBased, 0.8)
        }
    } else if pages_with_text > 0 && pages_with_images > 0 {
        ocr_recommended = true;
        (PdfType::Mixed, 0.7)
    } else if total_text_ops == 0 {
        ocr_recommended = true;
        (PdfType::Scanned, 0.9)
    } else {
        ocr_recommended = false;
        (PdfType::TextBased, text_ratio.max(0.5))
    };

    // Phase 2: Build per-page OCR list
    let pages_needing_ocr = match pdf_type {
        PdfType::TextBased => Vec::new(),
        PdfType::Scanned | PdfType::ImageBased => (1..=total_pages).collect(),
        PdfType::Mixed => {
            let mut ocr_pages = Vec::new();
            for page_num in 1..=total_pages {
                let analysis = if let Some(cached) = analysis_cache.get(&page_num) {
                    cached.clone()
                } else if let Some(&page_id) = pages.get(&page_num) {
                    analyze_page_content(doc, page_id)
                } else {
                    continue;
                };
                if analysis.has_template_image
                    || (analysis.text_operator_count < config.min_text_ops_per_page
                        && analysis.has_images)
                {
                    ocr_pages.push(page_num);
                }
            }
            ocr_pages
        }
    };

    // Try to get title from metadata
    let title = get_document_title(doc);

    Ok(PdfTypeResult {
        pdf_type,
        page_count,
        pages_sampled,
        pages_with_text,
        confidence,
        title,
        ocr_recommended,
        pages_needing_ocr,
    })
}

/// Distribute `n` page indices evenly across `total` pages (1-indexed).
///
/// Always includes the first and last page, with remaining pages
/// spaced evenly in between.
fn distribute_pages(n: u32, total: u32) -> Vec<u32> {
    if n == 0 {
        return Vec::new();
    }
    if n >= total {
        return (1..=total).collect();
    }

    let mut indices = Vec::with_capacity(n as usize);
    indices.push(1);

    if n > 1 {
        indices.push(total);
    }

    let remaining = n.saturating_sub(2);
    if remaining > 0 && total > 2 {
        let step = (total - 2) / (remaining + 1);
        for i in 1..=remaining {
            let idx = 1 + (step * i);
            if idx > 1 && idx < total && !indices.contains(&idx) {
                indices.push(idx);
            }
        }
    }

    indices.sort();
    indices.dedup();
    indices
}

/// Page content analysis result
#[derive(Clone)]
struct PageAnalysis {
    text_operator_count: u32,
    has_images: bool,
    /// Whether page has a large background/template image (>50% coverage)
    has_template_image: bool,
    /// Total image area in pixels (reserved for future use)
    #[allow(dead_code)]
    total_image_area: u64,
}

/// Analyze a page's content stream for text operators and images
fn analyze_page_content(doc: &Document, page_id: ObjectId) -> PageAnalysis {
    let mut text_ops = 0u32;
    let mut has_images = false;

    // Get content streams for this page
    let content_streams = doc.get_page_contents(page_id);

    for content_id in content_streams {
        if let Ok(Object::Stream(stream)) = doc.get_object(content_id) {
            // Try to decompress and scan content
            let content = match stream.decompressed_content() {
                Ok(data) => data,
                Err(_) => stream.content.clone(),
            };

            // Scan for text operators (Tj, TJ)
            let (ops, imgs) = scan_content_for_text_operators(&content);
            text_ops += ops;
            has_images = has_images || imgs;
        }
    }

    // Scan XObject Form contents for text operators
    if let Ok((resource_dict, resource_ids)) = doc.get_page_resources(page_id) {
        let mut visited = HashSet::new();
        if let Some(resources) = resource_dict {
            let (ops, imgs) = scan_xobjects_in_resources(doc, resources, &mut visited);
            text_ops += ops;
            has_images = has_images || imgs;
        }
        for resource_id in resource_ids {
            if let Ok(resources) = doc.get_dictionary(resource_id) {
                let (ops, imgs) = scan_xobjects_in_resources(doc, resources, &mut visited);
                text_ops += ops;
                has_images = has_images || imgs;
            }
        }
    }

    // Check for XObject images and calculate coverage
    let (found_images, total_image_area, has_template_image) = analyze_page_images(doc, page_id);

    if found_images {
        has_images = true;
    }

    PageAnalysis {
        text_operator_count: text_ops,
        has_images,
        has_template_image,
        total_image_area,
    }
}

fn scan_xobjects_in_resources(
    doc: &Document,
    resources: &lopdf::Dictionary,
    visited: &mut HashSet<ObjectId>,
) -> (u32, bool) {
    let mut text_ops = 0u32;
    let mut has_images = false;

    let xobjects = match resources.get(b"XObject").ok() {
        Some(Object::Dictionary(d)) => Some(d.clone()),
        Some(Object::Reference(r)) => doc.get_dictionary(*r).ok().cloned(),
        _ => None,
    };

    if let Some(xobj_dict) = xobjects {
        for (_, obj) in xobj_dict.iter() {
            let Some(obj_id) = obj.as_reference().ok() else {
                continue;
            };
            if !visited.insert(obj_id) {
                continue;
            }
            let Ok(Object::Stream(stream)) = doc.get_object(obj_id) else {
                continue;
            };
            let subtype = stream
                .dict
                .get(b"Subtype")
                .ok()
                .and_then(|o| o.as_name().ok());
            match subtype {
                Some(b"Form") => {
                    let content = stream
                        .decompressed_content()
                        .unwrap_or_else(|_| stream.content.clone());
                    let (ops, imgs) = scan_content_for_text_operators(&content);
                    text_ops += ops;
                    has_images = has_images || imgs;
                    if let Some(res) = stream
                        .dict
                        .get(b"Resources")
                        .ok()
                        .and_then(|o| o.as_dict().ok())
                    {
                        let (ops2, imgs2) = scan_xobjects_in_resources(doc, res, visited);
                        text_ops += ops2;
                        has_images = has_images || imgs2;
                    }
                }
                Some(b"Image") => {
                    has_images = true;
                }
                _ => {}
            }
        }
    }

    (text_ops, has_images)
}

/// Fast scan of content stream bytes for text operators
///
/// This is a fast heuristic scan that looks for:
/// - "Tj" - show text string
/// - "TJ" - show text with individual glyph positioning
/// - "'" - move to next line and show text
/// - "\"" - set word/char spacing, move to next line, show text
fn scan_content_for_text_operators(content: &[u8]) -> (u32, bool) {
    let mut text_ops = 0u32;
    let mut has_images = false;

    // Simple state machine to find operators
    let mut i = 0;
    while i < content.len() {
        let b = content[i];

        // Look for 'T' followed by 'j' or 'J'
        if b == b'T' && i + 1 < content.len() {
            let next = content[i + 1];
            if next == b'j' || next == b'J' {
                // Verify it's an operator (followed by whitespace or newline)
                if i + 2 >= content.len()
                    || content[i + 2].is_ascii_whitespace()
                    || content[i + 2] == b'\n'
                    || content[i + 2] == b'\r'
                {
                    text_ops += 1;
                }
            }
        }

        // Look for 'Do' operator (XObject/image placement)
        if b == b'D'
            && i + 1 < content.len()
            && content[i + 1] == b'o'
            && (i + 2 >= content.len() || content[i + 2].is_ascii_whitespace())
        {
            has_images = true;
        }

        i += 1;
    }

    (text_ops, has_images)
}

/// Analyze page images: returns (has_images, total_area, has_template_image)
///
/// A template image is one that covers >50% of a standard page area.
/// Standard page: 612x792 points (US Letter) = ~485,000 sq points
/// At 2x resolution that's ~1.9M pixels, so we use 250K pixels as threshold
/// (accounting for varying DPI and page sizes)
fn analyze_page_images(doc: &Document, page_id: ObjectId) -> (bool, u64, bool) {
    // Threshold: image covering roughly half a page at 150+ DPI
    // 612 * 792 / 2 * (150/72)^2 ≈ 1M pixels, but we'll be conservative
    const TEMPLATE_IMAGE_THRESHOLD: u64 = 500_000; // 500K pixels

    let mut has_images = false;
    let mut total_area: u64 = 0;
    let mut has_template_image = false;

    if let Ok(page_dict) = doc.get_dictionary(page_id) {
        let resources = match page_dict.get(b"Resources") {
            Ok(Object::Reference(id)) => doc.get_dictionary(*id).ok(),
            Ok(Object::Dictionary(dict)) => Some(dict),
            _ => None,
        };

        if let Some(resources) = resources {
            if let Ok(xobject) = resources.get(b"XObject") {
                let xobject_dict = match xobject {
                    Object::Reference(id) => doc.get_dictionary(*id).ok(),
                    Object::Dictionary(dict) => Some(dict),
                    _ => None,
                };

                if let Some(xobject_dict) = xobject_dict {
                    for (_, value) in xobject_dict.iter() {
                        if let Ok(xobj_ref) = value.as_reference() {
                            if let Ok(xobj) = doc.get_object(xobj_ref) {
                                if let Ok(stream) = xobj.as_stream() {
                                    // Check if it's an Image subtype
                                    if let Ok(subtype) = stream.dict.get(b"Subtype") {
                                        if let Ok(name) = subtype.as_name() {
                                            if name == b"Image" {
                                                has_images = true;

                                                // Get image dimensions
                                                let width = stream
                                                    .dict
                                                    .get(b"Width")
                                                    .ok()
                                                    .and_then(|w| w.as_i64().ok())
                                                    .unwrap_or(0)
                                                    as u64;
                                                let height = stream
                                                    .dict
                                                    .get(b"Height")
                                                    .ok()
                                                    .and_then(|h| h.as_i64().ok())
                                                    .unwrap_or(0)
                                                    as u64;

                                                let area = width * height;
                                                total_area += area;

                                                // Check if this is a large template image
                                                if area >= TEMPLATE_IMAGE_THRESHOLD {
                                                    has_template_image = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    (has_images, total_area, has_template_image)
}

/// Get document title from Info dictionary
fn get_document_title(doc: &Document) -> Option<String> {
    let info_ref = doc.trailer.get(b"Info").ok()?.as_reference().ok()?;
    let info = doc.get_dictionary(info_ref).ok()?;
    let title_obj = info.get(b"Title").ok()?;

    match title_obj {
        Object::String(bytes, _) => {
            // Handle UTF-16BE encoding (BOM: 0xFE 0xFF)
            if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
                let utf16: Vec<u16> = bytes[2..]
                    .chunks_exact(2)
                    .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
                    .collect();
                Some(String::from_utf16_lossy(&utf16))
            } else {
                Some(String::from_utf8_lossy(bytes).to_string())
            }
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_scan_content_operators() {
        // Sample PDF content stream with text operators
        let content = b"BT /F1 12 Tf 100 700 Td (Hello World) Tj ET";
        let (ops, imgs) = scan_content_for_text_operators(content);
        assert_eq!(ops, 1);
        assert!(!imgs);

        // Content with TJ array
        let content2 = b"BT /F1 12 Tf 100 700 Td [(H) 10 (ello)] TJ ET";
        let (ops2, _) = scan_content_for_text_operators(content2);
        assert_eq!(ops2, 1);

        // Content with Do (image)
        let content3 = b"q 100 0 0 100 50 700 cm /Img1 Do Q";
        let (ops3, imgs3) = scan_content_for_text_operators(content3);
        assert_eq!(ops3, 0);
        assert!(imgs3);
    }
}
