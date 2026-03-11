//! PDF page content-stream extraction wrapper.

use crate::pdf_inspector::PdfError;
use crate::pdf_inspector::tounicode::FontCMaps;
use crate::pdf_inspector::types::{PdfRect, TextItem};
use lopdf::{Document, ObjectId, content::Content};

use super::fonts::CMapDecisionCache;
use super::interpreter::{MarkedContentMode, build_font_context, extract_content_stream};
use super::xobjects::{extract_form_xobject_content_with_depth, get_page_xobjects};

pub(crate) fn extract_page_text_items(
    doc: &Document,
    page_id: ObjectId,
    page_num: u32,
    font_cmaps: &FontCMaps,
) -> Result<(Vec<TextItem>, Vec<PdfRect>), PdfError> {
    let fonts = doc.get_page_fonts(page_id).unwrap_or_default();
    let font_ctx = build_font_context(doc, &fonts);
    let xobjects = get_page_xobjects(doc, page_id);
    let mut cmap_decisions = CMapDecisionCache::new();

    let content_data = doc
        .get_page_content(page_id)
        .map_err(|e| PdfError::Parse(e.to_string()))?;
    let content = Content::decode(&content_data).map_err(|e| PdfError::Parse(e.to_string()))?;

    let sink = extract_content_stream(
        doc,
        &content.operations,
        page_num,
        font_cmaps,
        &font_ctx,
        &xobjects,
        [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
        MarkedContentMode::ActualText,
        &mut cmap_decisions,
        |form_id, parent_ctm, decisions| {
            extract_form_xobject_content_with_depth(
                doc, form_id, page_num, font_cmaps, parent_ctm, decisions, 0,
            )
        },
    );

    let items = super::merge_text_items(super::deduplicate_overlapping_text_items(sink.items));
    Ok((items, sink.rects))
}
