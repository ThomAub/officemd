//! PDF content-stream operator state machine.
//!
//! Walks the page's content stream, tracking the graphics state and text
//! matrix, and emits `TextItem`s and `PdfRect`s.

use crate::pdf_inspector::PdfError;
use crate::pdf_inspector::text_utils::{
    decode_text_string, effective_font_size, expand_ligatures, is_bold_font, is_italic_font,
};
use crate::pdf_inspector::tounicode::FontCMaps;
use crate::pdf_inspector::types::{ItemType, PdfRect, TextItem};
use log::trace;
use lopdf::{Document, Encoding, Object, ObjectId};
use std::collections::HashMap;

use super::fonts::{
    CMapDecisionCache, build_font_encodings, build_font_widths, compute_string_width_ts,
    extract_text_from_operand, get_font_file2_obj_num, get_operand_bytes,
};
use super::xobjects::{XObjectType, extract_form_xobject_text, get_page_xobjects};
use super::{get_number, multiply_matrices};

pub(crate) fn extract_page_text_items(
    doc: &Document,
    page_id: ObjectId,
    page_num: u32,
    font_cmaps: &FontCMaps,
) -> Result<(Vec<TextItem>, Vec<PdfRect>), PdfError> {
    use lopdf::content::Content;

    let mut items = Vec::new();
    let mut rects: Vec<PdfRect> = Vec::new();

    // Get fonts for encoding
    let fonts = doc.get_page_fonts(page_id).unwrap_or_default();

    // Build font encoding maps from Differences arrays
    let font_encodings = build_font_encodings(doc, &fonts);

    // Build font width info for accurate text positioning
    let font_widths = build_font_widths(doc, &fonts);

    // Build maps of font resource names to their base font names and ToUnicode object refs
    let mut font_base_names: std::collections::HashMap<String, String> =
        std::collections::HashMap::new();
    let mut font_tounicode_refs: std::collections::HashMap<String, u32> =
        std::collections::HashMap::new();
    let mut inline_cmaps: std::collections::HashMap<
        String,
        crate::pdf_inspector::tounicode::CMapEntry,
    > = std::collections::HashMap::new();
    for (font_name, font_dict) in &fonts {
        let resource_name = String::from_utf8_lossy(font_name).to_string();
        if let Ok(base_font) = font_dict.get(b"BaseFont") {
            if let Ok(name) = base_font.as_name() {
                let base_name = String::from_utf8_lossy(name).to_string();
                font_base_names.insert(resource_name.clone(), base_name);
            }
        }
        // Track ToUnicode object reference, with FontFile2 fallback for Identity-H/V.
        // Also handle inline ToUnicode streams.
        match font_dict.get(b"ToUnicode") {
            Ok(tounicode) => {
                if let Ok(obj_ref) = tounicode.as_reference() {
                    font_tounicode_refs.insert(resource_name, obj_ref.0);
                } else if let Object::Stream(s) = tounicode {
                    if let Ok(data) = s.decompressed_content() {
                        if let Some(entry) =
                            crate::pdf_inspector::tounicode::build_cmap_entry_from_stream(
                                &data, font_dict, doc, 0,
                            )
                        {
                            inline_cmaps.insert(resource_name, entry);
                        }
                    }
                }
            }
            Err(_) => {
                if let Some(ff2_obj_num) = get_font_file2_obj_num(doc, font_dict) {
                    font_tounicode_refs.insert(resource_name, ff2_obj_num);
                }
            }
        }
    }

    // Cache font encodings from lopdf (once per font, not per text operand).
    // This avoids re-parsing ToUnicode CMap streams for every Tj/TJ operator.
    let mut encoding_cache: HashMap<String, Encoding<'_>> = HashMap::new();
    for (font_name, font_dict) in &fonts {
        let name = String::from_utf8_lossy(font_name).to_string();
        if let Ok(enc) = font_dict.get_font_encoding(doc) {
            encoding_cache.insert(name, enc);
        }
    }

    let mut cmap_decisions = CMapDecisionCache::new();

    // Get XObjects (images) from page resources
    let xobjects = get_page_xobjects(doc, page_id);

    // Get content
    let content_data = doc
        .get_page_content(page_id)
        .map_err(|e| PdfError::Parse(e.to_string()))?;

    let content = Content::decode(&content_data).map_err(|e| PdfError::Parse(e.to_string()))?;

    // Graphics state tracking
    let mut ctm = [1.0f32, 0.0, 0.0, 1.0, 0.0, 0.0]; // Current Transformation Matrix
    let mut fill_is_white = false; // Fill color is white (invisible text)
    let mut text_rendering_mode: i32 = 0; // 0=fill, 1=stroke, 2=fill+stroke, 3=invisible
    let mut gstate_stack: Vec<([f32; 6], bool, i32)> = Vec::new();

    // Text state tracking
    let mut current_font = String::new();
    let mut current_font_size: f32 = 12.0;
    let mut text_leading: f32 = 0.0; // TL parameter (in text-space units)
    let mut text_matrix = [1.0f32, 0.0, 0.0, 1.0, 0.0, 0.0];
    let mut line_matrix = [1.0f32, 0.0, 0.0, 1.0, 0.0, 0.0];
    let mut in_text_block = false;

    // Marked content (ActualText) tracking
    let mut marked_content_stack: Vec<Option<String>> = Vec::new();
    let mut suppress_glyph_extraction = false;
    let mut actual_text_start_tm: Option<[f32; 6]> = None; // text matrix at BDC entry

    for op in &content.operations {
        trace!("{} {:?}", op.operator, op.operands);
        match op.operator.as_str() {
            "q" => {
                // Save graphics state
                gstate_stack.push((ctm, fill_is_white, text_rendering_mode));
            }
            "Q" => {
                // Restore graphics state
                if let Some((saved_ctm, saved_fill, saved_tr)) = gstate_stack.pop() {
                    ctm = saved_ctm;
                    fill_is_white = saved_fill;
                    text_rendering_mode = saved_tr;
                }
            }
            "cm" => {
                // Concatenate matrix to CTM
                if op.operands.len() >= 6 {
                    let new_matrix = [
                        get_number(&op.operands[0]).unwrap_or(1.0),
                        get_number(&op.operands[1]).unwrap_or(0.0),
                        get_number(&op.operands[2]).unwrap_or(0.0),
                        get_number(&op.operands[3]).unwrap_or(1.0),
                        get_number(&op.operands[4]).unwrap_or(0.0),
                        get_number(&op.operands[5]).unwrap_or(0.0),
                    ];
                    ctm = multiply_matrices(&new_matrix, &ctm);
                }
            }
            "g" => {
                // Set grayscale fill color (1.0 = white)
                if let Some(gray) = op.operands.first().and_then(get_number) {
                    fill_is_white = gray > 0.95;
                }
            }
            "rg" => {
                // Set RGB fill color
                if op.operands.len() >= 3 {
                    let r = get_number(&op.operands[0]).unwrap_or(0.0);
                    let g = get_number(&op.operands[1]).unwrap_or(0.0);
                    let b = get_number(&op.operands[2]).unwrap_or(0.0);
                    fill_is_white = r > 0.95 && g > 0.95 && b > 0.95;
                }
            }
            "k" => {
                // Set CMYK fill color (0,0,0,0 = white)
                if op.operands.len() >= 4 {
                    let c = get_number(&op.operands[0]).unwrap_or(1.0);
                    let m = get_number(&op.operands[1]).unwrap_or(1.0);
                    let y = get_number(&op.operands[2]).unwrap_or(1.0);
                    let k = get_number(&op.operands[3]).unwrap_or(1.0);
                    fill_is_white = c < 0.05 && m < 0.05 && y < 0.05 && k < 0.05;
                }
            }
            "BT" => {
                // Begin text block
                in_text_block = true;
                text_matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
                line_matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
                text_rendering_mode = 0;
            }
            "ET" => {
                // End text block
                in_text_block = false;
            }
            "Tf" => {
                // Set font and size
                if op.operands.len() >= 2 {
                    if let Ok(name) = op.operands[0].as_name() {
                        current_font = String::from_utf8_lossy(name).to_string();
                    }
                    if let Ok(size) = op.operands[1].as_f32() {
                        current_font_size = size;
                    } else if let Ok(size) = op.operands[1].as_i64() {
                        current_font_size = size as f32;
                    }
                }
            }
            "TL" => {
                // Set text leading (used by T*, ', and " operators)
                if let Some(tl) = op.operands.first().and_then(get_number) {
                    text_leading = tl;
                }
            }
            "Tr" => {
                // Set text rendering mode (3 = invisible / OCR overlay)
                if let Some(mode) = op.operands.first().and_then(get_number) {
                    text_rendering_mode = mode as i32;
                }
            }
            "Td" | "TD" => {
                // Move text position: TLM = T(tx,ty) × TLM; Tm = TLM
                // tx,ty are in text space — must be scaled by the text line matrix
                if op.operands.len() >= 2 {
                    let tx = get_number(&op.operands[0]).unwrap_or(0.0);
                    let ty = get_number(&op.operands[1]).unwrap_or(0.0);
                    line_matrix[4] += tx * line_matrix[0] + ty * line_matrix[2];
                    line_matrix[5] += tx * line_matrix[1] + ty * line_matrix[3];
                    text_matrix = line_matrix;
                    if op.operator == "TD" {
                        text_leading = -ty;
                    }
                }
            }
            "Tm" => {
                // Set text matrix
                if op.operands.len() >= 6 {
                    for (i, operand) in op.operands.iter().take(6).enumerate() {
                        text_matrix[i] =
                            get_number(operand).unwrap_or(if i == 0 || i == 3 { 1.0 } else { 0.0 });
                    }
                    line_matrix = text_matrix;
                }
            }
            "T*" => {
                // Move to start of next line: equivalent to 0 -TL Td
                let tl = if text_leading != 0.0 {
                    text_leading
                } else {
                    current_font_size * 1.2
                };
                line_matrix[4] += (-tl) * line_matrix[2]; // Usually 0 for non-rotated text
                line_matrix[5] += (-tl) * line_matrix[3];
                text_matrix = line_matrix;
            }
            "Tj" => {
                // Show text string
                if in_text_block && !op.operands.is_empty() {
                    // Advance text matrix regardless of visibility
                    let w_ts_opt = font_widths.get(&current_font).and_then(|fi| {
                        get_operand_bytes(&op.operands[0])
                            .map(|raw| compute_string_width_ts(raw, fi, current_font_size))
                    });
                    // ActualText: suppress glyph extraction, just advance text matrix
                    if suppress_glyph_extraction {
                        if let Some(w_ts) = w_ts_opt {
                            text_matrix[4] += w_ts * text_matrix[0];
                            text_matrix[5] += w_ts * text_matrix[1];
                        }
                        continue;
                    }
                    // Skip invisible (white/Tr=3) text but still advance text matrix
                    if fill_is_white || text_rendering_mode == 3 {
                        if let Some(w_ts) = w_ts_opt {
                            text_matrix[4] += w_ts * text_matrix[0];
                            text_matrix[5] += w_ts * text_matrix[1];
                        }
                        continue;
                    }
                    if let Some(text) = extract_text_from_operand(
                        &op.operands[0],
                        &current_font,
                        font_base_names.get(&current_font).map(|s| s.as_str()),
                        font_cmaps,
                        &font_tounicode_refs,
                        &inline_cmaps,
                        &font_encodings,
                        &encoding_cache,
                        &mut cmap_decisions,
                    ) {
                        let combined = multiply_matrices(&text_matrix, &ctm);
                        let rendered_size = effective_font_size(current_font_size, &combined);
                        let (x, y) = (combined[4], combined[5]);
                        let width = if let Some(w_ts) = w_ts_opt {
                            text_matrix[4] += w_ts * text_matrix[0];
                            text_matrix[5] += w_ts * text_matrix[1];
                            (w_ts * (text_matrix[0] * ctm[0] + text_matrix[1] * ctm[2])).abs()
                        } else {
                            0.0
                        };
                        // Only create text item for non-whitespace; whitespace
                        // still advances the text matrix above so gap detection works
                        if !text.trim().is_empty() {
                            let base_font = font_base_names
                                .get(&current_font)
                                .map(|s| s.as_str())
                                .unwrap_or(&current_font);
                            items.push(TextItem {
                                text: expand_ligatures(&text),
                                x,
                                y,
                                width,
                                height: rendered_size,
                                font: current_font.clone(),
                                font_size: rendered_size,
                                page: page_num,
                                is_bold: is_bold_font(base_font),
                                is_italic: is_italic_font(base_font),
                                item_type: ItemType::Text,
                            });
                        }
                    }
                }
            }
            "TJ" => {
                // Show text with positioning — split at column-sized gaps
                if in_text_block && !op.operands.is_empty() {
                    if let Ok(array) = op.operands[0].as_array() {
                        let font_info = font_widths.get(&current_font);
                        let is_invisible =
                            fill_is_white || text_rendering_mode == 3 || suppress_glyph_extraction;

                        // Compute space threshold based on font metrics when available
                        let space_threshold = if let Some(font_info) = font_info {
                            let space_em = font_info.space_width as f32 * font_info.units_scale;
                            let threshold = space_em * 1000.0 * 0.4;
                            threshold.max(80.0)
                        } else {
                            120.0
                        };
                        let column_gap_threshold = space_threshold * 4.0;

                        // Track sub-items for column-gap splitting:
                        // (text, start_width_ts, end_width_ts)
                        let mut sub_items: Vec<(String, f32, f32)> = Vec::new();
                        let mut current_text = String::new();
                        let mut sub_start_width_ts: f32 = 0.0;
                        let mut total_width_ts: f32 = 0.0;
                        for element in array {
                            match element {
                                Object::Integer(n) => {
                                    let n_val = *n as f32;
                                    let displacement = -n_val / 1000.0 * current_font_size;
                                    if !is_invisible
                                        && n_val < -column_gap_threshold
                                        && !current_text.is_empty()
                                    {
                                        // Column gap: flush current segment
                                        sub_items.push((
                                            std::mem::take(&mut current_text),
                                            sub_start_width_ts,
                                            total_width_ts,
                                        ));
                                        total_width_ts += displacement;
                                        sub_start_width_ts = total_width_ts;
                                    } else {
                                        total_width_ts += displacement;
                                        if !is_invisible
                                            && n_val < -space_threshold
                                            && !current_text.is_empty()
                                            && !current_text.ends_with(' ')
                                        {
                                            current_text.push(' ');
                                        }
                                    }
                                    continue;
                                }
                                Object::Real(n) => {
                                    let n_val = *n;
                                    let displacement = -n_val / 1000.0 * current_font_size;
                                    if !is_invisible
                                        && n_val < -column_gap_threshold
                                        && !current_text.is_empty()
                                    {
                                        sub_items.push((
                                            std::mem::take(&mut current_text),
                                            sub_start_width_ts,
                                            total_width_ts,
                                        ));
                                        total_width_ts += displacement;
                                        sub_start_width_ts = total_width_ts;
                                    } else {
                                        total_width_ts += displacement;
                                        if !is_invisible
                                            && n_val < -space_threshold
                                            && !current_text.is_empty()
                                            && !current_text.ends_with(' ')
                                        {
                                            current_text.push(' ');
                                        }
                                    }
                                    continue;
                                }
                                _ => {}
                            }
                            if let Some(fi) = font_info {
                                if let Some(raw_bytes) = get_operand_bytes(element) {
                                    total_width_ts +=
                                        compute_string_width_ts(raw_bytes, fi, current_font_size);
                                }
                            }
                            if !is_invisible {
                                if let Some(text) = extract_text_from_operand(
                                    element,
                                    &current_font,
                                    font_base_names.get(&current_font).map(|s| s.as_str()),
                                    font_cmaps,
                                    &font_tounicode_refs,
                                    &inline_cmaps,
                                    &font_encodings,
                                    &encoding_cache,
                                    &mut cmap_decisions,
                                ) {
                                    current_text.push_str(&text);
                                }
                            }
                        }
                        // Flush remaining text
                        if !is_invisible && !current_text.trim().is_empty() {
                            sub_items.push((current_text, sub_start_width_ts, total_width_ts));
                        }
                        // Emit one TextItem per sub-item
                        if !sub_items.is_empty() {
                            let combined = multiply_matrices(&text_matrix, &ctm);
                            let rendered_size = effective_font_size(current_font_size, &combined);
                            let base_font = font_base_names
                                .get(&current_font)
                                .map(|s| s.as_str())
                                .unwrap_or(&current_font);
                            let scale_x = text_matrix[0] * ctm[0] + text_matrix[1] * ctm[2];
                            for (text, start_w, end_w) in &sub_items {
                                let offset_tm = [
                                    text_matrix[0],
                                    text_matrix[1],
                                    text_matrix[2],
                                    text_matrix[3],
                                    text_matrix[4] + start_w * text_matrix[0],
                                    text_matrix[5] + start_w * text_matrix[1],
                                ];
                                let combined = multiply_matrices(&offset_tm, &ctm);
                                let (x, y) = (combined[4], combined[5]);
                                let width = if font_info.is_some() {
                                    ((end_w - start_w) * scale_x).abs()
                                } else {
                                    0.0
                                };
                                items.push(TextItem {
                                    text: expand_ligatures(text),
                                    x,
                                    y,
                                    width,
                                    height: rendered_size,
                                    font: current_font.clone(),
                                    font_size: rendered_size,
                                    page: page_num,
                                    is_bold: is_bold_font(base_font),
                                    is_italic: is_italic_font(base_font),
                                    item_type: ItemType::Text,
                                });
                            }
                        }
                        // Always advance text matrix by total width
                        if font_info.is_some() {
                            text_matrix[4] += total_width_ts * text_matrix[0];
                            text_matrix[5] += total_width_ts * text_matrix[1];
                        }
                    }
                }
            }
            "'" => {
                // Move to next line and show text (equivalent to T* then Tj)
                let tl = if text_leading != 0.0 {
                    text_leading
                } else {
                    current_font_size * 1.2
                };
                line_matrix[4] += (-tl) * line_matrix[2];
                line_matrix[5] += (-tl) * line_matrix[3];
                text_matrix = line_matrix;
                if !(fill_is_white
                    || text_rendering_mode == 3
                    || suppress_glyph_extraction
                    || op.operands.is_empty())
                {
                    if let Some(text) = extract_text_from_operand(
                        &op.operands[0],
                        &current_font,
                        font_base_names.get(&current_font).map(|s| s.as_str()),
                        font_cmaps,
                        &font_tounicode_refs,
                        &inline_cmaps,
                        &font_encodings,
                        &encoding_cache,
                        &mut cmap_decisions,
                    ) {
                        if !text.trim().is_empty() {
                            let combined = multiply_matrices(&text_matrix, &ctm);
                            let rendered_size = effective_font_size(current_font_size, &combined);
                            let (x, y) = (combined[4], combined[5]);
                            let base_font = font_base_names
                                .get(&current_font)
                                .map(|s| s.as_str())
                                .unwrap_or(&current_font);
                            items.push(TextItem {
                                text: expand_ligatures(&text),
                                x,
                                y,
                                width: 0.0,
                                height: rendered_size,
                                font: current_font.clone(),
                                font_size: rendered_size,
                                page: page_num,
                                is_bold: is_bold_font(base_font),
                                is_italic: is_italic_font(base_font),
                                item_type: ItemType::Text,
                            });
                        }
                    }
                }
            }
            "Do" => {
                // XObject invocation - could be an image or form
                if !op.operands.is_empty() {
                    if let Ok(name) = op.operands[0].as_name() {
                        let xobj_name = String::from_utf8_lossy(name).to_string();

                        if let Some(xobj_type) = xobjects.get(&xobj_name) {
                            match xobj_type {
                                XObjectType::Image => {
                                    // Skip images — text extraction only
                                }
                                XObjectType::Form(form_id) => {
                                    // Extract text from Form XObject
                                    let form_items = extract_form_xobject_text(
                                        doc,
                                        *form_id,
                                        page_num,
                                        font_cmaps,
                                        &ctm,
                                        &mut cmap_decisions,
                                    );
                                    items.extend(form_items);
                                }
                            }
                        }
                    }
                }
            }
            "BMC" => {
                // Begin Marked Content (no properties)
                marked_content_stack.push(None);
            }
            "BDC" => {
                // Begin Marked Content with properties — extract ActualText
                let mut actual_text: Option<String> = None;
                if op.operands.len() >= 2 {
                    let dict = match &op.operands[1] {
                        Object::Dictionary(d) => Some(d.clone()),
                        Object::Reference(id) => doc.get_dictionary(*id).ok().cloned(),
                        _ => None,
                    };
                    if let Some(d) = dict {
                        if let Ok(val) = d.get(b"ActualText") {
                            actual_text = match val {
                                Object::String(bytes, _) => Some(decode_text_string(bytes)),
                                _ => None,
                            };
                        }
                    }
                }
                if actual_text.is_some() {
                    suppress_glyph_extraction = true;
                    actual_text_start_tm = Some(text_matrix);
                }
                marked_content_stack.push(actual_text);
            }
            "EMC" => {
                // End Marked Content — emit ActualText item with correct width
                if let Some(Some(at)) = marked_content_stack.pop() {
                    // Compute width from text matrix advancement during BDC..EMC
                    if let Some(start_tm) = actual_text_start_tm.take() {
                        let combined = multiply_matrices(&start_tm, &ctm);
                        let rendered_size = effective_font_size(current_font_size, &combined);
                        let (x, y) = (combined[4], combined[5]);
                        // Width in device space from text matrix delta
                        let delta_ts = text_matrix[4] - start_tm[4];
                        let scale_x = start_tm[0] * ctm[0] + start_tm[1] * ctm[2];
                        let width = (delta_ts * scale_x).abs();
                        if !at.trim().is_empty() {
                            let base_font = font_base_names
                                .get(&current_font)
                                .map(|s| s.as_str())
                                .unwrap_or(&current_font);
                            items.push(TextItem {
                                text: expand_ligatures(&at),
                                x,
                                y,
                                width,
                                height: rendered_size,
                                font: current_font.clone(),
                                font_size: rendered_size,
                                page: page_num,
                                is_bold: is_bold_font(base_font),
                                is_italic: is_italic_font(base_font),
                                item_type: ItemType::Text,
                            });
                        }
                    }
                    suppress_glyph_extraction = marked_content_stack.iter().any(|a| a.is_some());
                }
            }
            "re" => {
                // Rectangle operator: collect for table-grid detection
                if op.operands.len() >= 4 {
                    let rx = get_number(&op.operands[0]).unwrap_or(0.0);
                    let ry = get_number(&op.operands[1]).unwrap_or(0.0);
                    let rw = get_number(&op.operands[2]).unwrap_or(0.0);
                    let rh = get_number(&op.operands[3]).unwrap_or(0.0);
                    // Transform origin to device space
                    let x_dev = rx * ctm[0] + ry * ctm[2] + ctm[4];
                    let y_dev = rx * ctm[1] + ry * ctm[3] + ctm[5];
                    let w_dev = rw * ctm[0];
                    let h_dev = rh * ctm[3];
                    rects.push(PdfRect {
                        x: x_dev,
                        y: y_dev,
                        width: w_dev,
                        height: h_dev,
                        page: page_num,
                    });
                }
            }
            _ => {}
        }
    }

    let items = super::merge_text_items(items);
    Ok((items, rects))
}
