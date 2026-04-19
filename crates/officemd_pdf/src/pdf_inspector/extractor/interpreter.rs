//! Shared PDF content-stream interpreter for page streams and Form XObjects.

use std::collections::{BTreeMap, HashMap};

use log::trace;
use lopdf::{Dictionary, Document, Encoding, Object, ObjectId, content::Operation};

use crate::pdf_inspector::text_utils::{
    decode_text_string, effective_font_size, expand_ligatures, is_bold_font, is_italic_font,
};
use crate::pdf_inspector::tounicode::{CMapEntry, FontCMaps};
use crate::pdf_inspector::types::{ItemType, PageFontEncodings, PageFontWidths, PdfRect, TextItem};

use super::fonts::{
    CMapDecisionCache, build_font_encodings, build_font_widths, compute_string_width_ts,
    extract_text_from_operand, get_font_file2_obj_num, get_operand_bytes,
};
use super::xobjects::XObjectType;
use super::{
    get_number, is_rotated_text_transform, multiply_matrices, parse_matrix_from_operands,
    parse_rect_operator,
};

#[derive(Debug, Default)]
pub(crate) struct ExtractionSink {
    pub(crate) items: Vec<TextItem>,
    pub(crate) rects: Vec<PdfRect>,
}

/// One entry on the q/Q graphics-state stack. Per PDF 32000-1 §8.4.2, the
/// saved state includes the text state (Tc, Tw, Th, TL, Tf, Tfs, Tr, Trise),
/// so we capture the font name + size alongside the CTM so that `Q` restores
/// the font that was active before the matching `q` — otherwise a `/Fx Tf`
/// inside a nested graphics block leaks out and misroutes later decoding
/// through the wrong font's ToUnicode CMap.
#[derive(Debug, Clone)]
pub(crate) struct GraphicsStackEntry {
    ctm: [f32; 6],
    fill_is_white: bool,
    text_rendering_mode: i32,
    current_font: String,
    current_font_size: f32,
}

#[derive(Debug)]
pub(crate) struct GraphicsState {
    pub(crate) ctm: [f32; 6],
    fill_is_white: bool,
    text_rendering_mode: i32,
    stack: Vec<GraphicsStackEntry>,
}

impl GraphicsState {
    fn new(initial_ctm: [f32; 6]) -> Self {
        Self {
            ctm: initial_ctm,
            fill_is_white: false,
            text_rendering_mode: 0,
            stack: Vec::new(),
        }
    }

    fn is_invisible(&self) -> bool {
        self.fill_is_white || self.text_rendering_mode == 3
    }
}

#[derive(Debug, Default)]
pub(crate) struct TextState {
    current_font: String,
    current_font_size: f32,
    text_leading: f32,
    text_matrix: [f32; 6],
    line_matrix: [f32; 6],
    in_text_block: bool,
}

impl TextState {
    fn new() -> Self {
        Self {
            current_font: String::new(),
            current_font_size: 12.0,
            text_leading: 0.0,
            text_matrix: [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
            line_matrix: [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
            in_text_block: false,
        }
    }

    fn reset_text_block(&mut self) {
        self.in_text_block = true;
        self.text_matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
        self.line_matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
    }
}

pub(crate) struct FontContext<'a> {
    font_encodings: PageFontEncodings,
    font_widths: PageFontWidths,
    font_base_names: HashMap<String, String>,
    font_tounicode_refs: HashMap<String, u32>,
    inline_cmaps: HashMap<String, CMapEntry>,
    encoding_cache: HashMap<String, Encoding<'a>>,
}

impl<'a> FontContext<'a> {
    fn base_font_name<'b>(&'b self, current_font: &'b str) -> &'b str {
        self.font_base_names
            .get(current_font)
            .map(|s| s.as_str())
            .unwrap_or(current_font)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum MarkedContentMode {
    Disabled,
    ActualText,
}

#[derive(Debug, Clone)]
struct MarkedContentState {
    actual_text: Option<String>,
    saw_visible_text: bool,
    start_tm: Option<[f32; 6]>,
}

pub(crate) fn build_font_context<'a>(
    doc: &'a Document,
    fonts: &BTreeMap<Vec<u8>, &'a Dictionary>,
) -> FontContext<'a> {
    let font_encodings = build_font_encodings(doc, fonts);
    let font_widths = build_font_widths(doc, fonts);

    let mut font_base_names = HashMap::new();
    let mut font_tounicode_refs = HashMap::new();
    let mut inline_cmaps = HashMap::new();
    let mut encoding_cache = HashMap::new();

    for (font_name, font_dict) in fonts {
        let resource_name = String::from_utf8_lossy(font_name).to_string();

        if let Ok(base_font) = font_dict.get(b"BaseFont") {
            if let Ok(name) = base_font.as_name() {
                let base_name = String::from_utf8_lossy(name).to_string();
                font_base_names.insert(resource_name.clone(), base_name);
            }
        }

        match font_dict.get(b"ToUnicode") {
            Ok(tounicode) => {
                if let Ok(obj_ref) = tounicode.as_reference() {
                    font_tounicode_refs.insert(resource_name.clone(), obj_ref.0);
                } else if let Object::Stream(s) = tounicode {
                    if let Ok(data) = s.decompressed_content() {
                        if let Some(entry) =
                            crate::pdf_inspector::tounicode::build_cmap_entry_from_stream(
                                &data, font_dict, doc, 0,
                            )
                        {
                            inline_cmaps.insert(resource_name.clone(), entry);
                        }
                    }
                }
            }
            Err(_) => {
                if let Some(ff2_obj_num) = get_font_file2_obj_num(doc, font_dict) {
                    font_tounicode_refs.insert(resource_name.clone(), ff2_obj_num);
                }
            }
        }

        if let Ok(enc) = font_dict.get_font_encoding(doc) {
            encoding_cache.insert(resource_name, enc);
        }
    }

    FontContext {
        font_encodings,
        font_widths,
        font_base_names,
        font_tounicode_refs,
        inline_cmaps,
        encoding_cache,
    }
}

fn advance_text_matrix(text_state: &mut TextState, width_ts: f32) {
    text_state.text_matrix[4] += width_ts * text_state.text_matrix[0];
    text_state.text_matrix[5] += width_ts * text_state.text_matrix[1];
}

fn note_visible_text(marked_content_stack: &mut [MarkedContentState], start_tm: [f32; 6]) {
    if let Some(entry) = marked_content_stack
        .iter_mut()
        .rev()
        .find(|entry| entry.actual_text.is_some())
    {
        entry.saw_visible_text = true;
        entry.start_tm.get_or_insert(start_tm);
    }
}

fn emit_text_item(
    sink: &mut ExtractionSink,
    text_state: &TextState,
    graphics_state: &GraphicsState,
    page_num: u32,
    base_font: &str,
    text: &str,
    width: f32,
    origin_tm: [f32; 6],
) {
    let combined = multiply_matrices(&origin_tm, &graphics_state.ctm);
    let rendered_size = effective_font_size(text_state.current_font_size, &combined);
    let (x, y) = (combined[4], combined[5]);

    sink.items.push(TextItem {
        text: expand_ligatures(text),
        x,
        y,
        width,
        height: rendered_size,
        font: text_state.current_font.clone(),
        font_size: rendered_size,
        page: page_num,
        is_rotated: is_rotated_text_transform(&combined),
        is_bold: is_bold_font(base_font),
        is_italic: is_italic_font(base_font),
        item_type: ItemType::Text,
    });
}

fn handle_show_text(
    sink: &mut ExtractionSink,
    text_state: &mut TextState,
    graphics_state: &GraphicsState,
    font_ctx: &FontContext<'_>,
    font_cmaps: &FontCMaps,
    operand: &Object,
    page_num: u32,
    cmap_decisions: &mut CMapDecisionCache,
    marked_content_stack: &mut Vec<MarkedContentState>,
) {
    if !text_state.in_text_block {
        return;
    }

    let width_ts = font_ctx
        .font_widths
        .get(&text_state.current_font)
        .and_then(|fi| {
            get_operand_bytes(operand)
                .map(|raw| compute_string_width_ts(raw, fi, text_state.current_font_size))
        });

    if graphics_state.is_invisible() {
        if let Some(width_ts) = width_ts {
            advance_text_matrix(text_state, width_ts);
        }
        return;
    }

    if let Some(text) = extract_text_from_operand(
        operand,
        &text_state.current_font,
        font_ctx
            .font_base_names
            .get(&text_state.current_font)
            .map(|s| s.as_str()),
        font_cmaps,
        &font_ctx.font_tounicode_refs,
        &font_ctx.inline_cmaps,
        &font_ctx.font_encodings,
        &font_ctx.encoding_cache,
        cmap_decisions,
    ) {
        let origin_tm = text_state.text_matrix;
        let width = if let Some(width_ts) = width_ts {
            advance_text_matrix(text_state, width_ts);
            (width_ts
                * (text_state.text_matrix[0] * graphics_state.ctm[0]
                    + text_state.text_matrix[1] * graphics_state.ctm[2]))
                .abs()
        } else {
            0.0
        };

        if !text.trim().is_empty() {
            note_visible_text(marked_content_stack, text_state.text_matrix);
            emit_text_item(
                sink,
                text_state,
                graphics_state,
                page_num,
                font_ctx.base_font_name(&text_state.current_font),
                &text,
                width,
                origin_tm,
            );
        }
    }
}

fn handle_show_text_array(
    sink: &mut ExtractionSink,
    text_state: &mut TextState,
    graphics_state: &GraphicsState,
    font_ctx: &FontContext<'_>,
    font_cmaps: &FontCMaps,
    array: &[Object],
    page_num: u32,
    cmap_decisions: &mut CMapDecisionCache,
    marked_content_stack: &mut Vec<MarkedContentState>,
) {
    if !text_state.in_text_block {
        return;
    }

    let font_info = font_ctx.font_widths.get(&text_state.current_font);
    let is_invisible = graphics_state.is_invisible();

    let space_threshold = if let Some(font_info) = font_info {
        let space_em = font_info.space_width as f32 * font_info.units_scale;
        let threshold = space_em * 1000.0 * 0.4;
        threshold.max(80.0)
    } else {
        120.0
    };
    let column_gap_threshold = space_threshold * 4.0;

    let mut sub_items: Vec<(String, f32, f32)> = Vec::new();
    let mut current_text = String::new();
    let mut sub_start_width_ts = 0.0;
    let mut total_width_ts = 0.0;

    for element in array {
        match element {
            Object::Integer(n) => {
                let n_val = *n as f32;
                let displacement = -n_val / 1000.0 * text_state.current_font_size;
                if !is_invisible && n_val < -column_gap_threshold && !current_text.is_empty() {
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
                let displacement = -*n / 1000.0 * text_state.current_font_size;
                if !is_invisible && *n < -column_gap_threshold && !current_text.is_empty() {
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
                        && *n < -space_threshold
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
                    compute_string_width_ts(raw_bytes, fi, text_state.current_font_size);
            }
        }
        if !is_invisible {
            if let Some(text) = extract_text_from_operand(
                element,
                &text_state.current_font,
                font_ctx
                    .font_base_names
                    .get(&text_state.current_font)
                    .map(|s| s.as_str()),
                font_cmaps,
                &font_ctx.font_tounicode_refs,
                &font_ctx.inline_cmaps,
                &font_ctx.font_encodings,
                &font_ctx.encoding_cache,
                cmap_decisions,
            ) {
                current_text.push_str(&text);
            }
        }
    }

    if !is_invisible && !current_text.trim().is_empty() {
        sub_items.push((current_text, sub_start_width_ts, total_width_ts));
    }

    if !sub_items.is_empty() {
        let base_font = font_ctx.base_font_name(&text_state.current_font);
        let scale_x = text_state.text_matrix[0] * graphics_state.ctm[0]
            + text_state.text_matrix[1] * graphics_state.ctm[2];
        for (text, start_w, end_w) in &sub_items {
            let offset_tm = [
                text_state.text_matrix[0],
                text_state.text_matrix[1],
                text_state.text_matrix[2],
                text_state.text_matrix[3],
                text_state.text_matrix[4] + start_w * text_state.text_matrix[0],
                text_state.text_matrix[5] + start_w * text_state.text_matrix[1],
            ];
            let width = if font_info.is_some() {
                ((end_w - start_w) * scale_x).abs()
            } else {
                0.0
            };
            note_visible_text(marked_content_stack, offset_tm);
            emit_text_item(
                sink,
                text_state,
                graphics_state,
                page_num,
                base_font,
                text,
                width,
                offset_tm,
            );
        }
    }

    if font_info.is_some() {
        advance_text_matrix(text_state, total_width_ts);
    }
}

fn handle_next_line_show_text(
    sink: &mut ExtractionSink,
    text_state: &mut TextState,
    graphics_state: &GraphicsState,
    font_ctx: &FontContext<'_>,
    font_cmaps: &FontCMaps,
    operand: &Object,
    page_num: u32,
    cmap_decisions: &mut CMapDecisionCache,
    marked_content_stack: &mut Vec<MarkedContentState>,
) {
    let tl = if text_state.text_leading != 0.0 {
        text_state.text_leading
    } else {
        text_state.current_font_size * 1.2
    };
    text_state.line_matrix[4] += (-tl) * text_state.line_matrix[2];
    text_state.line_matrix[5] += (-tl) * text_state.line_matrix[3];
    text_state.text_matrix = text_state.line_matrix;

    if graphics_state.is_invisible() || !text_state.in_text_block {
        return;
    }

    if let Some(text) = extract_text_from_operand(
        operand,
        &text_state.current_font,
        font_ctx
            .font_base_names
            .get(&text_state.current_font)
            .map(|s| s.as_str()),
        font_cmaps,
        &font_ctx.font_tounicode_refs,
        &font_ctx.inline_cmaps,
        &font_ctx.font_encodings,
        &font_ctx.encoding_cache,
        cmap_decisions,
    ) {
        if !text.trim().is_empty() {
            note_visible_text(marked_content_stack, text_state.text_matrix);
            emit_text_item(
                sink,
                text_state,
                graphics_state,
                page_num,
                font_ctx.base_font_name(&text_state.current_font),
                &text,
                0.0,
                text_state.text_matrix,
            );
        }
    }
}

fn handle_marked_content_begin(
    doc: &Document,
    op: &Operation,
    text_state: &TextState,
    marked_content_stack: &mut Vec<MarkedContentState>,
) {
    let mut actual_text = None;
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
    marked_content_stack.push(MarkedContentState {
        start_tm: actual_text.as_ref().map(|_| text_state.text_matrix),
        actual_text,
        saw_visible_text: false,
    });
}

fn handle_marked_content_end(
    sink: &mut ExtractionSink,
    text_state: &TextState,
    graphics_state: &GraphicsState,
    font_ctx: &FontContext<'_>,
    page_num: u32,
    marked_content_stack: &mut Vec<MarkedContentState>,
) {
    let Some(state) = marked_content_stack.pop() else {
        return;
    };
    let Some(actual_text) = state.actual_text else {
        return;
    };
    if state.saw_visible_text {
        return;
    }
    let Some(start_tm) = state.start_tm else {
        return;
    };
    if actual_text.trim().is_empty() {
        return;
    }

    let delta_ts = text_state.text_matrix[4] - start_tm[4];
    let scale_x = start_tm[0] * graphics_state.ctm[0] + start_tm[1] * graphics_state.ctm[2];
    let width = (delta_ts * scale_x).abs();
    emit_text_item(
        sink,
        text_state,
        graphics_state,
        page_num,
        font_ctx.base_font_name(&text_state.current_font),
        &actual_text,
        width,
        start_tm,
    );
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn extract_content_stream<F>(
    doc: &Document,
    operations: &[Operation],
    page_num: u32,
    font_cmaps: &FontCMaps,
    font_ctx: &FontContext<'_>,
    xobjects: &HashMap<String, XObjectType>,
    initial_ctm: [f32; 6],
    marked_content_mode: MarkedContentMode,
    cmap_decisions: &mut CMapDecisionCache,
    mut extract_nested_form: F,
) -> ExtractionSink
where
    F: FnMut(ObjectId, &[f32; 6], &mut CMapDecisionCache) -> (Vec<TextItem>, Vec<PdfRect>),
{
    let mut sink = ExtractionSink::default();
    let mut graphics_state = GraphicsState::new(initial_ctm);
    let mut text_state = TextState::new();
    let mut marked_content_stack: Vec<MarkedContentState> = Vec::new();

    for op in operations {
        trace!("{} {:?}", op.operator, op.operands);
        match op.operator.as_str() {
            "q" => {
                graphics_state.stack.push(GraphicsStackEntry {
                    ctm: graphics_state.ctm,
                    fill_is_white: graphics_state.fill_is_white,
                    text_rendering_mode: graphics_state.text_rendering_mode,
                    current_font: text_state.current_font.clone(),
                    current_font_size: text_state.current_font_size,
                });
            }
            "Q" => {
                if let Some(saved) = graphics_state.stack.pop() {
                    graphics_state.ctm = saved.ctm;
                    graphics_state.fill_is_white = saved.fill_is_white;
                    graphics_state.text_rendering_mode = saved.text_rendering_mode;
                    text_state.current_font = saved.current_font;
                    text_state.current_font_size = saved.current_font_size;
                }
            }
            "cm" => {
                if let Some(new_matrix) = parse_matrix_from_operands(&op.operands) {
                    graphics_state.ctm = multiply_matrices(&new_matrix, &graphics_state.ctm);
                }
            }
            "g" => {
                if let Some(gray) = op.operands.first().and_then(get_number) {
                    graphics_state.fill_is_white = gray > 0.95;
                }
            }
            "rg" => {
                if op.operands.len() >= 3 {
                    let r = get_number(&op.operands[0]).unwrap_or(0.0);
                    let g = get_number(&op.operands[1]).unwrap_or(0.0);
                    let b = get_number(&op.operands[2]).unwrap_or(0.0);
                    graphics_state.fill_is_white = r > 0.95 && g > 0.95 && b > 0.95;
                }
            }
            "k" => {
                if op.operands.len() >= 4 {
                    let c = get_number(&op.operands[0]).unwrap_or(1.0);
                    let m = get_number(&op.operands[1]).unwrap_or(1.0);
                    let y = get_number(&op.operands[2]).unwrap_or(1.0);
                    let k = get_number(&op.operands[3]).unwrap_or(1.0);
                    graphics_state.fill_is_white = c < 0.05 && m < 0.05 && y < 0.05 && k < 0.05;
                }
            }
            "BT" => {
                text_state.reset_text_block();
                graphics_state.text_rendering_mode = 0;
            }
            "ET" => {
                text_state.in_text_block = false;
            }
            "Tf" => {
                if op.operands.len() >= 2 {
                    if let Ok(name) = op.operands[0].as_name() {
                        text_state.current_font = String::from_utf8_lossy(name).to_string();
                    }
                    text_state.current_font_size = get_number(&op.operands[1]).unwrap_or(12.0);
                }
            }
            "TL" => {
                if let Some(tl) = op.operands.first().and_then(get_number) {
                    text_state.text_leading = tl;
                }
            }
            "Tr" => {
                if let Some(mode) = op.operands.first().and_then(get_number) {
                    graphics_state.text_rendering_mode = mode as i32;
                }
            }
            "Td" | "TD" => {
                if op.operands.len() >= 2 {
                    let tx = get_number(&op.operands[0]).unwrap_or(0.0);
                    let ty = get_number(&op.operands[1]).unwrap_or(0.0);
                    text_state.line_matrix[4] +=
                        tx * text_state.line_matrix[0] + ty * text_state.line_matrix[2];
                    text_state.line_matrix[5] +=
                        tx * text_state.line_matrix[1] + ty * text_state.line_matrix[3];
                    text_state.text_matrix = text_state.line_matrix;
                    if op.operator == "TD" {
                        text_state.text_leading = -ty;
                    }
                }
            }
            "Tm" => {
                if op.operands.len() >= 6 {
                    for (i, operand) in op.operands.iter().take(6).enumerate() {
                        text_state.text_matrix[i] =
                            get_number(operand).unwrap_or(if i == 0 || i == 3 { 1.0 } else { 0.0 });
                    }
                    text_state.line_matrix = text_state.text_matrix;
                }
            }
            "T*" => {
                let tl = if text_state.text_leading != 0.0 {
                    text_state.text_leading
                } else {
                    text_state.current_font_size * 1.2
                };
                text_state.line_matrix[4] += (-tl) * text_state.line_matrix[2];
                text_state.line_matrix[5] += (-tl) * text_state.line_matrix[3];
                text_state.text_matrix = text_state.line_matrix;
            }
            "Tj" => handle_show_text(
                &mut sink,
                &mut text_state,
                &graphics_state,
                font_ctx,
                font_cmaps,
                &op.operands[0],
                page_num,
                cmap_decisions,
                &mut marked_content_stack,
            ),
            "TJ" => {
                if let Ok(array) = op.operands[0].as_array() {
                    handle_show_text_array(
                        &mut sink,
                        &mut text_state,
                        &graphics_state,
                        font_ctx,
                        font_cmaps,
                        array,
                        page_num,
                        cmap_decisions,
                        &mut marked_content_stack,
                    );
                }
            }
            "'" => handle_next_line_show_text(
                &mut sink,
                &mut text_state,
                &graphics_state,
                font_ctx,
                font_cmaps,
                &op.operands[0],
                page_num,
                cmap_decisions,
                &mut marked_content_stack,
            ),
            "Do" => {
                if !op.operands.is_empty() {
                    if let Ok(name) = op.operands[0].as_name() {
                        let xobj_name = String::from_utf8_lossy(name).to_string();
                        if let Some(xobj_type) = xobjects.get(&xobj_name) {
                            match xobj_type {
                                XObjectType::Image => {}
                                XObjectType::Form(form_id) => {
                                    let (items, rects) = extract_nested_form(
                                        *form_id,
                                        &graphics_state.ctm,
                                        cmap_decisions,
                                    );
                                    sink.items.extend(items);
                                    sink.rects.extend(rects);
                                }
                            }
                        }
                    }
                }
            }
            "BMC" if marked_content_mode == MarkedContentMode::ActualText => {
                marked_content_stack.push(MarkedContentState {
                    actual_text: None,
                    saw_visible_text: false,
                    start_tm: None,
                });
            }
            "BDC" if marked_content_mode == MarkedContentMode::ActualText => {
                handle_marked_content_begin(doc, op, &text_state, &mut marked_content_stack);
            }
            "EMC" if marked_content_mode == MarkedContentMode::ActualText => {
                handle_marked_content_end(
                    &mut sink,
                    &text_state,
                    &graphics_state,
                    font_ctx,
                    page_num,
                    &mut marked_content_stack,
                );
            }
            "re" => {
                if let Some(rect) = parse_rect_operator(&op.operands, &graphics_state.ctm, page_num)
                {
                    sink.rects.push(rect);
                }
            }
            _ => {}
        }
    }

    sink
}
