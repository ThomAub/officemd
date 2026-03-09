//! Font width parsing, encoding, and text decoding.

use crate::pdf_inspector::glyph_names::glyph_to_char;
use crate::pdf_inspector::tounicode::FontCMaps;
use crate::pdf_inspector::types::{
    FontEncodingMap, FontWidthInfo, PageFontEncodings, PageFontWidths,
};
use log::debug;
use lopdf::{Document, Encoding, Object};
use std::collections::HashMap;

#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub(crate) enum CMapChoice {
    Primary,
    Remapped,
}

#[derive(Debug, Default, Clone)]
pub(crate) struct CMapDecisionCache {
    decisions: HashMap<u32, CMapDecision>,
}

#[derive(Debug, Default, Clone)]
struct CMapDecision {
    primary_sample: String,
    remapped_sample: String,
    sample_bytes: usize,
    choice: Option<CMapChoice>,
}

impl CMapDecisionCache {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn get_choice(&self, obj_num: u32) -> Option<CMapChoice> {
        self.decisions.get(&obj_num).and_then(|d| d.choice)
    }

    pub(crate) fn consider(
        &mut self,
        obj_num: u32,
        primary: &str,
        remapped: &str,
        bytes_len: usize,
    ) -> Option<CMapChoice> {
        const SAMPLE_TARGET_BYTES: usize = 240;

        let entry = self.decisions.entry(obj_num).or_default();
        entry.sample_bytes = entry.sample_bytes.saturating_add(bytes_len);
        entry.primary_sample.push_str(primary);
        entry.remapped_sample.push_str(remapped);

        if entry.choice.is_none() && entry.sample_bytes >= SAMPLE_TARGET_BYTES {
            let score_primary = score_text(&entry.primary_sample);
            let score_remap = score_text(&entry.remapped_sample);
            entry.choice = if score_remap > score_primary + 5 {
                Some(CMapChoice::Remapped)
            } else {
                Some(CMapChoice::Primary)
            };
        }

        entry.choice
    }
}

/// Resolve a PDF object reference to an array
pub(crate) fn resolve_array<'a>(doc: &'a Document, obj: &'a Object) -> Option<&'a Vec<Object>> {
    match obj {
        Object::Array(arr) => Some(arr),
        Object::Reference(r) => {
            if let Ok(Object::Array(arr)) = doc.get_object(*r) {
                Some(arr)
            } else {
                None
            }
        }
        _ => None,
    }
}

/// Resolve a PDF object reference to a dictionary
pub(crate) fn resolve_dict<'a>(
    doc: &'a Document,
    obj: &'a Object,
) -> Option<&'a lopdf::Dictionary> {
    match obj {
        Object::Dictionary(d) => Some(d),
        Object::Reference(r) => doc.get_dictionary(*r).ok(),
        _ => None,
    }
}

/// Build font width info for all fonts on a page
pub(crate) fn build_font_widths(
    doc: &Document,
    fonts: &std::collections::BTreeMap<Vec<u8>, &lopdf::Dictionary>,
) -> PageFontWidths {
    let mut widths = PageFontWidths::new();

    for (font_name, font_dict) in fonts {
        let resource_name = String::from_utf8_lossy(font_name).to_string();

        let subtype = font_dict
            .get(b"Subtype")
            .ok()
            .and_then(|o| o.as_name().ok())
            .map(|n| String::from_utf8_lossy(n).to_string())
            .unwrap_or_default();
        let base_font = font_dict
            .get(b"BaseFont")
            .ok()
            .and_then(|o| o.as_name().ok())
            .map(|n| String::from_utf8_lossy(n).to_string())
            .unwrap_or_default();
        let has_tounicode = font_dict.get(b"ToUnicode").is_ok();
        let has_descendants = font_dict.get(b"DescendantFonts").is_ok();
        let encoding_str = font_dict
            .get(b"Encoding")
            .ok()
            .map(|o| match o {
                Object::Name(n) => String::from_utf8_lossy(n).to_string(),
                Object::Reference(_) => "ref(dict)".to_string(),
                Object::Dictionary(_) => "dict".to_string(),
                _ => format!("{:?}", o),
            })
            .unwrap_or_else(|| "none".to_string());

        debug!(
            "font {:<10} sub={:<12} base={:<45} toUni={:<6} enc={:<20} cid={}",
            resource_name, subtype, base_font, has_tounicode, encoding_str, has_descendants
        );

        if let Some(info) = parse_font_widths(doc, font_dict) {
            widths.insert(resource_name, info);
        }
    }

    widths
}

/// Parse font widths from a font dictionary, dispatching by Subtype
pub(crate) fn parse_font_widths(
    doc: &Document,
    font_dict: &lopdf::Dictionary,
) -> Option<FontWidthInfo> {
    // Get the font subtype
    let subtype = font_dict.get(b"Subtype").ok()?;
    let subtype_name = subtype.as_name().ok()?;

    match subtype_name {
        b"Type0" => parse_type0_widths(doc, font_dict),
        b"Type1" | b"TrueType" | b"MMType1" | b"Type3" => parse_simple_font_widths(doc, font_dict),
        _ => None,
    }
}

/// Parse widths for simple fonts (Type1, TrueType, MMType1, Type3)
/// Reads FirstChar, LastChar, and Widths array.
/// For Type3 fonts, reads FontMatrix to determine the correct units_scale.
pub(crate) fn parse_simple_font_widths(
    doc: &Document,
    font_dict: &lopdf::Dictionary,
) -> Option<FontWidthInfo> {
    let first_char = font_dict.get(b"FirstChar").ok().and_then(|o| match o {
        Object::Integer(n) => Some(*n as u16),
        Object::Reference(r) => doc.get_object(*r).ok().and_then(|o| {
            if let Object::Integer(n) = o {
                Some(*n as u16)
            } else {
                None
            }
        }),
        _ => None,
    })?;

    let last_char = font_dict.get(b"LastChar").ok().and_then(|o| match o {
        Object::Integer(n) => Some(*n as u16),
        Object::Reference(r) => doc.get_object(*r).ok().and_then(|o| {
            if let Object::Integer(n) = o {
                Some(*n as u16)
            } else {
                None
            }
        }),
        _ => None,
    })?;

    let widths_obj = font_dict.get(b"Widths").ok()?;
    let widths_array = resolve_array(doc, widths_obj)?;

    let mut widths = HashMap::new();
    let mut space_width: u16 = 0;

    for (i, w_obj) in widths_array.iter().enumerate() {
        let code = first_char + i as u16;
        if code > last_char {
            break;
        }
        let w = match w_obj {
            Object::Integer(n) => *n as u16,
            Object::Real(n) => *n as u16,
            Object::Reference(r) => {
                if let Ok(obj) = doc.get_object(*r) {
                    match obj {
                        Object::Integer(n) => *n as u16,
                        Object::Real(n) => *n as u16,
                        _ => continue,
                    }
                } else {
                    continue;
                }
            }
            _ => continue,
        };
        if code == 32 {
            space_width = w;
        }
        widths.insert(code, w);
    }

    // Determine units_scale: for Type3 fonts, use FontMatrix[0]; for others, use 1/1000
    let units_scale = if let Ok(fm) = font_dict.get(b"FontMatrix") {
        if let Some(arr) = resolve_array(doc, fm) {
            if !arr.is_empty() {
                match &arr[0] {
                    Object::Real(r) => r.abs(),
                    Object::Integer(i) => (*i as f32).abs(),
                    _ => 0.001,
                }
            } else {
                0.001
            }
        } else {
            0.001
        }
    } else {
        0.001 // Standard 1000-unit system
    };

    // If space width wasn't found in the table, estimate from font metrics.
    // The default of 250 is calibrated for standard 1000-unit fonts (units_scale=0.001).
    // For Type3 fonts with different coordinate systems, use average glyph width instead.
    if space_width == 0 {
        if !widths.is_empty() && (units_scale - 0.001).abs() > 0.0005 {
            // Non-standard scale: estimate space as ~45% of average glyph width
            let sum: u32 = widths.values().map(|&w| w as u32).sum();
            let avg = sum as f32 / widths.len() as f32;
            space_width = (avg * 0.45).max(1.0) as u16;
        } else {
            space_width = 250;
        }
    }

    Some(FontWidthInfo {
        widths,
        default_width: 0,
        space_width,
        is_cid: false,
        units_scale,
        wmode: 0,
    })
}

/// Parse widths for Type0 (composite/CID) fonts
/// Reads DescendantFonts → CIDFont → W array and DW value
pub(crate) fn parse_type0_widths(
    doc: &Document,
    font_dict: &lopdf::Dictionary,
) -> Option<FontWidthInfo> {
    let desc_fonts_obj = font_dict.get(b"DescendantFonts").ok()?;
    let desc_fonts = resolve_array(doc, desc_fonts_obj)?;

    if desc_fonts.is_empty() {
        return None;
    }

    // Get the first descendant font dictionary
    let cid_font_dict = resolve_dict(doc, &desc_fonts[0])?;

    // Get DW (default width)
    let default_width = cid_font_dict
        .get(b"DW")
        .ok()
        .and_then(|o| match o {
            Object::Integer(n) => Some(*n as u16),
            Object::Real(n) => Some(*n as u16),
            _ => None,
        })
        .unwrap_or(1000);

    let mut widths = HashMap::new();

    // Parse W array if present
    if let Ok(w_obj) = cid_font_dict.get(b"W") {
        if let Some(w_array) = resolve_array(doc, w_obj) {
            parse_cid_w_array(doc, w_array, &mut widths);
        }
    }

    // Try to determine space width (CID 32 or CID 3 are common for space)
    let space_width = widths
        .get(&32)
        .or_else(|| widths.get(&3))
        .copied()
        .unwrap_or(if default_width > 0 {
            default_width / 4
        } else {
            250
        });

    let wmode = font_dict
        .get(b"WMode")
        .ok()
        .and_then(|o| match o {
            Object::Integer(n) => Some(*n as u8),
            _ => None,
        })
        .unwrap_or(0);

    Some(FontWidthInfo {
        widths,
        default_width,
        space_width,
        is_cid: true,
        units_scale: 0.001, // CID fonts use standard 1000-unit system
        wmode,
    })
}

/// Parse a CID W array into widths map
/// Format: [c [w1 w2 ...]] (consecutive from c) or [c_first c_last w] (range with same width)
pub(crate) fn parse_cid_w_array(
    doc: &Document,
    w_array: &[Object],
    widths: &mut HashMap<u16, u16>,
) {
    let mut i = 0;
    while i < w_array.len() {
        let start_cid = match &w_array[i] {
            Object::Integer(n) => *n as u16,
            Object::Real(n) => *n as u16,
            _ => {
                i += 1;
                continue;
            }
        };
        i += 1;
        if i >= w_array.len() {
            break;
        }

        // Check if next element is an array (consecutive widths) or integer (range)
        match &w_array[i] {
            Object::Array(arr) => {
                // [c [w1 w2 ...]] — consecutive widths starting at c
                for (j, w_obj) in arr.iter().enumerate() {
                    let w = match w_obj {
                        Object::Integer(n) => *n as u16,
                        Object::Real(n) => *n as u16,
                        _ => continue,
                    };
                    widths.insert(start_cid + j as u16, w);
                }
                i += 1;
            }
            Object::Reference(r) => {
                // Could be a reference to an array
                if let Ok(Object::Array(arr)) = doc.get_object(*r) {
                    for (j, w_obj) in arr.iter().enumerate() {
                        let w = match w_obj {
                            Object::Integer(n) => *n as u16,
                            Object::Real(n) => *n as u16,
                            _ => continue,
                        };
                        widths.insert(start_cid + j as u16, w);
                    }
                    i += 1;
                } else {
                    // Treat as c_first c_last w
                    i += 1; // skip this
                }
            }
            Object::Integer(end_cid) => {
                // [c_first c_last w] — range with uniform width
                let end = *end_cid as u16;
                i += 1;
                if i >= w_array.len() {
                    break;
                }
                let w = match &w_array[i] {
                    Object::Integer(n) => *n as u16,
                    Object::Real(n) => *n as u16,
                    _ => {
                        i += 1;
                        continue;
                    }
                };
                for cid in start_cid..=end {
                    widths.insert(cid, w);
                }
                i += 1;
            }
            Object::Real(end_cid) => {
                let end = *end_cid as u16;
                i += 1;
                if i >= w_array.len() {
                    break;
                }
                let w = match &w_array[i] {
                    Object::Integer(n) => *n as u16,
                    Object::Real(n) => *n as u16,
                    _ => {
                        i += 1;
                        continue;
                    }
                };
                for cid in start_cid..=end {
                    widths.insert(cid, w);
                }
                i += 1;
            }
            _ => {
                i += 1;
            }
        }
    }
}

/// Compute the width of a string in text space units,
/// given raw bytes and font width info.
/// Returns width in text space units (font_units * units_scale * font_size).
pub(crate) fn compute_string_width_ts(
    bytes: &[u8],
    font_info: &FontWidthInfo,
    font_size: f32,
) -> f32 {
    let mut total: f32 = 0.0;
    if font_info.is_cid {
        // 2-byte (big-endian) character codes
        let mut j = 0;
        while j + 1 < bytes.len() {
            let cid = u16::from_be_bytes([bytes[j], bytes[j + 1]]);
            let w = font_info
                .widths
                .get(&cid)
                .copied()
                .unwrap_or(font_info.default_width);
            total += w as f32;
            j += 2;
        }
    } else {
        // 1-byte character codes
        for &b in bytes {
            let code = b as u16;
            let w = font_info
                .widths
                .get(&code)
                .copied()
                .unwrap_or(font_info.default_width);
            total += w as f32;
        }
    }
    // Convert from font units to text space using the font's scale factor
    total * font_info.units_scale * font_size
}

/// Extract raw bytes from a PDF operand (String object)
pub(crate) fn get_operand_bytes(obj: &Object) -> Option<&[u8]> {
    if let Object::String(bytes, _) = obj {
        Some(bytes)
    } else {
        None
    }
}

/// Build encoding maps for all fonts on a page
pub(crate) fn build_font_encodings(
    doc: &Document,
    fonts: &std::collections::BTreeMap<Vec<u8>, &lopdf::Dictionary>,
) -> PageFontEncodings {
    let mut encodings = PageFontEncodings::new();

    for (font_name, font_dict) in fonts {
        let resource_name = String::from_utf8_lossy(font_name).to_string();

        if let Some(encoding_map) = parse_font_encoding(doc, font_dict) {
            encodings.insert(resource_name, encoding_map);
        }
    }

    encodings
}

/// Parse font encoding from a font dictionary
pub(crate) fn parse_font_encoding(
    doc: &Document,
    font_dict: &lopdf::Dictionary,
) -> Option<FontEncodingMap> {
    let encoding_obj = font_dict.get(b"Encoding").ok()?;

    // Encoding can be a name or a dictionary
    match encoding_obj {
        Object::Name(_name) => {
            // Standard encoding name (e.g., MacRomanEncoding, WinAnsiEncoding)
            // For standard encodings, we can use the standard tables
            // But we still need to check for Differences
            None // Let lopdf handle standard encodings
        }
        Object::Reference(obj_ref) => {
            // Reference to encoding dictionary
            if let Ok(enc_dict) = doc.get_dictionary(*obj_ref) {
                parse_encoding_dictionary(doc, enc_dict)
            } else {
                None
            }
        }
        Object::Dictionary(enc_dict) => parse_encoding_dictionary(doc, enc_dict),
        _ => None,
    }
}

/// Parse an encoding dictionary with Differences array
pub(crate) fn parse_encoding_dictionary(
    doc: &Document,
    enc_dict: &lopdf::Dictionary,
) -> Option<FontEncodingMap> {
    let differences = enc_dict.get(b"Differences").ok()?;

    let diff_array = match differences {
        Object::Array(arr) => arr.clone(),
        Object::Reference(obj_ref) => {
            if let Ok(Object::Array(arr)) = doc.get_object(*obj_ref) {
                arr.clone()
            } else {
                return None;
            }
        }
        _ => return None,
    };

    let mut encoding_map = FontEncodingMap::new();
    let mut current_code: u8 = 0;
    let mut ligature_count = 0u32;

    for item in diff_array {
        match item {
            Object::Integer(n) => {
                // This sets the starting code for subsequent glyph names
                current_code = n as u8;
            }
            Object::Name(name) => {
                // Map current code to glyph name -> Unicode
                let glyph_name = String::from_utf8_lossy(&name).to_string();
                if glyph_name == "fi"
                    || glyph_name == "fl"
                    || glyph_name == "ffi"
                    || glyph_name == "ffl"
                {
                    debug!(
                        "  Differences: code=0x{:02X} glyph={:?} (ligature)",
                        current_code, glyph_name
                    );
                    ligature_count += 1;
                }
                if let Some(ch) = glyph_to_char(&glyph_name) {
                    encoding_map.insert(current_code, ch);
                } else {
                    debug!(
                        "  Differences: code=0x{:02X} glyph={:?} (unmapped)",
                        current_code, glyph_name
                    );
                }
                current_code = current_code.wrapping_add(1);
            }
            _ => {}
        }
    }

    if ligature_count > 0 {
        debug!(
            "  Differences: {} total entries, {} ligatures",
            encoding_map.len(),
            ligature_count
        );
    }

    if encoding_map.is_empty() {
        None
    } else {
        Some(encoding_map)
    }
}

/// Get the CMap lookup key for an Identity-H/V CID font without ToUnicode.
/// Returns the object number used by `collect_cmaps_from_fonts` to store the CMap:
/// - FontFile2 or FontFile3 obj_num (for embedded font cmap)
/// - CIDFont dict obj_num (for predefined CIDSystemInfo-based mapping)
pub(crate) fn get_font_file2_obj_num(doc: &Document, font_dict: &lopdf::Dictionary) -> Option<u32> {
    let subtype = font_dict
        .get(b"Subtype")
        .ok()
        .and_then(|o| o.as_name().ok());

    // Type0 (CID) fonts
    if subtype == Some(b"Type0") {
        let encoding = font_dict.get(b"Encoding").ok()?.as_name().ok()?;
        if encoding != b"Identity-H" && encoding != b"Identity-V" {
            return None;
        }
        let desc_fonts_obj = font_dict.get(b"DescendantFonts").ok()?;
        let desc_fonts = resolve_array(doc, desc_fonts_obj)?;
        if desc_fonts.is_empty() {
            return None;
        }
        let cid_font_dict = resolve_dict(doc, &desc_fonts[0])?;
        let font_descriptor_obj = cid_font_dict.get(b"FontDescriptor").ok()?;
        let font_descriptor = resolve_dict(doc, font_descriptor_obj)?;

        // Try FontFile2 (TrueType), then FontFile3 (OpenType/CFF)
        if let Some(ff_ref) = font_descriptor
            .get(b"FontFile2")
            .ok()
            .and_then(|o| o.as_reference().ok())
            .or_else(|| {
                font_descriptor
                    .get(b"FontFile3")
                    .ok()
                    .and_then(|o| o.as_reference().ok())
            })
        {
            return Some(ff_ref.0);
        }

        // Fallback: use DescendantFonts[0] obj_num (for predefined CIDSystemInfo mapping)
        if let Object::Reference(r) = &desc_fonts[0] {
            return Some(r.0);
        }
        return None;
    }

    // Simple fonts: use embedded font file if available
    let font_descriptor_obj = font_dict.get(b"FontDescriptor").ok()?;
    let font_descriptor = resolve_dict(doc, font_descriptor_obj)?;
    font_descriptor
        .get(b"FontFile2")
        .ok()
        .and_then(|o| o.as_reference().ok())
        .or_else(|| {
            font_descriptor
                .get(b"FontFile3")
                .ok()
                .and_then(|o| o.as_reference().ok())
        })
        .map(|r| r.0)
}

/// Decode text from a PDF string operand using font CMaps, encodings, and fallbacks.
#[allow(clippy::too_many_arguments)]
pub(crate) fn extract_text_from_operand(
    obj: &Object,
    current_font: &str,
    base_font_name: Option<&str>,
    font_cmaps: &FontCMaps,
    font_tounicode_refs: &std::collections::HashMap<String, u32>,
    inline_cmaps: &std::collections::HashMap<String, crate::pdf_inspector::tounicode::CMapEntry>,
    font_encodings: &PageFontEncodings,
    encoding_cache: &HashMap<String, Encoding<'_>>,
    cmap_decisions: &mut CMapDecisionCache,
) -> Option<String> {
    let result = (|| -> Option<String> {
        if let Object::String(bytes, _) = obj {
            let mut decode_with_entry =
                |entry: &crate::pdf_inspector::tounicode::CMapEntry| -> Option<String> {
                    // For single-byte CMaps, merge CMap + Differences at the byte level:
                    // try CMap first, then Differences, then Latin-1 fallback per byte.
                    // This prevents partial CMap results from blocking the Differences path.
                    if entry.primary.code_byte_length == 1 {
                        let encoding_map = font_encodings.get(current_font);
                        let decoded: String = bytes
                            .iter()
                            .filter_map(|&b| {
                                let code = b as u16;
                                // 1. Primary CMap
                                if let Some(s) = entry.primary.lookup(code) {
                                    if !s.contains('\u{FFFD}') {
                                        return Some(s);
                                    }
                                }
                                // 2. Fallback CMap (embedded font cmap)
                                if let Some(fb) =
                                    entry.fallback.as_ref().and_then(|c| c.lookup(code))
                                {
                                    if !fb.contains('\u{FFFD}') {
                                        return Some(fb);
                                    }
                                }
                                // 3. Differences mapped it? Use Differences result
                                if let Some(map) = encoding_map {
                                    if let Some(&ch) = map.get(&b) {
                                        return Some(ch.to_string());
                                    }
                                }
                                // 4. Printable ASCII/Latin-1 fallback
                                if b >= 0x20 {
                                    return Some((b as char).to_string());
                                }
                                None
                            })
                            .collect();
                        if !decoded.is_empty() {
                            return Some(decoded);
                        }
                        return None;
                    }

                    // 2-byte CMap: use standard decode_cids path
                    if bytes.len() % 2 == 1 {
                        // Some PDFs emit 1-byte codes even for Type0 fonts; try per-byte lookup
                        let lookups = entry.primary.lookup_bytes(bytes);
                        let decoded: String = lookups
                            .iter()
                            .filter_map(|&(_b, ref cmap_result)| cmap_result.clone())
                            .collect();
                        if !decoded.is_empty() {
                            return Some(decoded);
                        }
                    }
                    let decoded_primary = entry.primary.decode_cids(bytes);
                    if let Some(remapped) = entry.remapped.as_ref() {
                        let decoded_remap = remapped.decode_cids(bytes);
                        let decoded_fallback =
                            entry.fallback.as_ref().map(|c| c.decode_cids(bytes));

                        if let Some(choice) = cmap_decisions
                            .get_choice(font_tounicode_refs.get(current_font).copied().unwrap_or(0))
                        {
                            let decoded = match choice {
                                CMapChoice::Primary => decoded_primary.clone(),
                                CMapChoice::Remapped => decoded_remap.clone(),
                            };
                            if !decoded.is_empty() {
                                return Some(decoded);
                            }
                        }

                        let choice = cmap_decisions.consider(
                            font_tounicode_refs.get(current_font).copied().unwrap_or(0),
                            &decoded_primary,
                            &decoded_remap,
                            bytes.len(),
                        );
                        let mut decoded = match choice {
                            Some(CMapChoice::Primary) => decoded_primary,
                            Some(CMapChoice::Remapped) => decoded_remap,
                            None => choose_best_cmap_decode(decoded_primary, decoded_remap),
                        };
                        if let Some(fb) = decoded_fallback {
                            let expected = bytes.len() / 2;
                            let decoded_len = decoded.chars().count();
                            let prefer_fallback = (!fb.is_empty() && decoded.is_empty())
                                || (!fb.is_empty() && expected > 0 && decoded_len * 2 < expected);
                            if prefer_fallback || score_text(&fb) > score_text(&decoded) + 3 {
                                decoded = fb;
                            }
                        }
                        if !decoded.is_empty() {
                            return Some(decoded);
                        }
                    } else if !decoded_primary.is_empty() {
                        if let Some(fb) = entry.fallback.as_ref().map(|c| c.decode_cids(bytes)) {
                            let expected = bytes.len() / 2;
                            let decoded_len = decoded_primary.chars().count();
                            let prefer_fallback = (!fb.is_empty() && decoded_primary.is_empty())
                                || (!fb.is_empty() && expected > 0 && decoded_len * 2 < expected);
                            if prefer_fallback || score_text(&fb) > score_text(&decoded_primary) + 3
                            {
                                return Some(fb);
                            }
                        }
                        return Some(decoded_primary);
                    }

                    None
                };

            if let Some(entry) = inline_cmaps.get(current_font) {
                if let Some(decoded) = decode_with_entry(entry) {
                    return Some(decoded);
                }
            }

            // Look up CMap by ToUnicode object reference
            if let Some(&obj_num) = font_tounicode_refs.get(current_font) {
                if let Some(entry) = font_cmaps.get_by_obj(obj_num) {
                    if let Some(decoded) = decode_with_entry(entry) {
                        return Some(decoded);
                    }
                }
            }

            // Try our custom encoding map from Differences arrays.
            // The Differences array overrides specific codes in a base encoding (typically
            // WinAnsiEncoding). We must combine Differences entries with the base encoding
            // rather than using filter_map which silently drops unmapped bytes.
            if let Some(encoding_map) = font_encodings.get(current_font) {
                let has_diff_match = bytes.iter().any(|b| encoding_map.contains_key(b));
                if has_diff_match {
                    let decoded: String = bytes
                        .iter()
                        .filter_map(|&b| {
                            if let Some(&ch) = encoding_map.get(&b) {
                                Some(ch)
                            } else if b >= 0x20 {
                                // Base encoding fallback for printable bytes.
                                // For codes 0x20-0x7E this matches all standard PDF encodings.
                                Some(b as char)
                            } else {
                                None // Skip unmapped control characters
                            }
                        })
                        .collect();
                    if !decoded.is_empty() {
                        return Some(decoded);
                    }
                }
            }

            // Fallback: try UTF-16BE then Latin-1
            if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
                let utf16: Vec<u16> = bytes[2..]
                    .chunks_exact(2)
                    .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
                    .collect();
                let text = String::from_utf16_lossy(&utf16);
                if text.contains('\u{FFFD}') {
                    debug!(
                        "utf16 loss produced replacement for font={} bytes_len={}",
                        current_font,
                        bytes.len()
                    );
                }
                return Some(text);
            }

            // Heuristic UTF-16BE decode when bytes look like UTF-16 (even length, null-heavy)
            if bytes.len() >= 4 && bytes.len() % 2 == 0 {
                let nulls = bytes.iter().filter(|&&b| b == 0).count();
                if nulls * 4 > bytes.len() {
                    let utf16: Vec<u16> = bytes
                        .chunks_exact(2)
                        .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
                        .collect();
                    let text = String::from_utf16_lossy(&utf16);
                    if score_text(&text) > 0 {
                        return Some(text);
                    }
                }
            }

            // Check for UTF-8 encoded strings before single-byte encoding decoding.
            // Some PDFs incorrectly embed UTF-8 bytes in single-byte encoded fonts
            // (e.g. "José" as UTF-8 [C3 A9] instead of WinAnsi [E9]).
            if bytes.iter().any(|&b| b > 0x7F) {
                if let Ok(text) = std::str::from_utf8(bytes) {
                    return Some(text.to_string());
                }
            }

            // Try to decode using cached font encoding from lopdf
            if let Some(encoding) = encoding_cache.get(current_font) {
                if let Ok(text) = Document::decode_text(encoding, bytes) {
                    if text.contains('\u{FFFD}') {
                        debug!(
                            "decode_text produced replacement for font={} bytes_len={}",
                            current_font,
                            bytes.len()
                        );
                        if bytes.len() <= 8 {
                            let hex: String = bytes.iter().map(|b| format!("{:02X}", b)).collect();
                            debug!(
                                "decode_text replacement bytes font={} base={:?} hex={}",
                                current_font, base_font_name, hex
                            );
                        }
                        if bytes.iter().all(|&b| (0x20..=0x7E).contains(&b)) {
                            return Some(bytes.iter().map(|&b| b as char).collect());
                        }
                        if let Some(symbol_text) = decode_symbol_fallback(bytes, base_font_name) {
                            return Some(symbol_text);
                        }
                    }
                    return Some(text);
                }
            }

            if let Some(symbol_text) = decode_symbol_fallback(bytes, base_font_name) {
                return Some(symbol_text);
            }

            // Latin-1 fallback
            Some(bytes.iter().map(|&b| b as char).collect())
        } else {
            None
        }
    })();
    result.map(clean_symbol_pua)
}

/// Replace PUA characters in the F000-F0FF range with standard Unicode equivalents.
/// These come from Symbol/Wingdings fonts whose ToUnicode CMaps map to PUA.
fn clean_symbol_pua(text: String) -> String {
    if !text.chars().any(|c| ('\u{F000}'..='\u{F0FF}').contains(&c)) {
        return text;
    }
    text.chars()
        .map(|c| {
            let code = c as u32;
            if !(0xF000..=0xF0FF).contains(&code) {
                return c;
            }
            let low = code - 0xF000;
            match low {
                // Common bullets
                0xA1 | 0xA7 | 0xB7 => '\u{2022}',
                // Checkmark
                0xFC => '\u{2713}',
                // Printable ASCII range and Latin-1 above: strip F000 offset
                0x20..=0xFF => char::from_u32(low).unwrap_or(c),
                _ => c,
            }
        })
        .collect()
}

fn decode_symbol_fallback(bytes: &[u8], base_font_name: Option<&str>) -> Option<String> {
    let name = base_font_name?.to_ascii_lowercase();
    if !name.contains("symbol") && !name.contains("wingdings") && !name.contains("zapfdingbats") {
        return None;
    }
    let mut out = String::new();
    for &b in bytes {
        if b < 0x20 {
            continue;
        }
        if let Some(ch) = char::from_u32(0xF000 + b as u32) {
            out.push(ch);
        }
    }
    if out.is_empty() { None } else { Some(out) }
}

fn choose_best_cmap_decode(primary: String, remapped: String) -> String {
    if primary.is_empty() {
        return remapped;
    }
    if remapped.is_empty() {
        return primary;
    }
    let score_primary = score_text(&primary);
    let score_remap = score_text(&remapped);
    if score_remap > score_primary + 3 {
        remapped
    } else {
        primary
    }
}

fn score_text(text: &str) -> i32 {
    const COMMON_WORDS: [&str; 22] = [
        "the", "and", "of", "to", "in", "a", "is", "that", "for", "with", "on", "as", "by", "from",
        "this", "be", "are", "at", "or", "not", "it", "our",
    ];

    let mut letters = 0i32;
    let mut spaces = 0i32;
    let mut digits = 0i32;
    let mut other = 0i32;
    let mut word_hits = 0i32;

    let mut current = String::new();
    for ch in text.chars() {
        if ch.is_ascii_alphabetic() {
            letters += 1;
            current.push(ch.to_ascii_lowercase());
        } else {
            if !current.is_empty() {
                if COMMON_WORDS.iter().any(|w| *w == current) {
                    word_hits += 1;
                }
                current.clear();
            }
            if ch == ' ' {
                spaces += 1;
            } else if ch.is_ascii_digit() {
                digits += 1;
            } else if ch.is_control() || ch == '\u{FFFD}' {
                other += 3;
            } else {
                other += 1;
            }
        }
    }
    if !current.is_empty() && COMMON_WORDS.iter().any(|w| *w == current) {
        word_hits += 1;
    }

    let mut score = word_hits * 10 + letters + spaces * 2 + digits - other * 2;
    if letters > 15 && word_hits == 0 {
        score -= 15;
    }
    score
}
