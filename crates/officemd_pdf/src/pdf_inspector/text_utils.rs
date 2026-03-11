//! Character classification and text utility functions.
//!
//! Pure helpers that operate on characters, strings, or `TextItem` slices.
//! No PDF parsing happens here — these are shared across the extraction
//! and markdown pipelines.

use crate::pdf_inspector::types::TextItem;

/// Check if a character is CJK (Chinese, Japanese, Korean).
/// CJK languages don't use spaces between words, so word-boundary
/// heuristics should not apply when CJK characters are involved.
pub(crate) fn is_cjk_char(c: char) -> bool {
    matches!(c,
        '\u{1100}'..='\u{11FF}'   // Hangul Jamo
        | '\u{3000}'..='\u{303F}' // CJK Symbols and Punctuation
        | '\u{3040}'..='\u{309F}' // Hiragana
        | '\u{30A0}'..='\u{30FF}' // Katakana
        | '\u{3130}'..='\u{318F}' // Hangul Compatibility Jamo
        | '\u{4E00}'..='\u{9FFF}' // CJK Unified Ideographs
        | '\u{AC00}'..='\u{D7AF}' // Hangul Syllables
        | '\u{F900}'..='\u{FAFF}' // CJK Compatibility Ideographs
        | '\u{FF00}'..='\u{FFEF}' // Halfwidth and Fullwidth Forms
    )
}

pub(crate) fn is_rtl_char(c: char) -> bool {
    matches!(c,
        '\u{0590}'..='\u{05FF}'   // Hebrew
        | '\u{0600}'..='\u{06FF}' // Arabic
        | '\u{0700}'..='\u{074F}' // Syriac
        | '\u{0750}'..='\u{077F}' // Arabic Supplement
        | '\u{0780}'..='\u{07BF}' // Thaana
        | '\u{07C0}'..='\u{07FF}' // NKo
        | '\u{0800}'..='\u{083F}' // Samaritan
        | '\u{0840}'..='\u{085F}' // Mandaic
        | '\u{08A0}'..='\u{08FF}' // Arabic Extended-A
        | '\u{FB1D}'..='\u{FB4F}' // Hebrew Presentation Forms
        | '\u{FB50}'..='\u{FDFF}' // Arabic Presentation Forms-A
        | '\u{FE70}'..='\u{FEFF}' // Arabic Presentation Forms-B
    )
}

pub(crate) fn is_rtl_text<I, S>(texts: I) -> bool
where
    I: Iterator<Item = S>,
    S: AsRef<str>,
{
    let (mut rtl, mut ltr) = (0u32, 0u32);
    for t in texts {
        for c in t.as_ref().chars() {
            if is_rtl_char(c) {
                rtl += 1;
            } else if c.is_alphabetic() && !is_cjk_char(c) {
                ltr += 1;
            }
        }
    }
    rtl > 0 && rtl > ltr
}

pub(crate) fn sort_line_items(items: &mut [TextItem]) {
    let rtl = is_rtl_text(items.iter().map(|i| &i.text));
    if rtl {
        items.sort_by(|a, b| b.x.partial_cmp(&a.x).unwrap_or(std::cmp::Ordering::Equal));
    } else {
        items.sort_by(|a, b| a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal));
    }
}

/// Detect if a font name indicates bold style
/// Common patterns: "Bold", "Bd", "Black", "Heavy", "Demi", "Semi" (semi-bold)
pub fn is_bold_font(font_name: &str) -> bool {
    let lower = font_name.to_lowercase();

    // Check for common bold indicators
    // Note: Need to be careful with "Oblique" not matching "Obl" + false positive for bold
    lower.contains("bold")
        || lower.contains("-bd")
        || lower.contains("_bd")
        || lower.contains("black")
        || lower.contains("heavy")
        || lower.contains("demibold")
        || lower.contains("semibold")
        || lower.contains("demi-bold")
        || lower.contains("semi-bold")
        || lower.contains("extrabold")
        || lower.contains("ultrabold")
        || lower.contains("medium") && !lower.contains("mediumitalic") // Some fonts use Medium for semi-bold
}

/// Detect if a font name indicates italic/oblique style
/// Common patterns: "Italic", "It", "Oblique", "Obl", "Slant", "Inclined"
pub fn is_italic_font(font_name: &str) -> bool {
    let lower = font_name.to_lowercase();

    // Check for common italic indicators
    lower.contains("italic")
        || lower.contains("oblique")
        || lower.contains("-it")
        || lower.contains("_it")
        || lower.contains("slant")
        || lower.contains("inclined")
        || lower.contains("kursiv") // German for italic
}

/// Expand Unicode ligature characters to their component characters.
/// This makes extracted text more searchable and semantically correct.
pub(crate) fn expand_ligatures(text: &str) -> String {
    // Strip null bytes and other control characters (except newline/tab)
    let text = if text
        .bytes()
        .any(|b| b < 0x20 && b != b'\n' && b != b'\r' && b != b'\t')
    {
        text.chars()
            .filter(|&c| c >= ' ' || c == '\n' || c == '\r' || c == '\t')
            .collect::<String>()
    } else {
        text.to_string()
    };

    let mut result = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '\u{FB00}' => result.push_str("ff"),
            '\u{FB01}' => result.push_str("fi"),
            '\u{FB02}' => result.push_str("fl"),
            '\u{FB03}' => result.push_str("ffi"),
            '\u{FB04}' => result.push_str("ffl"),
            '\u{FB05}' | '\u{FB06}' => result.push_str("st"),
            // Strip invisible Unicode characters that pollute markdown output
            '\u{00AD}' => {}              // soft hyphen
            '\u{200B}' => {}              // zero-width space
            '\u{FEFF}' => {}              // BOM / zero-width no-break space
            '\u{200C}' | '\u{200D}' => {} // ZWNJ / ZWJ
            '\u{2060}' => {}              // word joiner
            // Normalize typographic spaces to ASCII space so downstream
            // spacing logic (should_join_items) can detect word boundaries.
            // Excludes NBSP (U+00A0) which is common in PDFs and handled
            // correctly by existing coordinate-based spacing.
            '\u{2000}'..='\u{200A}' => result.push(' '), // en/em/thin/hair spaces etc.
            _ => result.push(ch),
        }
    }
    collapse_doubled_glyph_runs(&result)
}

fn same_letter_fold(a: char, b: char) -> bool {
    a.to_lowercase().to_string() == b.to_lowercase().to_string()
}

fn collapse_doubled_token(token: &str) -> Option<String> {
    let chars: Vec<char> = token.chars().collect();
    if chars.len() < 6 || !chars.iter().all(|c| c.is_alphabetic()) {
        return None;
    }

    let mut collapsed = String::with_capacity(chars.len());
    let mut duplicate_pairs = 0usize;
    let mut singles = 0usize;
    let mut i = 0usize;

    while i < chars.len() {
        let ch = chars[i];
        if i + 1 < chars.len() && same_letter_fold(ch, chars[i + 1]) {
            collapsed.push(ch);
            duplicate_pairs += 1;
            i += 2;
        } else {
            collapsed.push(ch);
            singles += 1;
            i += 1;
        }
    }

    let collapsed_len = collapsed.chars().count();
    let unit_count = duplicate_pairs + singles;
    if collapsed_len < 4 || duplicate_pairs < 2 || unit_count == 0 {
        return None;
    }

    let duplicate_ratio = duplicate_pairs as f32 / unit_count as f32;
    (duplicate_ratio >= 0.60).then_some(collapsed)
}

fn collapse_doubled_glyph_runs(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    let mut token = String::new();

    let flush_token = |token: &mut String, result: &mut String| {
        if token.is_empty() {
            return;
        }
        if let Some(collapsed) = collapse_doubled_token(token) {
            result.push_str(&collapsed);
        } else {
            result.push_str(token);
        }
        token.clear();
    };

    for ch in text.chars() {
        if ch.is_alphabetic() {
            token.push(ch);
        } else {
            flush_token(&mut token, &mut result);
            result.push(ch);
        }
    }
    flush_token(&mut token, &mut result);

    result
}

/// Decode a PDF text string (ActualText, etc.) that may be UTF-16BE (BOM \xFE\xFF)
/// or PDFDocEncoding (Latin-1 superset).
pub(crate) fn decode_text_string(bytes: &[u8]) -> String {
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        // UTF-16BE with BOM
        let utf16: Vec<u16> = bytes[2..]
            .chunks_exact(2)
            .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16_lossy(&utf16)
    } else {
        // PDFDocEncoding — identical to Latin-1 for the byte range we care about
        bytes.iter().map(|&b| b as char).collect()
    }
}

/// Compute effective font size from base size and text matrix
/// Text matrix is [a, b, c, d, tx, ty] where a,d are scale factors
pub(crate) fn effective_font_size(base_size: f32, text_matrix: &[f32; 6]) -> f32 {
    // The scale factor is typically the magnitude of the transformation
    // For most PDFs, text_matrix[0] (a) is the horizontal scale
    // and text_matrix[3] (d) is the vertical scale
    let scale_x = (text_matrix[0].powi(2) + text_matrix[1].powi(2)).sqrt();
    let scale_y = (text_matrix[2].powi(2) + text_matrix[3].powi(2)).sqrt();
    // Use the larger of the two scales (usually they're equal for non-rotated text)
    let scale = scale_x.max(scale_y);
    base_size * scale
}

/// Estimate the width of a text item, falling back to a character-count heuristic when width is 0.
pub(crate) fn effective_width(item: &TextItem) -> f32 {
    if item.width > 0.0 {
        item.width
    } else {
        item.text.chars().count() as f32 * item.font_size * 0.5
    }
}

pub(crate) fn is_cid_font(font: &str) -> bool {
    font.starts_with("C2_") || font.starts_with("C0_")
}

/// Determine if two adjacent text items should be joined without a space
/// based on their physical positions on the page and character case.
/// Uses a hybrid approach: position-based with case-aware thresholds.
/// CID fonts emit one word per text operator with gaps ≈ 0 between words.
/// Non-CID (Type1/TrueType) fonts emit phrases or fragments.
pub(crate) fn should_join_items(prev_item: &TextItem, curr_item: &TextItem) -> bool {
    // If either text explicitly has leading/trailing spaces, respect them
    if prev_item.text.ends_with(' ') || curr_item.text.starts_with(' ') {
        return false;
    }

    // Get the last character of previous and first character of current
    let prev_last = prev_item.text.trim_end().chars().last();
    let curr_first = curr_item.text.trim_start().chars().next();

    // Always join if current starts with punctuation that typically follows without space
    // e.g., "www" + ".com" → "www.com", not "www .com"
    if let Some(c) = curr_first {
        if matches!(c, '.' | ',' | ';' | '!' | '?' | ')' | ']' | '}' | '\'') {
            return true;
        }
    }

    // After colons, add space if followed by alphanumeric (typical label:value pattern)
    // e.g., "Clave:" + "T9N2I6" → "Clave: T9N2I6"
    if let (Some(p), Some(c)) = (prev_last, curr_first) {
        if p == ':' && c.is_alphanumeric() {
            return false;
        }
    }

    // When we have accurate width from font metrics, use a tight threshold
    if prev_item.width > 0.0 {
        let gap = if prev_item.x <= curr_item.x {
            // LTR: prev is left of curr
            curr_item.x - (prev_item.x + prev_item.width)
        } else {
            // RTL: prev is right of curr
            prev_item.x - (curr_item.x + curr_item.width)
        };
        let font_size = prev_item.font_size;

        // Never join across column-scale gaps
        if gap > font_size * 3.0 {
            return false;
        }

        // CID fonts (C2_*, C0_*) emit one word per text operator with gaps ≈ 0
        // between words. Detect these and add spaces. Only applies to CID fonts —
        // non-CID fonts (Type1/TrueType) emit phrases or fragments with small gaps
        // from positioning imprecision and should NOT trigger this.
        // Skip for CJK text — CJK languages don't use spaces between words.
        let prev_chars = prev_item.text.trim().chars().count();
        let curr_chars = curr_item.text.trim().chars().count();
        let prev_last_char = prev_item.text.trim().chars().last();
        let curr_first_char = curr_item.text.trim().chars().next();
        let is_cjk =
            prev_last_char.is_some_and(is_cjk_char) || curr_first_char.is_some_and(is_cjk_char);

        if !is_cjk && gap >= 0.0 && gap < font_size * 0.01 && is_cid_font(&prev_item.font) {
            let prev_word_count = prev_item.text.split_whitespace().count();

            if prev_word_count >= 3 {
                // Multi-word phrase from a line-level CID operator — likely mid-word boundary
                return gap < font_size * 0.15;
            }

            // CID font: each text operator is a separate word. Always add space.
            return false;
        }

        // Numeric continuity: digits, commas, periods, and percent signs that
        // are positioned close together are almost always a single number.
        // e.g., "34,20" + "8" → "34,208", "+13." + "0" + "%" → "+13.0%"
        // Use a generous threshold since word spaces in numbers are rare.
        if let (Some(p), Some(c)) = (prev_last, curr_first) {
            let prev_is_numeric = p.is_ascii_digit() || p == ',' || p == '.';
            let curr_is_numeric = c.is_ascii_digit() || c == '%' || c == '.';
            if prev_is_numeric && curr_is_numeric {
                return gap < font_size * 0.3;
            }
            // Sign characters (+/-) followed by digits
            if (p == '+' || p == '-') && c.is_ascii_digit() {
                return gap < font_size * 0.3;
            }
        }

        // Single-character fragment joined to a multi-character item: use a
        // moderately generous threshold to rejoin split words like "b" + "illion"
        // or "C" + "ultural". Gap near 0 = same word; gap ~0.2+ = different words.
        if (prev_chars == 1) != (curr_chars == 1) {
            return gap < font_size * 0.20;
        }

        // Both single-char: per-glyph positioning (character-by-character rendering).
        // Intra-word gaps are ≈ 0, word boundaries are ≈ 0.15× font_size.
        // For numeric chars (digits within "100,000"), use generous threshold.
        // For alphabetic, use tight threshold (0.10) to reliably detect word
        // boundaries in per-character PDFs like SEC filings.
        if prev_chars == 1 && curr_chars == 1 {
            if let (Some(p), Some(c)) = (prev_last, curr_first) {
                let p_numeric = p.is_ascii_digit() || matches!(p, ',' | '.' | '%' | '+' | '-');
                let c_numeric = c.is_ascii_digit() || matches!(c, ',' | '.' | '%');
                if p_numeric && c_numeric {
                    return gap < font_size * 0.25;
                }
            }
            return gap < font_size * 0.10;
        }

        // With accurate widths, a gap < 15% of font size means glyphs are
        // adjacent (same word). Anything larger is a deliberate space.
        // For multi-char items with a lowercase→lowercase junction, use a
        // slightly wider threshold (0.18) to avoid mid-word space injection
        // with imprecise CID font metrics (e.g. "enterta"+"inment").
        // All-caps or mixed-case junctions keep the tighter 0.15 threshold
        // to preserve word boundaries (e.g. "LCOE"+"WITH").
        if prev_item.text.trim().chars().count() >= 2 && curr_item.text.trim().chars().count() >= 2
        {
            let prev_ends_lower = prev_item
                .text
                .trim()
                .chars()
                .last()
                .is_some_and(|c| c.is_lowercase());
            let curr_starts_lower = curr_item
                .text
                .trim()
                .chars()
                .next()
                .is_some_and(|c| c.is_lowercase());
            if prev_ends_lower && curr_starts_lower {
                return gap < font_size * 0.18;
            }
        }
        return gap < font_size * 0.15;
    }

    // Fallback: estimate width from font size heuristics
    let char_width = prev_item.font_size * 0.45;

    let prev_text_len = prev_item.text.chars().count() as f32;
    let estimated_prev_width = prev_text_len * char_width;

    // Calculate expected end position of previous item
    let prev_end_x = prev_item.x + estimated_prev_width;

    // Calculate gap between items
    let gap = curr_item.x - prev_end_x;

    // Never join across column-scale gaps (fallback path)
    if gap > char_width * 6.0 {
        return false;
    }

    // CJK text: always join adjacent items — CJK languages don't use spaces between words.
    // The Latin case-based heuristics below would incorrectly insert spaces within CJK words.
    let is_cjk = prev_last.is_some_and(is_cjk_char) || curr_first.is_some_and(is_cjk_char);
    if is_cjk {
        return gap < char_width * 0.8;
    }

    // Use different thresholds based on character case
    // Same-case sequences (ALL CAPS or all lowercase) are more likely to be
    // word fragments that got split. Mixed case suggests word boundaries.
    match (prev_last, curr_first) {
        (Some(p), Some(c)) if p.is_alphabetic() && c.is_alphabetic() => {
            let same_case =
                (p.is_uppercase() && c.is_uppercase()) || (p.is_lowercase() && c.is_lowercase());
            if same_case {
                // Same case: use generous threshold (likely same word fragment)
                // e.g., "CONST" + "ANCIA" → "CONSTANCIA"
                gap < char_width * 0.8
            } else if p.is_lowercase() && c.is_uppercase() {
                // Lowercase to uppercase transition (e.g., "presente" → "CONSTANCIA")
                // This is typically a word boundary. In Spanish/English, words don't
                // transition from lowercase to uppercase mid-word.
                // Always add a space for this case, regardless of position.
                false
            } else {
                // Uppercase to lowercase (e.g., "REGISTRO" → "para")
                // Use stricter threshold (likely word boundary)
                gap < char_width * 0.3
            }
        }
        _ => {
            // Non-alphabetic: use moderate threshold
            gap < char_width * 0.5
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strip_soft_hyphen() {
        assert_eq!(expand_ligatures("con\u{00AD}tent"), "content");
    }

    #[test]
    fn strip_zero_width_space() {
        assert_eq!(expand_ligatures("hello\u{200B}world"), "helloworld");
    }

    #[test]
    fn strip_bom() {
        assert_eq!(expand_ligatures("\u{FEFF}text"), "text");
    }

    #[test]
    fn strip_zwnj_zwj_word_joiner() {
        assert_eq!(expand_ligatures("a\u{200C}b\u{200D}c\u{2060}d"), "abcd");
    }

    #[test]
    fn ligature_plus_invisible_chars() {
        assert_eq!(expand_ligatures("\u{FB01}rst\u{00AD}ly"), "firstly");
    }

    #[test]
    fn ligatures_still_expand() {
        assert_eq!(expand_ligatures("\u{FB00}\u{FB01}\u{FB02}"), "fffifl");
    }

    #[test]
    fn collapse_doubled_glyph_runs_on_stutter_token() {
        assert_eq!(expand_ligatures("Exxeerrcciiccee"), "Exercice");
    }

    #[test]
    fn keep_normal_double_letters() {
        assert_eq!(expand_ligatures("bookkeeper"), "bookkeeper");
        assert_eq!(expand_ligatures("committee"), "committee");
    }

    #[test]
    fn normalize_typographic_spaces() {
        // EM SPACE, EN SPACE, THIN SPACE → ASCII space
        assert_eq!(expand_ligatures("•\u{2003}text"), "• text");
        assert_eq!(expand_ligatures("a\u{2002}b"), "a b");
        assert_eq!(expand_ligatures("x\u{2009}y"), "x y");
    }

    #[test]
    fn nbsp_preserved() {
        // NBSP (U+00A0) should NOT be normalized
        assert_eq!(expand_ligatures("a\u{00A0}b"), "a\u{00A0}b");
    }
}
