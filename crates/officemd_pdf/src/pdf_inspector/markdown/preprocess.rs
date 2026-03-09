//! Line preprocessing: heading merging and drop cap handling.

use crate::pdf_inspector::types::TextLine;

use super::analysis::detect_header_level;

/// Merge consecutive heading lines at the same level into a single line.
///
/// When a heading wraps across multiple text lines (e.g., "About Glenair, the Mission-Critical"
/// and "Interconnect Company"), each fragment becomes a separate `# Header` in the output.
/// This function detects consecutive lines at the same heading tier on the same page
/// with a small Y gap and merges them into one line.
pub(crate) fn merge_heading_lines(
    lines: Vec<TextLine>,
    base_size: f32,
    heading_tiers: &[f32],
) -> Vec<TextLine> {
    if lines.is_empty() {
        return lines;
    }

    let mut result: Vec<TextLine> = Vec::with_capacity(lines.len());

    for line in lines {
        let line_font = line.items.first().map(|i| i.font_size).unwrap_or(base_size);
        let line_level = detect_header_level(line_font, base_size, heading_tiers);

        // Check if the previous line is a heading at the same level on the same page
        let should_merge = if let (Some(prev), Some(curr_level)) = (result.last(), line_level) {
            let prev_font = prev.items.first().map(|i| i.font_size).unwrap_or(base_size);
            let prev_level = detect_header_level(prev_font, base_size, heading_tiers);
            let same_page = prev.page == line.page;
            let same_level = prev_level == Some(curr_level);
            let y_gap = prev.y - line.y;
            // Merge if gap is within ~2x the font size (normal line wrap spacing)
            let close_enough = y_gap > 0.0 && y_gap < line_font * 2.0;
            same_page && same_level && close_enough
        } else {
            false
        };

        if should_merge {
            // Append this line's items to the previous line
            let prev = result.last_mut().unwrap();
            // Add a space-bearing TextItem to separate the merged text
            if let Some(first_item) = line.items.first() {
                let mut space_item = first_item.clone();
                space_item.text = format!(" {}", space_item.text.trim_start());
                prev.items.push(space_item);
            }
            for item in line.items.into_iter().skip(1) {
                prev.items.push(item);
            }
        } else {
            result.push(line);
        }
    }

    result
}

/// Merge drop caps with the appropriate line.
/// A drop cap is a single large letter at the start of a paragraph.
/// Due to PDF coordinate sorting, the drop cap may appear AFTER the line it belongs to.
pub(crate) fn merge_drop_caps(lines: Vec<TextLine>, base_size: f32) -> Vec<TextLine> {
    let mut result: Vec<TextLine> = Vec::with_capacity(lines.len());

    for line in &lines {
        let text = line.text();
        let trimmed = text.trim();

        // Check if this looks like a drop cap:
        // 1. Single character (or single char + space)
        // 2. Much larger than base font (3x or more)
        // 3. The character is uppercase
        let is_drop_cap = trimmed.len() <= 2
            && line.items.first().map(|i| i.font_size).unwrap_or(0.0) >= base_size * 2.5
            && trimmed
                .chars()
                .next()
                .map(|c| c.is_uppercase())
                .unwrap_or(false);

        if is_drop_cap {
            let drop_char = trimmed.chars().next().unwrap();

            // Find the first line that starts with lowercase and is at the START of a paragraph
            // (i.e., preceded by a header or non-lowercase-starting line)
            let mut target_idx: Option<usize> = None;

            for (idx, prev_line) in result.iter().enumerate() {
                if prev_line.page != line.page {
                    continue;
                }

                let prev_text = prev_line.text();
                let prev_trimmed = prev_text.trim();

                // Check if this line starts with lowercase
                if prev_trimmed
                    .chars()
                    .next()
                    .map(|c| c.is_lowercase())
                    .unwrap_or(false)
                {
                    // Check if previous line exists and doesn't start with lowercase
                    // (meaning this is the start of a paragraph)
                    let is_para_start = if idx == 0 {
                        true
                    } else {
                        let before = result[idx - 1].text();
                        let before_trimmed = before.trim();
                        !before_trimmed
                            .chars()
                            .next()
                            .map(|c| c.is_lowercase())
                            .unwrap_or(true)
                    };

                    if is_para_start {
                        target_idx = Some(idx);
                        break;
                    }
                }
            }

            // Merge with the target line
            if let Some(idx) = target_idx {
                if let Some(first_item) = result[idx].items.first_mut() {
                    let prev_text = first_item.text.trim().to_string();
                    first_item.text = format!("{}{}", drop_char, prev_text);
                }
            }
            // Don't add the drop cap line itself
            continue;
        }

        result.push(line.clone());
    }

    result
}
