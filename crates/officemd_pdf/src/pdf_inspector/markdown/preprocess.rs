//! Line preprocessing: heading merging, drop cap handling, and header/footer stripping.

use std::collections::{HashMap, HashSet};

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

/// Strip repeated headers and footers that appear across multiple pages.
///
/// Detects lines near the top/bottom of pages (within 8% of Y range) that
/// repeat on >= 50% of pages (minimum 3 pages). Text is normalized by
/// removing digit sequences (page numbers vary across pages) before comparison.
/// Lines with heading-tier font sizes are preserved.
pub(crate) fn strip_repeated_headers_footers(
    lines: Vec<TextLine>,
    base_size: f32,
    heading_tiers: &[f32],
) -> Vec<TextLine> {
    let pages: HashSet<u32> = lines.iter().map(|l| l.page).collect();
    let page_count = pages.len();
    if page_count < 3 {
        return lines;
    }

    // Compute Y bounds per page from text content
    let mut page_y_min: HashMap<u32, f32> = HashMap::new();
    let mut page_y_max: HashMap<u32, f32> = HashMap::new();
    for line in &lines {
        page_y_min
            .entry(line.page)
            .and_modify(|v| *v = v.min(line.y))
            .or_insert(line.y);
        page_y_max
            .entry(line.page)
            .and_modify(|v| *v = v.max(line.y))
            .or_insert(line.y);
    }

    let margin_fraction = 0.08;

    // Collect candidate header/footer texts with their page sets and Y positions
    struct Candidate {
        pages: HashSet<u32>,
        y_positions: Vec<f32>,
    }
    let mut candidates: HashMap<String, Candidate> = HashMap::new();

    for line in &lines {
        let text = line.text();
        let trimmed = text.trim();

        // Skip empty or long lines (unlikely to be headers/footers)
        if trimmed.is_empty() || trimmed.chars().count() > 120 {
            continue;
        }

        // Skip lines with heading-tier font sizes
        if let Some(first_item) = line.items.first() {
            if first_item.font_size / base_size >= 1.2
                && heading_tiers
                    .iter()
                    .any(|&t| (first_item.font_size - t).abs() < 0.5)
            {
                continue;
            }
        }

        // Check if line is in header or footer margin region
        let y_min = page_y_min.get(&line.page).copied().unwrap_or(0.0);
        let y_max = page_y_max.get(&line.page).copied().unwrap_or(842.0);
        let y_range = (y_max - y_min).max(1.0);

        let in_header_region = line.y >= y_max - y_range * margin_fraction;
        let in_footer_region = line.y <= y_min + y_range * margin_fraction;

        if !in_header_region && !in_footer_region {
            continue;
        }

        let normalized = normalize_header_footer_text(trimmed);
        if normalized.is_empty() {
            continue;
        }

        let entry = candidates.entry(normalized).or_insert_with(|| Candidate {
            pages: HashSet::new(),
            y_positions: Vec::new(),
        });
        entry.pages.insert(line.page);
        entry.y_positions.push(line.y);
    }

    // Determine which texts to strip: appear on >= 50% of pages with stable Y
    let min_pages = (page_count as f32 * 0.5).ceil() as usize;
    let y_tolerance = 5.0;
    let mut texts_to_strip: HashSet<String> = HashSet::new();

    for (text, candidate) in &candidates {
        if candidate.pages.len() < min_pages {
            continue;
        }
        // Check Y-position stability across pages
        if candidate.y_positions.len() >= 2 {
            let mean_y: f32 =
                candidate.y_positions.iter().sum::<f32>() / candidate.y_positions.len() as f32;
            if candidate
                .y_positions
                .iter()
                .all(|&y| (y - mean_y).abs() <= y_tolerance)
            {
                texts_to_strip.insert(text.clone());
            }
        }
    }

    if texts_to_strip.is_empty() {
        return lines;
    }

    lines
        .into_iter()
        .filter(|line| {
            let text = line.text();
            let trimmed = text.trim();
            if trimmed.is_empty() {
                return true;
            }
            let normalized = normalize_header_footer_text(trimmed);
            !texts_to_strip.contains(&normalized)
        })
        .collect()
}

/// Normalize header/footer text for cross-page comparison.
/// Strips all digit sequences (page numbers vary across pages) and lowercases.
fn normalize_header_footer_text(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    for ch in text.chars() {
        if !ch.is_ascii_digit() {
            result.push(ch);
        }
    }
    result
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .to_lowercase()
}
