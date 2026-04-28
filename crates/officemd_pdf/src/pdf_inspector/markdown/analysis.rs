//! Font statistics, heading detection, and document structure analysis.

use std::collections::HashMap;

use crate::pdf_inspector::types::{TextItem, TextLine};
use log::debug;

// ── Page margins & alignment ────────────────────────────────────────

/// Page margin info computed from bulk text statistics.
#[derive(Debug, Clone)]
pub(crate) struct PageMargins {
    /// Typical left edge of text (mode of line-start X, bucketed to 5pt).
    pub left: f32,
    /// Typical right edge of text (90th percentile of line-end X).
    pub right: f32,
    /// Midpoint of the text area.
    pub center: f32,
    /// Width of the text area (right − left).
    pub width: f32,
    /// Fraction of qualifying lines that start within ±10pt of the mode left margin.
    /// When most lines share the same left margin (ratio > 0.6), a line that is
    /// substantially to the right is a genuine indented block. When lines are
    /// scattered (ratio < 0.4), there is no dominant margin to compare against.
    pub left_margin_concentration: f32,
    /// Whether the page has multiple strong line-start clusters that suggest
    /// a multi-column layout rather than a single dominant text flow.
    pub likely_multi_column: bool,
}

/// Alignment classification for a text line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum LineAlignment {
    Left,
    Center,
}

/// Compute typical page margins from grouped text lines.
///
/// For each page, uses the **mode** (most common) left-margin bucket to define
/// the normal text indent, and the 90th-percentile right edge for the right
/// margin. The mode is more robust than percentiles because in many PDFs the
/// body text has a consistent indent while only a few lines (headings, bullets)
/// start further left. Bucketing to the nearest 5pt handles minor X-jitter.
pub(crate) fn compute_page_margins(lines: &[TextLine]) -> HashMap<u32, PageMargins> {
    let mut page_starts: HashMap<u32, Vec<f32>> = HashMap::new();
    let mut page_ends: HashMap<u32, Vec<f32>> = HashMap::new();

    for line in lines {
        // Skip very short lines (labels, page numbers) that would skew margins
        if line.items.len() < 3 {
            continue;
        }
        let first = &line.items[0];
        let last = line.items.last().unwrap();
        let line_start = first.x;
        let line_end = last.x + last.width;

        page_starts.entry(line.page).or_default().push(line_start);
        page_ends.entry(line.page).or_default().push(line_end);
    }

    let mut result = HashMap::new();
    for (&page, starts) in &mut page_starts {
        let ends = match page_ends.get_mut(&page) {
            Some(e) => e,
            None => continue,
        };
        if starts.len() < 3 {
            continue;
        }

        // Left margin: mode of X-start positions bucketed to nearest 5pt
        let (left, left_margin_concentration, likely_multi_column) = {
            let mut buckets: HashMap<i32, usize> = HashMap::new();
            for &x in starts.iter() {
                let key = (x / 5.0).round() as i32;
                *buckets.entry(key).or_default() += 1;
            }
            let (best_key, best_count) = buckets
                .iter()
                .max_by(|a, b| a.1.cmp(b.1).then_with(|| b.0.cmp(a.0)))
                .map(|(&k, &c)| (k, c))
                .unwrap_or((0, 0));
            let mode_x = best_key as f32 * 5.0;
            // Count lines within ±10pt of the mode (2 adjacent buckets)
            let near_count: usize = buckets
                .iter()
                .filter(|(k, _)| (*k - best_key).abs() <= 2)
                .map(|(_, &c)| c)
                .sum();
            let concentration = if starts.is_empty() {
                0.0
            } else {
                near_count as f32 / starts.len() as f32
            };
            let significant_threshold = ((starts.len() as f32) * 0.2).ceil() as usize;
            let mut significant_clusters: Vec<i32> = buckets
                .iter()
                .filter_map(|(&bucket, &count)| (count >= significant_threshold).then_some(bucket))
                .collect();
            significant_clusters.sort_unstable();
            let likely_multi_column = significant_clusters.windows(2).any(|pair| {
                let gap = (pair[1] - pair[0]) as f32 * 5.0;
                gap >= 120.0
            });
            let _ = best_count; // used indirectly via near_count
            (mode_x, concentration, likely_multi_column)
        };

        // Right margin: 90th percentile of X-end positions
        ends.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
        let right = ends[ends.len() * 9 / 10];
        let width = (right - left).max(1.0);

        result.insert(
            page,
            PageMargins {
                left,
                right,
                center: left + width / 2.0,
                width,
                left_margin_concentration,
                likely_multi_column,
            },
        );
    }

    result
}

#[cfg(test)]
mod tests {
    use super::compute_page_margins;
    use crate::pdf_inspector::types::{ItemType, TextItem, TextLine};

    fn make_line(page: u32, x: f32, y: f32, text: &str) -> TextLine {
        TextLine {
            items: vec![
                TextItem {
                    text: text.to_string(),
                    x,
                    y,
                    width: 80.0,
                    height: 12.0,
                    font: "F1".into(),
                    font_size: 12.0,
                    page,
                    is_bold: false,
                    is_italic: false,
                    item_type: ItemType::Text,
                    mcid: None,
                },
                TextItem {
                    text: "tail".into(),
                    x: x + 90.0,
                    y,
                    width: 30.0,
                    height: 12.0,
                    font: "F1".into(),
                    font_size: 12.0,
                    page,
                    is_bold: false,
                    is_italic: false,
                    item_type: ItemType::Text,
                    mcid: None,
                },
                TextItem {
                    text: "end".into(),
                    x: x + 130.0,
                    y,
                    width: 25.0,
                    height: 12.0,
                    font: "F1".into(),
                    font_size: 12.0,
                    page,
                    is_bold: false,
                    is_italic: false,
                    item_type: ItemType::Text,
                    mcid: None,
                },
            ],
            y,
            page,
            adaptive_threshold: 0.10,
        }
    }

    #[test]
    fn page_margins_mark_two_column_layouts() {
        let mut lines = Vec::new();
        for i in 0..12 {
            let y = 700.0 - i as f32 * 14.0;
            lines.push(make_line(1, 72.0, y, "Left column"));
            lines.push(make_line(1, 350.0, y, "Right column"));
        }

        let margins = compute_page_margins(&lines);
        assert!(margins.get(&1).unwrap().likely_multi_column);
    }

    #[test]
    fn page_margins_keep_single_column_layouts_enabled() {
        let mut lines = Vec::new();
        for i in 0..20 {
            let y = 700.0 - i as f32 * 14.0;
            lines.push(make_line(1, 72.0, y, "Body text"));
        }

        let margins = compute_page_margins(&lines);
        assert!(!margins.get(&1).unwrap().likely_multi_column);
    }
}

/// Detect whether a line is centered relative to page margins.
pub(crate) fn detect_line_alignment(line: &TextLine, margins: &PageMargins) -> LineAlignment {
    if line.items.is_empty() || margins.width < 50.0 {
        return LineAlignment::Left;
    }

    let first = &line.items[0];
    let last = line.items.last().unwrap();
    let line_start = first.x;
    let line_end = last.x + last.width;
    let line_width = line_end - line_start;
    let line_center = (line_start + line_end) / 2.0;

    // Line must be significantly shorter than the full text area
    if line_width > margins.width * 0.80 {
        return LineAlignment::Left;
    }

    // Line must not start near the normal left margin
    if line_start < margins.left + margins.width * 0.10 {
        return LineAlignment::Left;
    }

    // Line center must be close to page center
    let center_offset = (line_center - margins.center).abs();
    if center_offset > margins.width * 0.08 {
        return LineAlignment::Left;
    }

    // Limit to reasonably short text (avoid wrapped paragraphs)
    let char_count: usize = line.items.iter().map(|i| i.text.len()).sum();
    if char_count > 120 {
        return LineAlignment::Left;
    }

    LineAlignment::Center
}

/// Font statistics for a document
pub(crate) struct FontStats {
    pub(crate) most_common_size: f32,
    /// Font size frequency distribution (size_key → line count).
    /// Used for rarity-based heading detection.
    pub(crate) size_counts: HashMap<i32, usize>,
    /// Total number of lines counted.
    pub(crate) total_lines: usize,
}

/// Compute how rare a font size is in the document (0.0 = most common, 1.0 = unique).
/// Mirrors opendataloader's font rarity boosting approach: heading fonts appear on
/// far fewer lines than body text, so their percentile rank is high.
pub(crate) fn font_size_rarity(font_size: f32, stats: &FontStats) -> f32 {
    if stats.total_lines == 0 {
        return 0.0;
    }
    let key = (font_size * 10.0) as i32;
    let count = stats.size_counts.get(&key).copied().unwrap_or(0);
    // Rarity = 1 - (frequency ratio). A size used on 1/100 lines has rarity ~0.99.
    1.0 - (count as f32 / stats.total_lines as f32)
}

/// Calculate font stats directly from items (before grouping into lines)
pub(crate) fn calculate_font_stats_from_items(items: &[TextItem]) -> FontStats {
    let mut size_counts: HashMap<i32, usize> = HashMap::new();

    for item in items {
        if item.font_size >= 9.0 {
            let size_key = (item.font_size * 10.0) as i32;
            *size_counts.entry(size_key).or_insert(0) += 1;
        }
    }

    let total_lines = size_counts.values().sum();

    // Break ties by preferring the smaller font size for deterministic output
    let most_common_size = size_counts
        .iter()
        .max_by(|(size_a, count_a), (size_b, count_b)| {
            count_a.cmp(count_b).then_with(|| size_b.cmp(size_a))
        })
        .map(|(size, _)| *size as f32 / 10.0)
        .unwrap_or(12.0);

    FontStats {
        most_common_size,
        size_counts,
        total_lines,
    }
}

/// Calculate font stats from grouped lines
pub(crate) fn calculate_font_stats(lines: &[TextLine]) -> FontStats {
    let mut size_counts: HashMap<i32, usize> = HashMap::new();

    for line in lines {
        // Count once per line (first item) to give each line equal weight
        // Prevents small captions/footnotes from skewing the base
        if let Some(first) = line.items.first() {
            if first.font_size >= 9.0 {
                let size_key = (first.font_size * 10.0) as i32;
                *size_counts.entry(size_key).or_insert(0) += 1;
            }
        }
    }

    let total_lines = size_counts.values().sum();

    // Break ties by preferring the smaller font size for deterministic output
    let most_common_size = size_counts
        .iter()
        .max_by(|(size_a, count_a), (size_b, count_b)| {
            count_a.cmp(count_b).then_with(|| size_b.cmp(size_a))
        })
        .map(|(size, _)| *size as f32 / 10.0)
        .unwrap_or(12.0);

    FontStats {
        most_common_size,
        size_counts,
        total_lines,
    }
}

/// Determine the heading level for a bold-only line that didn't meet the font-size
/// threshold.  These are common in academic papers where section headings are bold
/// at the same size as body text.
///
/// Returns a level below the lowest font-size tier (or H2 when no tiers exist).
pub(crate) fn bold_heading_level(heading_tiers: &[f32]) -> usize {
    let level = heading_tiers.len() + 1;
    // Clamp to 1..=6 — if no font-size tiers, bold headings become H2
    // (H1 is reserved for titles which are typically larger)
    level.clamp(2, 6)
}

/// Detect TOC-style lines that contain dot leaders (e.g., "Section Name .... 42").
/// These lines should never be joined with adjacent lines into a paragraph.
/// Handles both consecutive dots ("....") and spaced dots ("...   ...").
pub(crate) fn has_dot_leaders(text: &str) -> bool {
    // Consecutive dots (4+)
    if text.contains("....") {
        return true;
    }
    // Spaced dot leaders: "..." followed by whitespace and more dots
    // Count occurrences of "..." (3+ dots) — if 2+ groups, it's a dot leader
    let mut dot_groups = 0;
    let mut consecutive_dots = 0;
    for ch in text.chars() {
        if ch == '.' {
            consecutive_dots += 1;
        } else {
            if consecutive_dots >= 3 {
                dot_groups += 1;
            }
            consecutive_dots = 0;
        }
    }
    if consecutive_dots >= 3 {
        dot_groups += 1;
    }
    dot_groups >= 2
}

/// Compute the Y-gap threshold for paragraph break detection.
///
/// Instead of using a fixed multiple of base_size (which fails for double-spaced
/// documents), we compute the document's typical (median) line spacing and use
/// a multiplier on that. A gap significantly larger than typical indicates a
/// paragraph break.
///
/// Fallback: if we can't compute typical spacing, use base_size * 1.8.
pub(crate) fn compute_paragraph_threshold(lines: &[TextLine], base_size: f32) -> f32 {
    let fallback = base_size * 1.8;

    // Collect Y gaps between consecutive lines on the same page
    let mut gaps: Vec<f32> = Vec::new();
    let mut prev_y: Option<(u32, f32)> = None;

    for line in lines {
        if let Some((prev_page, py)) = prev_y {
            if line.page == prev_page {
                let gap = py - line.y;
                // Only consider positive gaps within a reasonable range
                // (skip huge gaps from page headers/footers)
                if gap > 0.0 && gap < base_size * 10.0 {
                    gaps.push(gap);
                }
            }
        }
        prev_y = Some((line.page, line.y));
    }

    if gaps.len() < 5 {
        return fallback;
    }

    gaps.sort_by(|a, b| a.total_cmp(b));

    let median = gaps[gaps.len() / 2];

    let threshold = (median * 1.3).max(base_size * 1.5);

    debug!(
        "paragraph_threshold: base_size={:.1} median_gap={:.1} threshold={:.1} ({} gaps sampled)",
        base_size,
        median,
        threshold,
        gaps.len()
    );

    if log::log_enabled!(log::Level::Debug) {
        // Gap histogram
        let buckets: &[f32] = &[0.0, 0.5, 1.0, 1.2, 1.5, 1.8, 2.0, 2.5, 3.0, 5.0, 10.0];
        for i in 0..buckets.len() - 1 {
            let count = gaps
                .iter()
                .filter(|&&g| {
                    let r = g / base_size;
                    r >= buckets[i] && r < buckets[i + 1]
                })
                .count();
            if count > 0 {
                debug!(
                    "  gap_ratio {:.1}-{:.1}: {}",
                    buckets[i],
                    buckets[i + 1],
                    count
                );
            }
        }
        let over = gaps.iter().filter(|&&g| g / base_size >= 10.0).count();
        if over > 0 {
            debug!("  gap_ratio 10.0+: {}", over);
        }
    }

    // Per-line detail: Y position, gap, ratio, bold, text preview, paragraph marker
    if log::log_enabled!(log::Level::Trace) {
        let mut prev: Option<(u32, f32)> = None;
        for line in lines {
            let font_size = line.items.first().map(|i| i.font_size).unwrap_or(0.0);
            let is_bold = line.items.first().map(|i| i.is_bold).unwrap_or(false);
            let text = line.text();
            let display: String = text.chars().take(80).collect();

            let (gap_str, ratio_str, marker) = if let Some((pp, py)) = prev {
                if pp == line.page {
                    let gap = py - line.y;
                    let ratio = gap / base_size;
                    let is_para = gap > threshold;
                    (
                        format!("{:8.1}", gap),
                        format!("{:8.2}", ratio),
                        if is_para { " <<PARA>>" } else { "" },
                    )
                } else {
                    ("     ---".to_string(), "     ---".to_string(), "")
                }
            } else {
                ("     ---".to_string(), "     ---".to_string(), "")
            };

            log::trace!(
                "  p={} y={:8.1} gap={} ratio={} fs={:5.1} {}  {}{}",
                line.page,
                line.y,
                gap_str,
                ratio_str,
                font_size,
                if is_bold { "B" } else { " " },
                display,
                marker
            );

            prev = Some((line.page, line.y));
        }
    }

    threshold
}

/// Discover distinct heading font-size tiers in the document.
/// Returns tiers sorted largest-first (tier 0 = H1, tier 1 = H2, …).
/// Sizes within 0.5pt are clustered into the same tier. Capped at 4 tiers.
pub(crate) fn compute_heading_tiers(lines: &[TextLine], base_size: f32) -> Vec<f32> {
    let mut heading_sizes: Vec<f32> = Vec::new();

    for line in lines {
        if let Some(first) = line.items.first() {
            if first.font_size / base_size >= 1.2 {
                heading_sizes.push(first.font_size);
            }
        }
    }

    // Sort descending
    heading_sizes.sort_by(|a, b| b.total_cmp(a));

    // Cluster sizes within 0.5pt into same tier (use first value as representative)
    let mut tiers: Vec<f32> = Vec::new();
    for size in heading_sizes {
        let already_in_tier = tiers.iter().any(|&t| (t - size).abs() < 0.5);
        if !already_in_tier {
            tiers.push(size);
        }
    }

    // Cap at 4 tiers
    tiers.truncate(4);
    tiers
}

/// Detect header level from font size using document-specific heading tiers.
/// When tiers are available, maps tier 0→H1, tier 1→H2, etc.
/// Falls back to ratio-based thresholds when no tiers exist.
pub(crate) fn detect_header_level(
    font_size: f32,
    base_size: f32,
    heading_tiers: &[f32],
) -> Option<usize> {
    let ratio = font_size / base_size;

    if ratio < 1.2 {
        return None; // Regular text
    }

    if !heading_tiers.is_empty() {
        // Match font_size to a tier (within 0.5pt tolerance)
        for (i, &tier_size) in heading_tiers.iter().enumerate() {
            if (font_size - tier_size).abs() < 0.5 {
                return Some(i + 1); // tier 0 → H1, tier 1 → H2, etc.
            }
        }
        // No tier match but large ratio — assign level after last tier
        if ratio >= 1.5 {
            let level = (heading_tiers.len() + 1).min(4);
            return Some(level);
        }
        // No tier match and small ratio — not a heading
        return None;
    }

    // Fallback: original ratio-based thresholds (no tiers discovered)
    if ratio >= 2.0 {
        Some(1)
    } else if ratio >= 1.5 {
        Some(2)
    } else if ratio >= 1.25 {
        Some(3)
    } else {
        Some(4)
    }
}
