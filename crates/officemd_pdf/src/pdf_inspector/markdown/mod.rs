//! Markdown conversion with structure detection.
//!
//! Converts extracted text to markdown, detecting:
//! - Headers (by font size)
//! - Lists (bullet points, numbered lists)
//! - Code blocks (monospace fonts, indentation)
//! - Paragraphs

mod analysis;
mod classify;
mod convert;
mod postprocess;
mod preprocess;

pub use convert::to_markdown_from_lines;

use std::collections::{HashMap, HashSet};

use crate::pdf_inspector::extractor::group_into_lines;
use crate::pdf_inspector::types::{TextItem, TextLine};

use crate::pdf_inspector::tables::Table;
use analysis::calculate_font_stats_from_items;
use classify::{format_list_item, is_code_like, is_list_item};
use convert::{merge_continuation_tables, to_markdown_from_lines_with_tables_and_images};

fn item_in_hint_region(
    item: &TextItem,
    hint: &crate::pdf_inspector::tables::RectHintRegion,
    padding: f32,
) -> bool {
    item.x + item.width >= hint.x_left - padding
        && item.x <= hint.x_right + padding
        && item.y >= hint.y_bottom - padding
        && item.y <= hint.y_top + padding
}

fn merge_hint_regions_into_bands(
    hints: &[crate::pdf_inspector::tables::RectHintRegion],
) -> Vec<crate::pdf_inspector::tables::RectHintRegion> {
    if hints.len() < 2 {
        return hints.to_vec();
    }

    let mut sorted = hints.to_vec();
    sorted.sort_by(|a, b| {
        b.y_top
            .partial_cmp(&a.y_top)
            .unwrap_or(std::cmp::Ordering::Equal)
    });

    let mut merged: Vec<crate::pdf_inspector::tables::RectHintRegion> = Vec::new();
    for hint in sorted {
        let vertical_overlap =
            |a: &crate::pdf_inspector::tables::RectHintRegion,
             b: &crate::pdf_inspector::tables::RectHintRegion| {
                (a.y_top.min(b.y_top) - a.y_bottom.max(b.y_bottom)).max(0.0)
            };
        let horizontal_gap =
            |a: &crate::pdf_inspector::tables::RectHintRegion,
             b: &crate::pdf_inspector::tables::RectHintRegion| {
                if a.x_right < b.x_left {
                    b.x_left - a.x_right
                } else if b.x_right < a.x_left {
                    a.x_left - b.x_right
                } else {
                    0.0
                }
            };

        let should_merge = merged.last_mut().is_some_and(|prev| {
            let overlap = vertical_overlap(prev, &hint);
            let min_height = (prev.y_top - prev.y_bottom)
                .min(hint.y_top - hint.y_bottom)
                .max(1.0);
            let same_band = overlap >= min_height * 0.45
                || (prev.y_bottom - hint.y_top).abs() <= 14.0
                || (hint.y_bottom - prev.y_top).abs() <= 14.0;
            same_band && horizontal_gap(prev, &hint) <= 60.0
        });

        if should_merge {
            let prev = merged.last_mut().unwrap();
            prev.x_left = prev.x_left.min(hint.x_left);
            prev.x_right = prev.x_right.max(hint.x_right);
            prev.y_bottom = prev.y_bottom.min(hint.y_bottom);
            prev.y_top = prev.y_top.max(hint.y_top);
            prev.rect_count += hint.rect_count;
        } else {
            merged.push(hint);
        }
    }

    merged
}

/// Detect single uppercase characters with impossibly narrow width - typically
/// watermark or decorative fragments from PDF rendering artifacts.
fn is_probably_watermark_fragment(item: &TextItem) -> bool {
    let trimmed = item.text.trim();
    if trimmed.is_empty() {
        return false;
    }
    let char_count = trimmed.chars().count();
    if char_count != 1 {
        return false;
    }
    let width = item.width.abs();
    let font_size = item.font_size.max(1.0);
    width <= font_size * 0.35 && trimmed.chars().all(|c| c.is_ascii_uppercase())
}

fn is_probably_rotated_or_decorative(item: &TextItem) -> bool {
    if item.is_rotated {
        return true;
    }

    let trimmed = item.text.trim();
    if trimmed.is_empty() {
        return false;
    }

    let char_count = trimmed.chars().count();
    let alpha_count = trimmed.chars().filter(|c| c.is_alphabetic()).count();
    let width = item.width.abs();
    let font_size = item.font_size.max(1.0);
    let tiny_width = width <= font_size * 0.35;

    if char_count >= 2 {
        return tiny_width && alpha_count > 0;
    }

    is_probably_watermark_fragment(item)
}

fn is_probably_table_margin_noise(item: &TextItem) -> bool {
    let trimmed = item.text.trim();
    if trimmed.is_empty() {
        return false;
    }

    if trimmed == "SPNPRNRORR" {
        return true;
    }

    if trimmed.chars().count() == 1 && trimmed.chars().all(|c| c.is_ascii_uppercase()) {
        return item.font_size >= 16.0 || item.x < 120.0;
    }

    trimmed.len() >= 8
        && trimmed.chars().all(|c| c.is_ascii_uppercase())
        && !trimmed.contains(' ')
        && item.x < 220.0
}

fn salvage_boxed_section_table(page_items: &[TextItem], subset_indices: &[usize]) -> Option<Table> {
    let subset: Vec<(usize, &TextItem)> = subset_indices
        .iter()
        .copied()
        .filter_map(|idx| {
            let item = page_items.get(idx)?;
            (!is_probably_rotated_or_decorative(item)
                && !is_probably_table_margin_noise(item)
                && !item.text.trim().is_empty())
            .then_some((idx, item))
        })
        .collect();

    if subset.len() < 6 {
        return None;
    }

    let mut sorted_by_y = subset.clone();
    sorted_by_y.sort_by(|a, b| {
        b.1.y
            .partial_cmp(&a.1.y)
            .unwrap_or(std::cmp::Ordering::Equal)
            .then_with(|| {
                a.1.x
                    .partial_cmp(&b.1.x)
                    .unwrap_or(std::cmp::Ordering::Equal)
            })
    });

    let row_tolerance = sorted_by_y
        .iter()
        .map(|(_, item)| item.font_size)
        .fold(7.0f32, |acc, size| acc.max(size * 0.85))
        .clamp(7.0, 14.0);

    let mut row_groups: Vec<Vec<(usize, &TextItem)>> = Vec::new();
    let mut row_centers: Vec<f32> = Vec::new();
    for (idx, item) in sorted_by_y {
        if let Some((row_idx, _)) = row_centers
            .iter()
            .enumerate()
            .find(|(_, center)| (item.y - **center).abs() <= row_tolerance)
        {
            let group = &mut row_groups[row_idx];
            let count = group.len() as f32;
            row_centers[row_idx] = ((row_centers[row_idx] * count) + item.y) / (count + 1.0);
            group.push((idx, item));
        } else {
            row_centers.push(item.y);
            row_groups.push(vec![(idx, item)]);
        }
    }

    if row_groups.len() < 2 {
        return None;
    }

    for group in &mut row_groups {
        group.sort_by(|a, b| {
            a.1.x
                .partial_cmp(&b.1.x)
                .unwrap_or(std::cmp::Ordering::Equal)
        });
    }

    let mut x_positions: Vec<f32> = subset.iter().map(|(_, item)| item.x).collect();
    x_positions.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let x_range =
        x_positions.last().copied().unwrap_or(0.0) - x_positions.first().copied().unwrap_or(0.0);
    let avg_gap = if x_positions.len() > 1 {
        x_range / (x_positions.len() - 1) as f32
    } else {
        40.0
    };
    let col_tolerance = avg_gap.clamp(24.0, 70.0);

    let mut col_centers: Vec<f32> = Vec::new();
    let mut col_counts: Vec<usize> = Vec::new();
    for x in x_positions {
        if let Some((col_idx, _)) = col_centers
            .iter()
            .enumerate()
            .find(|(_, center)| (x - **center).abs() <= col_tolerance)
        {
            let count = col_counts[col_idx] as f32;
            col_centers[col_idx] = ((col_centers[col_idx] * count) + x) / (count + 1.0);
            col_counts[col_idx] += 1;
        } else {
            col_centers.push(x);
            col_counts.push(1);
        }
    }

    let min_items_per_col = (row_groups.len() / 3).max(2);
    let mut filtered_cols: Vec<f32> = col_centers
        .into_iter()
        .zip(col_counts)
        .filter_map(|(center, count)| (count >= min_items_per_col).then_some(center))
        .collect();
    filtered_cols.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    if filtered_cols.len() < 2 || filtered_cols.len() > 6 {
        return None;
    }

    let mut cells = vec![vec![String::new(); filtered_cols.len()]; row_groups.len()];
    let mut item_indices = Vec::new();
    let assignment_tolerance = col_tolerance * 1.35;

    for (row_idx, group) in row_groups.iter().enumerate() {
        for (item_idx, item) in group {
            let Some((col_idx, _)) = filtered_cols.iter().enumerate().min_by(|a, b| {
                (item.x - *a.1)
                    .abs()
                    .partial_cmp(&(item.x - *b.1).abs())
                    .unwrap_or(std::cmp::Ordering::Equal)
            }) else {
                continue;
            };

            if (item.x - filtered_cols[col_idx]).abs() > assignment_tolerance {
                continue;
            }

            let cell = &mut cells[row_idx][col_idx];
            if !cell.is_empty() {
                cell.push(' ');
            }
            cell.push_str(item.text.trim());
            item_indices.push(*item_idx);
        }
    }

    item_indices.sort_unstable();
    item_indices.dedup();

    let non_empty_rows = cells
        .iter()
        .filter(|row| row.iter().any(|cell| !cell.trim().is_empty()))
        .count();
    if non_empty_rows < 2 {
        return None;
    }

    let non_empty_cols = (0..filtered_cols.len())
        .filter(|&col| {
            cells
                .iter()
                .any(|row| row.get(col).is_some_and(|cell| !cell.trim().is_empty()))
        })
        .count();
    if non_empty_cols < 2 {
        return None;
    }

    let dense_rows = cells
        .iter()
        .filter(|row| row.iter().filter(|cell| !cell.trim().is_empty()).count() >= 2)
        .count();
    if dense_rows < 2 {
        return None;
    }

    let alpha_rows = cells
        .iter()
        .filter(|row| {
            row.first()
                .is_some_and(|cell| cell.chars().any(|c| c.is_alphabetic()))
                || row
                    .iter()
                    .any(|cell| cell.chars().filter(|c| c.is_alphabetic()).count() >= 6)
        })
        .count();
    if alpha_rows < 2 {
        return None;
    }

    Some(Table {
        columns: filtered_cols,
        rows: row_centers,
        cells,
        item_indices,
    })
}

#[derive(Default)]
struct PageContext {
    has_tabular_regions: bool,
    tables: Vec<(f32, String)>,
}

fn is_numeric_tick_token(token: &str) -> bool {
    let trimmed = token.trim_matches(|c: char| matches!(c, ',' | ';' | ':' | '.'));
    let trimmed = trimmed.strip_prefix('-').unwrap_or(trimmed);
    !trimmed.is_empty() && trimmed.chars().all(|c| c.is_ascii_digit())
}

fn is_year_marker_token(token: &str) -> bool {
    let trimmed = token.trim_matches(|c: char| matches!(c, ',' | ';' | ':' | '.'));
    trimmed.eq_ignore_ascii_case("exercice") || matches!(trimmed, "N" | "N-1" | "N-2" | "N-3" | "E")
}

fn is_probably_chart_noise_line(line: &TextLine) -> bool {
    let text = line.text();
    let tokens: Vec<&str> = text.split_whitespace().collect();
    if tokens.len() < 8 {
        return false;
    }

    let numeric_tokens = tokens.iter().filter(|t| is_numeric_tick_token(t)).count();
    let year_marker_tokens = tokens.iter().filter(|t| is_year_marker_token(t)).count();
    let alpha_tokens = tokens
        .iter()
        .filter(|token| token.chars().any(|c| c.is_alphabetic()))
        .count();

    (year_marker_tokens >= 2 && numeric_tokens >= 8) || (numeric_tokens >= 14 && alpha_tokens <= 6)
}

fn is_probably_chart_legend_line(line: &TextLine) -> bool {
    let text = line.text();
    let tokens: Vec<&str> = text.split_whitespace().collect();
    if tokens.len() < 5 {
        return false;
    }

    let numeric_tokens = tokens.iter().filter(|t| is_numeric_tick_token(t)).count();
    if numeric_tokens > 0 {
        return false;
    }

    let alpha_tokens = tokens
        .iter()
        .filter(|token| token.chars().any(|c| c.is_alphabetic()))
        .count();
    let short_or_abbrev_tokens = tokens
        .iter()
        .filter(|token| {
            let cleaned: String = token.chars().filter(|c| c.is_alphabetic()).collect();
            !cleaned.is_empty() && (cleaned.chars().count() <= 5 || token.contains('.'))
        })
        .count();

    alpha_tokens >= 5 && short_or_abbrev_tokens >= alpha_tokens.saturating_sub(1)
}

/// Options for markdown conversion
#[derive(Debug, Clone)]
pub struct MarkdownOptions {
    /// Detect headers by font size
    pub detect_headers: bool,
    /// Detect list items
    pub detect_lists: bool,
    /// Detect code blocks
    pub detect_code: bool,
    /// Base font size for comparison
    pub base_font_size: Option<f32>,
    /// Remove standalone page numbers
    pub remove_page_numbers: bool,
    /// Convert URLs to markdown links
    pub format_urls: bool,
    /// Fix hyphenation (broken words across lines)
    pub fix_hyphenation: bool,
    /// Detect and format bold text from font names
    pub detect_bold: bool,
    /// Detect and format italic text from font names
    pub detect_italic: bool,
    /// Include image placeholders in output
    pub include_images: bool,
    /// Include extracted hyperlinks
    pub include_links: bool,
    /// Insert page break markers (<!-- Page N -->) between pages
    pub include_page_numbers: bool,
    /// Detect centered text and wrap with `<center>` tags
    pub detect_alignment: bool,
    /// Detect indented blocks and render as `>` blockquotes
    pub detect_block_quotes: bool,
}

impl Default for MarkdownOptions {
    fn default() -> Self {
        Self {
            detect_headers: true,
            detect_lists: true,
            detect_code: true,
            base_font_size: None,
            remove_page_numbers: true,
            format_urls: true,
            fix_hyphenation: true,
            detect_bold: true,
            detect_italic: true,
            include_images: true,
            include_links: true,
            include_page_numbers: false,
            detect_alignment: true,
            detect_block_quotes: true,
        }
    }
}

/// Convert plain text to markdown (basic conversion)
pub fn to_markdown(text: &str, options: MarkdownOptions) -> String {
    let mut output = String::new();
    let mut in_list = false;
    let mut in_code_block = false;

    for line in text.lines() {
        let trimmed = line.trim();

        if trimmed.is_empty() {
            if in_list {
                in_list = false;
            }
            if in_code_block {
                output.push_str("```\n");
                in_code_block = false;
            }
            output.push('\n');
            continue;
        }

        // Detect list items
        if options.detect_lists && is_list_item(trimmed) {
            let formatted = format_list_item(trimmed);
            output.push_str(&formatted);
            output.push('\n');
            in_list = true;
            continue;
        }

        // Detect code blocks (indented lines)
        if options.detect_code && is_code_like(trimmed) {
            if !in_code_block {
                output.push_str("```\n");
                in_code_block = true;
            }
            output.push_str(trimmed);
            output.push('\n');
            continue;
        } else if in_code_block {
            output.push_str("```\n");
            in_code_block = false;
        }

        // Regular paragraph text
        output.push_str(trimmed);
        output.push('\n');
    }

    if in_code_block {
        output.push_str("```\n");
    }

    output
}

/// Convert positioned text items to markdown with structure detection
pub fn to_markdown_from_items(items: Vec<TextItem>, options: MarkdownOptions) -> String {
    to_markdown_from_items_with_rects(items, options, &[])
}

/// Convert positioned text items to markdown, using rectangle data for table detection
pub fn to_markdown_from_items_with_rects(
    items: Vec<TextItem>,
    options: MarkdownOptions,
    rects: &[crate::pdf_inspector::types::PdfRect],
) -> String {
    use crate::pdf_inspector::tables::{
        detect_tables, detect_tables_from_rects, table_to_markdown,
    };
    use crate::pdf_inspector::types::ItemType;

    if items.is_empty() {
        return String::new();
    }

    // Separate images and links from text items
    let mut images: Vec<TextItem> = Vec::new();
    let mut links: Vec<TextItem> = Vec::new();
    let mut text_items: Vec<TextItem> = Vec::new();

    for item in items {
        match &item.item_type {
            ItemType::Image => {
                if options.include_images {
                    images.push(item);
                }
            }
            ItemType::Link(_) => {
                if options.include_links {
                    links.push(item);
                }
            }
            ItemType::Text | ItemType::FormField => {
                text_items.push(item);
            }
        }
    }

    // Calculate base font size for table detection
    let font_stats = calculate_font_stats_from_items(&text_items);
    let base_size = options
        .base_font_size
        .unwrap_or(font_stats.most_common_size);

    // Detect tables on each page
    let mut table_items: HashSet<usize> = HashSet::new();
    let mut page_contexts: HashMap<u32, PageContext> = HashMap::new();

    // Store images by page and Y position for insertion
    let mut page_images: HashMap<u32, Vec<(f32, String)>> = HashMap::new();

    for img in &images {
        // Extract image name from "[Image: Im0]" format
        let img_name = img
            .text
            .strip_prefix("[Image: ")
            .and_then(|s| s.strip_suffix(']'))
            .unwrap_or(&img.text);
        let img_md = format!("![Image: {}](image)\n", img_name);
        page_images
            .entry(img.page)
            .or_default()
            .push((img.y, img_md));
    }

    // Pre-group items by page with their global indices (O(n) instead of O(pages*n))
    let mut page_groups: HashMap<u32, Vec<(usize, &TextItem)>> = HashMap::new();
    for (global_idx, item) in text_items.iter().enumerate() {
        page_groups
            .entry(item.page)
            .or_default()
            .push((global_idx, item));
    }

    let mut pages: Vec<u32> = page_groups.keys().copied().collect();
    pages.sort();

    for page in pages {
        let group = page_groups.get(&page).unwrap();
        let page_ctx = page_contexts.entry(page).or_default();
        let page_items: Vec<TextItem> = group.iter().map(|(_, item)| (*item).clone()).collect();
        let page_rect_count = rects.iter().filter(|rect| rect.page == page).count();

        // Track which local indices are claimed by rect-based tables
        let mut rect_claimed: HashSet<usize> = HashSet::new();

        // Try rectangle-based table detection first
        let (rect_tables, hint_regions) = detect_tables_from_rects(&page_items, rects, page);
        if !rect_tables.is_empty() || !hint_regions.is_empty() || page_rect_count >= 4 {
            page_ctx.has_tabular_regions = true;
        }
        for table in &rect_tables {
            for &idx in &table.item_indices {
                rect_claimed.insert(idx);
                if let Some(&(global_idx, _)) = group.get(idx) {
                    table_items.insert(global_idx);
                }
            }
            let table_y = table.rows.first().copied().unwrap_or(0.0);
            let table_md = table_to_markdown(table);
            page_ctx.tables.push((table_y, table_md));
        }

        // Helper: run heuristic on a subset of page-local indices, remapping indices back to page-space
        let mut run_heuristic =
            |subset_indices: &[usize], min_items: usize| -> Vec<(Table, Vec<usize>)> {
                let (filtered_items, filtered_map): (Vec<TextItem>, Vec<usize>) = subset_indices
                    .iter()
                    .copied()
                    .filter_map(|page_idx| {
                        let item = page_items.get(page_idx)?;
                        (!is_probably_rotated_or_decorative(item)).then(|| (item.clone(), page_idx))
                    })
                    .unzip();

                if filtered_items.len() < min_items {
                    return Vec::new();
                }

                detect_tables(&filtered_items, base_size, false)
                    .into_iter()
                    .map(|table| {
                        let mut claimed_page_indices: Vec<usize> = table
                            .item_indices
                            .iter()
                            .filter_map(|&idx| filtered_map.get(idx).copied())
                            .collect();
                        claimed_page_indices.sort_unstable();
                        claimed_page_indices.dedup();
                        (table, claimed_page_indices)
                    })
                    .collect()
            };

        // Run hint-scoped heuristics even when some rect tables were already
        // found on the page. Pages like boxed financial dashboards often have
        // one cluster that forms a valid rect-backed table and adjacent boxed
        // sections that only survive as hint regions.
        let hint_padding = 15.0;
        for hint in merge_hint_regions_into_bands(&hint_regions) {
            let inside_indices: Vec<usize> = page_items
                .iter()
                .enumerate()
                .filter(|(idx, item)| {
                    !rect_claimed.contains(idx) && item_in_hint_region(item, &hint, hint_padding)
                })
                .map(|(idx, _)| idx)
                .collect();
            let claimed_tables = run_heuristic(&inside_indices, 6);
            if claimed_tables.is_empty() {
                if let Some(table) = salvage_boxed_section_table(&page_items, &inside_indices) {
                    for &page_idx in &table.item_indices {
                        rect_claimed.insert(page_idx);
                        if let Some(&(global_idx, _)) = group.get(page_idx) {
                            table_items.insert(global_idx);
                        }
                    }
                    let table_y = table.rows.first().copied().unwrap_or(hint.y_top);
                    let table_md = table_to_markdown(&table);
                    if !table_md.trim().is_empty() {
                        page_ctx.tables.push((table_y, table_md));
                    }
                }
            } else {
                for (table, claimed) in claimed_tables {
                    for page_idx in claimed {
                        rect_claimed.insert(page_idx);
                        if let Some(&(global_idx, _)) = group.get(page_idx) {
                            table_items.insert(global_idx);
                        }
                    }
                    let table_y = table.rows.first().copied().unwrap_or(hint.y_top);
                    let table_md = table_to_markdown(&table);
                    if !table_md.trim().is_empty() {
                        page_ctx.tables.push((table_y, table_md));
                    }
                }
            }
        }

        if rect_tables.is_empty() {
            // Only run the broad page-wide fallback when nothing rect-backed was
            // recovered. Otherwise it tends to manufacture noisy tables from
            // leftover chart labels and headings.
            let unclaimed_indices: Vec<usize> = page_items
                .iter()
                .enumerate()
                .filter(|(idx, _)| !rect_claimed.contains(idx))
                .map(|(idx, _)| idx)
                .collect();
            for (table, claimed) in run_heuristic(&unclaimed_indices, 6) {
                for page_idx in claimed {
                    rect_claimed.insert(page_idx);
                    if let Some(&(global_idx, _)) = group.get(page_idx) {
                        table_items.insert(global_idx);
                    }
                }
                let table_y = table.rows.first().copied().unwrap_or(0.0);
                let table_md = table_to_markdown(&table);
                if !table_md.trim().is_empty() {
                    page_ctx.tables.push((table_y, table_md));
                }
            }
        }
    }

    let mut page_tables: HashMap<u32, Vec<(f32, String)>> = HashMap::new();
    for (&page, ctx) in &page_contexts {
        if !ctx.tables.is_empty() {
            page_tables.insert(page, ctx.tables.clone());
        }
    }

    // Filter out table items and process the rest
    let non_table_items: Vec<TextItem> = text_items
        .into_iter()
        .enumerate()
        .filter(|(idx, item)| {
            if table_items.contains(idx) {
                return false;
            }

            // Unconditionally filter watermark fragments (single uppercase chars
            // with impossibly narrow width) on all pages
            if is_probably_watermark_fragment(item) {
                return false;
            }

            let page_has_tabular_regions = page_contexts
                .get(&item.page)
                .is_some_and(|ctx| ctx.has_tabular_regions);
            !(page_has_tabular_regions
                && (is_probably_rotated_or_decorative(item)
                    || is_probably_table_margin_noise(item)))
        })
        .map(|(_, item)| item)
        .collect();

    // Find pages that are table-only (no remaining non-table text)
    let table_only_pages: HashSet<u32> = {
        let pages_with_text: HashSet<u32> = non_table_items.iter().map(|i| i.page).collect();
        page_tables
            .keys()
            .filter(|p| !pages_with_text.contains(p))
            .copied()
            .collect()
    };

    // Merge continuation tables across page breaks, but only for table-only pages
    merge_continuation_tables(&mut page_tables, &table_only_pages);

    let lines: Vec<TextLine> = group_into_lines(non_table_items)
        .into_iter()
        .filter(|line| !is_probably_chart_noise_line(line) && !is_probably_chart_legend_line(line))
        .collect();

    // Convert to markdown, inserting tables and images at appropriate positions
    to_markdown_from_lines_with_tables_and_images(lines, options, page_tables, page_images)
}

#[cfg(test)]
mod tests {
    use super::*;
    use analysis::detect_header_level;
    use classify::{is_code_like, is_list_item};

    #[test]
    fn test_is_list_item() {
        assert!(is_list_item("• Item one"));
        assert!(is_list_item("- Item two"));
        assert!(is_list_item("* Item three"));
        assert!(is_list_item("1. First"));
        assert!(is_list_item("2) Second"));
        assert!(is_list_item("a. Letter item"));
        assert!(!is_list_item("Regular text"));
    }

    #[test]
    fn test_format_list_item() {
        assert_eq!(format_list_item("• Item"), "- Item");
        assert_eq!(format_list_item("- Item"), "- Item");
        assert_eq!(format_list_item("1. First"), "1. First");
    }

    #[test]
    fn test_is_code_like() {
        assert!(is_code_like("const x = 5;"));
        assert!(is_code_like("function foo() {"));
        assert!(is_code_like("import React from 'react'"));
        assert!(!is_code_like("This is regular text."));
    }

    #[test]
    fn test_detect_header_level() {
        // With three tiers: 24→H1, 18→H2, 15→H3, 12→None
        let tiers = vec![24.0, 18.0, 15.0];
        assert_eq!(detect_header_level(24.0, 12.0, &tiers), Some(1));
        assert_eq!(detect_header_level(18.0, 12.0, &tiers), Some(2));
        assert_eq!(detect_header_level(15.0, 12.0, &tiers), Some(3));
        assert_eq!(detect_header_level(12.0, 12.0, &tiers), None);

        // Single tier: 15→H1 (ratio 1.25 ≥ 1.2), 14→None (ratio 1.17 < 1.2)
        let tiers = vec![15.0];
        assert_eq!(detect_header_level(15.0, 12.0, &tiers), Some(1));
        assert_eq!(detect_header_level(14.0, 12.0, &tiers), None);
        assert_eq!(detect_header_level(12.0, 12.0, &tiers), None);

        // No tiers (empty): falls back to ratio thresholds
        let tiers: Vec<f32> = vec![];
        assert_eq!(detect_header_level(24.0, 12.0, &tiers), Some(1));
        assert_eq!(detect_header_level(18.0, 12.0, &tiers), Some(2));
        assert_eq!(detect_header_level(15.0, 12.0, &tiers), Some(3));
        assert_eq!(detect_header_level(14.5, 12.0, &tiers), Some(4));
        assert_eq!(detect_header_level(14.0, 12.0, &tiers), None);
        assert_eq!(detect_header_level(12.0, 12.0, &tiers), None);

        // Body text excluded when tiers exist: 13pt (ratio 1.08) → None
        let tiers = vec![20.0];
        assert_eq!(detect_header_level(13.0, 12.0, &tiers), None);
    }

    #[test]
    fn test_to_markdown() {
        let text = "• First item\n• Second item\n\nRegular paragraph.";
        let md = to_markdown(text, MarkdownOptions::default());
        assert!(md.contains("- First item"));
        assert!(md.contains("- Second item"));
    }

    #[test]
    fn test_chart_noise_line_is_detected() {
        let line = TextLine {
            items: vec![TextItem {
                text: "Exercice N-2 34 000 N-1 32 000 Exercice N 30 000 28 000 26 000 24 000 22 000 20 000".into(),
                x: 0.0,
                y: 100.0,
                width: 400.0,
                height: 12.0,
                font: "F1".into(),
                font_size: 12.0,
                page: 1,
                is_rotated: false,
                is_bold: false,
                is_italic: false,
                item_type: crate::pdf_inspector::types::ItemType::Text,
            }],
            y: 100.0,
            page: 1,
        };

        assert!(is_probably_chart_noise_line(&line));
    }
}
