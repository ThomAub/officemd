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
use crate::pdf_inspector::types::TextItem;

use analysis::calculate_font_stats_from_items;
use classify::{format_list_item, is_code_like, is_list_item};
use convert::{merge_continuation_tables, to_markdown_from_lines_with_tables_and_images};

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
    let mut page_tables: HashMap<u32, Vec<(f32, String)>> = HashMap::new();

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
        let page_items: Vec<TextItem> = group.iter().map(|(_, item)| (*item).clone()).collect();

        // Track which local indices are claimed by rect-based tables
        let mut rect_claimed: HashSet<usize> = HashSet::new();

        // Try rectangle-based table detection first
        let (rect_tables, hint_regions) = detect_tables_from_rects(&page_items, rects, page);
        for table in &rect_tables {
            for &idx in &table.item_indices {
                rect_claimed.insert(idx);
                if let Some(&(global_idx, _)) = group.get(idx) {
                    table_items.insert(global_idx);
                }
            }
            let table_y = table.rows.first().copied().unwrap_or(0.0);
            let table_md = table_to_markdown(table);
            page_tables
                .entry(page)
                .or_default()
                .push((table_y, table_md));
        }

        // Helper: run heuristic on a subset of items, remapping indices back to page-space
        let mut run_heuristic =
            |subset_items: &[TextItem], index_map: &[usize], min_items: usize| {
                if subset_items.len() < min_items {
                    return;
                }
                let tables = detect_tables(subset_items, base_size, false);
                for table in tables {
                    for &idx in &table.item_indices {
                        if let Some(&page_idx) = index_map.get(idx) {
                            if let Some(&(global_idx, _)) = group.get(page_idx) {
                                table_items.insert(global_idx);
                            }
                        }
                    }
                    let table_y = table.rows.first().copied().unwrap_or(0.0);
                    let table_md = table_to_markdown(&table);
                    page_tables
                        .entry(page)
                        .or_default()
                        .push((table_y, table_md));
                }
            };

        // Run heuristic detection on unclaimed items
        if rect_claimed.is_empty() && hint_regions.is_empty() {
            // No rect tables or hints — run heuristic on all items
            let identity_map: Vec<usize> = (0..page_items.len()).collect();
            run_heuristic(&page_items, &identity_map, 6);
        } else if rect_claimed.is_empty() && !hint_regions.is_empty() {
            // No rect tables but hint regions exist — run heuristic separately
            // on items inside each hint region and on items outside all hints.
            // This prevents graph labels from being merged into nearby tables.
            let padding = 15.0; // include header lines slightly above rects
            for hint in &hint_regions {
                let (inside_items, inside_map): (Vec<TextItem>, Vec<usize>) = page_items
                    .iter()
                    .enumerate()
                    .filter(|(_, item)| {
                        item.y >= hint.y_bottom - padding && item.y <= hint.y_top + padding
                    })
                    .map(|(idx, item)| (item.clone(), idx))
                    .unzip();
                run_heuristic(&inside_items, &inside_map, 6);
                // Mark hint-region items as claimed so they aren't re-processed
                for &page_idx in &inside_map {
                    rect_claimed.insert(page_idx);
                }
            }
            // Run heuristic on remaining items outside all hint regions
            let (outside_items, outside_map): (Vec<TextItem>, Vec<usize>) = page_items
                .iter()
                .enumerate()
                .filter(|(idx, _)| !rect_claimed.contains(idx))
                .map(|(idx, item)| (item.clone(), idx))
                .unzip();
            run_heuristic(&outside_items, &outside_map, 6);
        } else {
            // Rect tables found — run heuristic on unclaimed items
            let (unclaimed_items, unclaimed_map): (Vec<TextItem>, Vec<usize>) = page_items
                .iter()
                .enumerate()
                .filter(|(idx, _)| !rect_claimed.contains(idx))
                .map(|(idx, item)| (item.clone(), idx))
                .unzip();
            run_heuristic(&unclaimed_items, &unclaimed_map, 6);
        }
    }

    // Filter out table items and process the rest
    let non_table_items: Vec<TextItem> = text_items
        .into_iter()
        .enumerate()
        .filter(|(idx, _)| !table_items.contains(idx))
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

    let lines = group_into_lines(non_table_items);

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
}
