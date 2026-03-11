//! Table detection and formatting.
//!
//! Detects tabular data in PDF text items and converts to markdown tables.

mod detect_heuristic;
mod detect_rects;
mod financial;
mod format;
mod grid;

pub use detect_heuristic::detect_tables;
pub use detect_rects::{RectHintRegion, detect_tables_from_rects};
pub use format::table_to_markdown;

/// Detection mode controls thresholds for table validation.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum TableDetectionMode {
    /// Existing behavior: items with font size smaller than body text
    SmallFont,
    /// New: body-font items with stricter structural criteria
    BodyFont,
}

/// A detected table.
#[derive(Debug, Clone)]
pub struct Table {
    /// Column boundaries (x positions)
    pub columns: Vec<f32>,
    /// Row boundaries (y positions, descending order)
    pub rows: Vec<f32>,
    /// Cell contents indexed by (row, col)
    pub cells: Vec<Vec<String>>,
    /// Items that belong to this table
    pub item_indices: Vec<usize>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pdf_inspector::types::{ItemType, TextItem};

    fn make_item(text: &str, x: f32, y: f32, font_size: f32) -> TextItem {
        TextItem {
            text: text.into(),
            x,
            y,
            width: 10.0,
            height: font_size,
            font: "F1".into(),
            font_size,
            page: 1,
            is_rotated: false,
            is_bold: false,
            is_italic: false,
            item_type: ItemType::Text,
        }
    }

    fn make_char(text: &str, x: f32, y: f32, font_size: f32, width: f32) -> TextItem {
        TextItem {
            text: text.into(),
            x,
            y,
            width,
            height: font_size,
            font: "F1".into(),
            font_size,
            page: 1,
            is_rotated: false,
            is_bold: false,
            is_italic: false,
            item_type: ItemType::Text,
        }
    }

    #[test]
    fn test_table_detection() {
        let items = vec![
            // Header row
            make_item("Subject", 100.0, 500.0, 8.0),
            make_item("Q1", 200.0, 500.0, 8.0),
            make_item("Q2", 280.0, 500.0, 8.0),
            make_item("Q3", 360.0, 500.0, 8.0),
            // Data row 1
            make_item("Math", 100.0, 480.0, 8.0),
            make_item("9.0", 200.0, 480.0, 8.0),
            make_item("8.5", 280.0, 480.0, 8.0),
            make_item("9.5", 360.0, 480.0, 8.0),
            // Data row 2
            make_item("Science", 100.0, 460.0, 8.0),
            make_item("8.0", 200.0, 460.0, 8.0),
            make_item("9.0", 280.0, 460.0, 8.0),
            make_item("8.5", 360.0, 460.0, 8.0),
            // Data row 3
            make_item("English", 100.0, 440.0, 8.0),
            make_item("9.5", 200.0, 440.0, 8.0),
            make_item("9.0", 280.0, 440.0, 8.0),
            make_item("9.5", 360.0, 440.0, 8.0),
        ];

        let tables = detect_tables(&items, 10.0, false);
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].columns.len(), 4);
        assert_eq!(tables[0].rows.len(), 4);
    }

    #[test]
    fn test_table_to_markdown() {
        let table = Table {
            columns: vec![100.0, 200.0],
            rows: vec![500.0, 480.0],
            cells: vec![
                vec!["Header 1".into(), "Header 2".into()],
                vec!["Cell 1".into(), "Cell 2".into()],
            ],
            item_indices: vec![],
        };

        let md = table_to_markdown(&table);
        assert!(md.contains("| Header 1"));
        assert!(md.contains("| ---"));
        assert!(md.contains("| Cell 1"));
    }

    #[test]
    fn test_body_font_table_detected() {
        let items = vec![
            // Header row
            make_item("Name", 100.0, 500.0, 10.0),
            make_item("Price", 200.0, 500.0, 10.0),
            make_item("Qty", 300.0, 500.0, 10.0),
            make_item("Total", 400.0, 500.0, 10.0),
            // Data row 1
            make_item("Widget", 100.0, 480.0, 10.0),
            make_item("5.00", 200.0, 480.0, 10.0),
            make_item("10", 300.0, 480.0, 10.0),
            make_item("50.00", 400.0, 480.0, 10.0),
            // Data row 2
            make_item("Gadget", 100.0, 460.0, 10.0),
            make_item("12.50", 200.0, 460.0, 10.0),
            make_item("4", 300.0, 460.0, 10.0),
            make_item("50.00", 400.0, 460.0, 10.0),
            // Data row 3
            make_item("Gizmo", 100.0, 440.0, 10.0),
            make_item("3.25", 200.0, 440.0, 10.0),
            make_item("20", 300.0, 440.0, 10.0),
            make_item("65.00", 400.0, 440.0, 10.0),
        ];

        let tables = detect_tables(&items, 10.0, false);
        assert_eq!(
            tables.len(),
            1,
            "Body-font table should be detected by Pass 2"
        );
        assert_eq!(tables[0].columns.len(), 4);
        assert!(tables[0].rows.len() >= 3);
    }

    #[test]
    fn test_paragraph_not_falsely_detected() {
        let items = vec![
            make_item(
                "This is a paragraph of text that spans the full width",
                72.0,
                500.0,
                10.0,
            ),
            make_item(
                "of the page and should not be detected as a table.",
                72.0,
                485.0,
                10.0,
            ),
            make_item(
                "It continues for several lines with normal body text",
                72.0,
                470.0,
                10.0,
            ),
            make_item(
                "that is left-aligned and has no columnar structure.",
                72.0,
                455.0,
                10.0,
            ),
            make_item(
                "The paragraph keeps going with more content here.",
                72.0,
                440.0,
                10.0,
            ),
            make_item(
                "And it has even more text on this line as well.",
                72.0,
                425.0,
                10.0,
            ),
            make_item(
                "Finally the paragraph concludes with this last line.",
                72.0,
                410.0,
                10.0,
            ),
            make_item(
                "One more line to have enough items for detection.",
                72.0,
                395.0,
                10.0,
            ),
            make_item(
                "And another line of plain paragraph text content.",
                72.0,
                380.0,
                10.0,
            ),
            make_item(
                "Last line of the paragraph ends here for the test.",
                72.0,
                365.0,
                10.0,
            ),
        ];

        let tables = detect_tables(&items, 10.0, false);
        assert_eq!(
            tables.len(),
            0,
            "Single-column paragraph must not be detected as table"
        );
    }

    #[test]
    fn test_word_level_paragraph_not_detected_as_table() {
        let items = vec![
            // Line 1
            make_item("We", 72.0, 500.0, 10.0),
            make_item("would", 95.0, 500.0, 10.0),
            make_item("like", 145.0, 500.0, 10.0),
            make_item("to", 180.0, 500.0, 10.0),
            make_item("thank", 200.0, 500.0, 10.0),
            make_item("all", 250.0, 500.0, 10.0),
            make_item("the", 278.0, 500.0, 10.0),
            make_item("practitioners", 305.0, 500.0, 10.0),
            // Line 2
            make_item("and", 72.0, 485.0, 10.0),
            make_item("researchers", 105.0, 485.0, 10.0),
            make_item("across", 185.0, 485.0, 10.0),
            make_item("the", 232.0, 485.0, 10.0),
            make_item("University", 260.0, 485.0, 10.0),
            make_item("of", 335.0, 485.0, 10.0),
            make_item("Leeds", 355.0, 485.0, 10.0),
            // Line 3
            make_item("Libraries", 72.0, 470.0, 10.0),
            make_item("whose", 142.0, 470.0, 10.0),
            make_item("contributions", 190.0, 470.0, 10.0),
            make_item("made", 290.0, 470.0, 10.0),
            make_item("this", 328.0, 470.0, 10.0),
            make_item("report", 360.0, 470.0, 10.0),
            // Line 4
            make_item("possible", 72.0, 455.0, 10.0),
            make_item("Both", 140.0, 455.0, 10.0),
            make_item("constituent", 178.0, 455.0, 10.0),
            make_item("studies", 262.0, 455.0, 10.0),
            make_item("were", 315.0, 455.0, 10.0),
            make_item("approved", 350.0, 455.0, 10.0),
        ];

        let tables = detect_tables(&items, 10.0, false);
        assert_eq!(
            tables.len(),
            0,
            "Word-level paragraph text must not be detected as table"
        );
    }

    #[test]
    fn test_large_data_table_not_rejected() {
        let mut items = Vec::new();
        // Header row
        items.push(make_item("Temp", 100.0, 800.0, 8.0));
        items.push(make_item("Pressure", 200.0, 800.0, 8.0));
        items.push(make_item("Volume", 300.0, 800.0, 8.0));
        items.push(make_item("Enthalpy", 400.0, 800.0, 8.0));

        // 49 data rows
        for i in 1..50 {
            let y = 800.0 - (i as f32 * 12.0);
            items.push(make_item(&format!("{}", -40 + i * 2), 100.0, y, 8.0));
            items.push(make_item(
                &format!("{:.1}", 100.0 + i as f32 * 5.0),
                200.0,
                y,
                8.0,
            ));
            items.push(make_item(
                &format!("{:.3}", 0.05 + i as f32 * 0.01),
                300.0,
                y,
                8.0,
            ));
            items.push(make_item(
                &format!("{:.1}", 150.0 + i as f32 * 2.5),
                400.0,
                y,
                8.0,
            ));
        }

        let tables = detect_tables(&items, 10.0, false);
        assert_eq!(tables.len(), 1, "Large data table should not be rejected");
        assert!(
            tables[0].rows.len() >= 40,
            "Large table should preserve most rows, got {}",
            tables[0].rows.len()
        );
    }

    #[test]
    fn test_uniform_spacing_rows_not_merged() {
        let companies = [
            "SC Priority LLC",
            "Craft Roofing Co",
            "Alpha Roofing Inc",
            "Beta Construction",
            "Gamma Builders",
            "Delta Roofing",
            "Epsilon Contractors",
        ];

        let mut items = Vec::new();

        // Header row at y=800
        items.push(make_item("No.", 50.0, 800.0, 8.0));
        items.push(make_item("Company", 120.0, 800.0, 8.0));
        items.push(make_item("Bid Amount", 350.0, 800.0, 8.0));

        // 7 data rows, each 10pt apart (exactly the old threshold)
        for (i, company) in companies.iter().enumerate() {
            let y = 790.0 - (i as f32 * 10.0);
            items.push(make_item(&format!("{}", i + 1), 50.0, y, 8.0));
            items.push(make_item(company, 120.0, y, 8.0));
            items.push(make_item(&format!("${},000", 100 + i * 10), 350.0, y, 8.0));
        }

        let tables = detect_tables(&items, 12.0, false);
        assert_eq!(tables.len(), 1, "Should detect one table");
        assert_eq!(
            tables[0].rows.len(),
            8,
            "Each company must be on its own row, got {} rows instead of 8",
            tables[0].rows.len()
        );
    }

    #[test]
    fn test_merge_adjacent_items() {
        let items = vec![
            make_char("J", 310.0, 532.0, 13.3, 4.0),
            make_char("u", 314.0, 532.0, 13.3, 4.4),
            make_char("n", 318.4, 532.0, 13.3, 4.4),
            make_char("e", 322.8, 532.0, 13.3, 3.5),
            // word gap (2pt)
            make_char("3", 328.3, 532.0, 13.3, 4.0),
            make_char("0", 332.3, 532.0, 13.3, 4.0),
            make_char(",", 336.3, 532.0, 13.3, 2.0),
            // large column gap (40pt)
            make_char("M", 378.3, 532.0, 13.3, 7.5),
            make_char("a", 385.8, 532.0, 13.3, 4.0),
            make_char("r", 389.8, 532.0, 13.3, 3.5),
        ];

        let (merged, map) = detect_heuristic::merge_adjacent_items(&items);

        assert_eq!(
            merged.len(),
            2,
            "Should produce 2 merged items, got {}",
            merged.len()
        );
        assert!(
            merged[0].text.contains("June") && merged[0].text.contains("30"),
            "First merged item should be 'June 30,' but got {:?}",
            merged[0].text
        );
        assert_eq!(merged[1].text, "Mar");

        assert_eq!(
            map[0].len(),
            7,
            "First merged item should map to 7 original chars"
        );
        assert_eq!(
            map[1].len(),
            3,
            "Second merged item should map to 3 original chars"
        );
    }

    #[test]
    fn test_per_char_financial_table_detected() {
        let mut items = Vec::new();

        // Per-character header row
        for (i, c) in "Col1".chars().enumerate() {
            items.push(make_char(
                &c.to_string(),
                300.0 + i as f32 * 5.0,
                540.0,
                13.0,
                5.0,
            ));
        }
        for (i, c) in "Col2".chars().enumerate() {
            items.push(make_char(
                &c.to_string(),
                400.0 + i as f32 * 5.0,
                540.0,
                13.0,
                5.0,
            ));
        }
        for (i, c) in "Col3".chars().enumerate() {
            items.push(make_char(
                &c.to_string(),
                500.0 + i as f32 * 5.0,
                540.0,
                13.0,
                5.0,
            ));
        }

        // Data rows with multi-word items
        let data = [
            ("Revenue", 520.0, "1,000", "2,000", "3,000"),
            ("Expenses", 505.0, "500", "800", "1,200"),
            ("Net Income", 490.0, "500", "1,200", "1,800"),
            ("Taxes", 475.0, "100", "200", "300"),
        ];

        for (label, y, v1, v2, v3) in &data {
            items.push(make_item(label, 50.0, *y, 12.0));
            items.push(make_item(v1, 310.0, *y, 12.0));
            items.push(make_item(v2, 410.0, *y, 12.0));
            items.push(make_item(v3, 510.0, *y, 12.0));
        }

        let tables = detect_tables(&items, 13.0, false);
        assert!(
            !tables.is_empty(),
            "Per-character financial table should be detected"
        );
    }

    #[test]
    fn test_short_subheader_not_merged_as_continuation() {
        // Simulate a table with section sub-headers (like month names) that have
        // an empty first column and short text in a single other column.
        // These should NOT be merged into the previous row as continuation text.
        let table = Table {
            columns: vec![50.0, 150.0, 300.0, 450.0],
            rows: vec![500.0, 480.0, 460.0, 440.0, 420.0, 400.0],
            cells: vec![
                // Header row
                vec!["No.".into(), "Date".into(), "Title".into(), "Amount".into()],
                // Sub-header: month name in 1 column, rest empty
                vec!["".into(), "JAN".into(), "".into(), "".into()],
                // Data row
                vec!["1".into(), "8/1".into(), "Item A".into(), "100".into()],
                vec!["2".into(), "15/1".into(), "Item B".into(), "200".into()],
                // Another sub-header
                vec!["".into(), "FEB".into(), "".into(), "".into()],
                // Data row
                vec!["3".into(), "5/2".into(), "Item C".into(), "300".into()],
            ],
            item_indices: vec![],
        };

        let md = table_to_markdown(&table);
        // JAN and FEB should be on their own rows, not merged into adjacent rows
        assert!(
            md.contains("| JAN"),
            "JAN should be on its own row, got:\n{}",
            md
        );
        assert!(
            md.contains("| FEB"),
            "FEB should be on its own row, got:\n{}",
            md
        );
        // Verify they're NOT merged into data rows
        assert!(
            !md.contains("15/1 FEB"),
            "FEB should not be merged into data row, got:\n{}",
            md
        );
        assert!(
            !md.contains("8/1 JAN"),
            "JAN should not be merged into data row, got:\n{}",
            md
        );
    }
}
