//! Table-to-markdown formatting and cell cleanup.

use super::Table;

/// Render a detected PDF table as GitHub-flavored Markdown.
///
/// Returns an empty string when the table has no usable cell content.
pub fn table_to_markdown(table: &Table) -> String {
    if table.cells.is_empty() || table.cells[0].is_empty() {
        return String::new();
    }

    // Clean up the table: merge continuation rows, extract footnotes, remove empty rows
    let (cleaned_cells, footnotes) = clean_table_cells(&table.cells);

    if cleaned_cells.is_empty() {
        return String::new();
    }

    let num_cols = cleaned_cells[0].len();
    let mut output = String::new();

    // Calculate column widths for alignment
    let col_widths: Vec<usize> = (0..num_cols)
        .map(|col| {
            cleaned_cells
                .iter()
                .map(|row| row.get(col).map(|c| c.len()).unwrap_or(0))
                .max()
                .unwrap_or(3)
                .max(3)
        })
        .collect();

    // Output each row
    for (row_idx, row) in cleaned_cells.iter().enumerate() {
        output.push('|');
        for (col_idx, cell) in row.iter().enumerate() {
            let width = col_widths[col_idx];
            output.push_str(&format!(" {:width$} |", cell, width = width));
        }
        output.push('\n');

        // Add separator after header row
        if row_idx == 0 {
            output.push('|');
            for width in &col_widths {
                output.push_str(&format!(" {} |", "-".repeat(*width)));
            }
            output.push('\n');
        }
    }

    // Add footnotes below the table
    if !footnotes.is_empty() {
        output.push('\n');
        for footnote in footnotes {
            output.push_str(&footnote);
            output.push('\n');
        }
    }

    output
}

/// Clean up table cells: merge continuation rows, extract footnotes, remove empty rows
fn clean_table_cells(cells: &[Vec<String>]) -> (Vec<Vec<String>>, Vec<String>) {
    let mut cleaned: Vec<Vec<String>> = Vec::new();
    let mut footnotes: Vec<String> = Vec::new();

    for row in cells {
        // Check if this row is empty
        if row.iter().all(|c| c.trim().is_empty()) {
            continue;
        }

        // Check if this row is a footnote (starts with (1), (2), etc. or just a number reference)
        let first_cell = row.first().map(|s| s.trim()).unwrap_or("");
        if is_footnote_row(first_cell) {
            // Combine all cells into a single footnote line
            let footnote_text: String = row
                .iter()
                .map(|c| c.trim())
                .filter(|c| !c.is_empty())
                .collect::<Vec<_>>()
                .join(" ");
            footnotes.push(footnote_text);
            continue;
        }

        // Check if this is a continuation row (first column is empty but others have content).
        // A row with only 1 short non-empty cell (besides the first) is more likely a
        // section sub-header (e.g. "JAN", "FEB") than overflow text — don't merge it.
        let non_first_cells: Vec<&str> = row
            .iter()
            .skip(1)
            .map(|c| c.trim())
            .filter(|c| !c.is_empty())
            .collect();
        let is_short_subheader = non_first_cells.len() == 1 && non_first_cells[0].len() <= 5;
        let is_continuation = first_cell.is_empty()
            && !non_first_cells.is_empty()
            && !is_short_subheader
            && cleaned.len() > 1; // Don't merge into the first row (header)

        if is_continuation {
            // Merge with previous row
            if let Some(prev_row) = cleaned.last_mut() {
                for (col_idx, cell) in row.iter().enumerate() {
                    let cell_text = cell.trim();
                    if !cell_text.is_empty() && col_idx < prev_row.len() {
                        if !prev_row[col_idx].is_empty() {
                            prev_row[col_idx].push(' ');
                        }
                        prev_row[col_idx].push_str(cell_text);
                    }
                }
            }
        } else {
            // Regular row - add as new row
            cleaned.push(row.iter().map(|c| c.trim().to_string()).collect());
        }
    }

    (cleaned, footnotes)
}

/// Check if a cell value indicates a footnote row
fn is_footnote_row(text: &str) -> bool {
    let trimmed = text.trim();

    // Check for common footnote patterns
    // (1), (2), etc.
    if trimmed.starts_with('(') && trimmed.len() >= 2 {
        let inside = &trimmed[1..];
        if let Some(close_idx) = inside.find(')') {
            let num_part = &inside[..close_idx];
            if num_part.chars().all(|c| c.is_ascii_digit()) {
                return true;
            }
        }
    }

    // 1), 2), etc.
    if trimmed.len() >= 2 {
        if let Some(paren_idx) = trimmed.find(')') {
            let num_part = &trimmed[..paren_idx];
            if !num_part.is_empty() && num_part.chars().all(|c| c.is_ascii_digit()) {
                return true;
            }
        }
    }

    // Check for "Note:" or "Notes:" at the start
    let lower = trimmed.to_lowercase();
    if lower.starts_with("note:") || lower.starts_with("notes:") {
        return true;
    }

    false
}
