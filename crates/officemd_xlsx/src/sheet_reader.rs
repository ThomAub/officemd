use std::collections::HashMap;

use officemd_core::ir::FormulaNote;
use officemd_core::opc::OpcPackage;
use quick_xml::Reader as XmlReader;
use quick_xml::events::{BytesStart, BytesText, Event};

use crate::error::XlsxError;
use crate::style_format::{StyleContext, ValueRenderMode, parse_cell_ref};

#[derive(Debug, Default)]
pub(crate) struct SheetTextGrid {
    pub(crate) cols: usize,
    pub(crate) rows: Vec<Vec<String>>,
    pub(crate) formulas: Vec<FormulaNote>,
}

#[derive(Debug, Default)]
struct StreamingCell {
    row: usize,
    col: usize,
    style_index: Option<usize>,
    cell_type: Option<String>,
    raw_value: String,
    inline_text: String,
    formula: String,
}

#[derive(Debug, Default)]
struct RowBuf {
    cells: Vec<(usize, String)>,
}

const MAX_DENSE_GRID_CELLS: usize = 2_000_000;

#[allow(clippy::too_many_lines)]
pub(crate) fn collect_sheet_text_grid(
    package: &mut OpcPackage<'_>,
    sheet_path: &str,
    style_context: &StyleContext,
    mode: ValueRenderMode,
) -> Result<SheetTextGrid, XlsxError> {
    let Some(sheet_xml_bytes) = package
        .read_part_bytes(sheet_path)
        .map_err(XlsxError::from)?
    else {
        return Ok(SheetTextGrid::default());
    };

    let mut reader = XmlReader::from_reader(std::io::Cursor::new(sheet_xml_bytes));
    reader.config_mut().trim_text(false);

    let mut rows_by_index: HashMap<usize, RowBuf> = HashMap::new();
    let mut formulas = Vec::new();

    let mut current_row = 0usize;
    let mut next_row = 0usize;
    let mut next_col = 0usize;

    let mut in_value = false;
    let mut in_formula = false;
    let mut in_inline_text = false;
    let mut current_cell: Option<StreamingCell> = None;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match local_name(e.name().as_ref()) {
                b"row" => {
                    let row_idx = attr_usize(e, b"r").map_or(next_row, |r| r.saturating_sub(1));
                    current_row = row_idx;
                    next_row = row_idx + 1;
                    next_col = 0;
                }
                b"c" => {
                    current_cell = Some(start_streaming_cell(e, current_row, next_col));
                    if let Some(cell) = &current_cell {
                        next_col = cell.col + 1;
                    }
                }
                b"v" => in_value = true,
                b"f" => in_formula = true,
                b"t" => {
                    if current_cell.as_ref().and_then(|c| c.cell_type.as_deref())
                        == Some("inlineStr")
                    {
                        in_inline_text = true;
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match local_name(e.name().as_ref()) {
                b"row" => {
                    if let Some(row_idx) = attr_usize(e, b"r").map(|r| r.saturating_sub(1)) {
                        next_row = row_idx + 1;
                    } else {
                        next_row += 1;
                    }
                }
                b"c" => {
                    let cell = start_streaming_cell(e, current_row, next_col);
                    next_col = cell.col + 1;
                    insert_streaming_cell(
                        &mut rows_by_index,
                        cell,
                        &mut formulas,
                        style_context,
                        mode,
                    );
                }
                _ => {}
            },
            Ok(Event::Text(t)) => {
                let text = unescape_text(&t)?;
                if in_value {
                    if let Some(cell) = &mut current_cell {
                        cell.raw_value.push_str(&text);
                    }
                } else if in_formula {
                    if let Some(cell) = &mut current_cell {
                        cell.formula.push_str(&text);
                    }
                } else if in_inline_text && let Some(cell) = &mut current_cell {
                    cell.inline_text.push_str(&text);
                }
            }
            Ok(Event::CData(t)) => {
                if let Ok(text) = std::str::from_utf8(t.as_ref()) {
                    if in_value {
                        if let Some(cell) = &mut current_cell {
                            cell.raw_value.push_str(text);
                        }
                    } else if in_formula {
                        if let Some(cell) = &mut current_cell {
                            cell.formula.push_str(text);
                        }
                    } else if in_inline_text && let Some(cell) = &mut current_cell {
                        cell.inline_text.push_str(text);
                    }
                }
            }
            Ok(Event::End(ref e)) => match local_name(e.name().as_ref()) {
                b"v" => in_value = false,
                b"f" => in_formula = false,
                b"t" => in_inline_text = false,
                b"c" => {
                    if let Some(cell) = current_cell.take() {
                        insert_streaming_cell(
                            &mut rows_by_index,
                            cell,
                            &mut formulas,
                            style_context,
                            mode,
                        );
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(XlsxError::Xml(e.to_string())),
        }
        buf.clear();
    }

    if rows_by_index.is_empty() {
        return Ok(SheetTextGrid::default());
    }

    let mut row_indices = rows_by_index.keys().copied().collect::<Vec<_>>();
    row_indices.sort_unstable();

    let mut unique_cols = rows_by_index
        .values()
        .flat_map(|row| row.cells.iter().map(|(col, _)| *col))
        .collect::<Vec<_>>();
    unique_cols.sort_unstable();
    unique_cols.dedup();

    if unique_cols.is_empty() {
        return Ok(SheetTextGrid::default());
    }

    let min_row = row_indices[0];
    let max_row = *row_indices.last().unwrap_or(&min_row);
    let min_col = unique_cols[0];
    let max_col = *unique_cols.last().unwrap_or(&min_col);

    let dense_rows = max_row.saturating_sub(min_row) + 1;
    let dense_cols = max_col.saturating_sub(min_col) + 1;
    let dense_cells = dense_rows.saturating_mul(dense_cols);

    let (cols, rows) = if dense_cells <= MAX_DENSE_GRID_CELLS {
        let cols = dense_cols;
        let mut rows = Vec::with_capacity(dense_rows);
        for row_idx in min_row..=max_row {
            let mut row = vec![String::new(); cols];
            if let Some(values) = rows_by_index.get(&row_idx) {
                for (absolute_col, text) in &values.cells {
                    if *absolute_col >= min_col && *absolute_col <= max_col {
                        row[*absolute_col - min_col].clone_from(text);
                    }
                }
            }
            rows.push(row);
        }
        (cols, rows)
    } else {
        let cols = unique_cols.len();
        let mut col_index = HashMap::with_capacity(cols);
        for (dense_idx, col) in unique_cols.iter().copied().enumerate() {
            col_index.insert(col, dense_idx);
        }

        let mut rows = Vec::with_capacity(row_indices.len());
        for row_idx in row_indices {
            let mut row = vec![String::new(); cols];
            if let Some(values) = rows_by_index.get(&row_idx) {
                for (absolute_col, text) in &values.cells {
                    if let Some(dense_col) = col_index.get(absolute_col) {
                        row[*dense_col].clone_from(text);
                    }
                }
            }
            rows.push(row);
        }
        (cols, rows)
    };

    Ok(SheetTextGrid {
        cols,
        rows,
        formulas,
    })
}

fn insert_streaming_cell(
    rows_by_index: &mut HashMap<usize, RowBuf>,
    mut cell: StreamingCell,
    formulas: &mut Vec<FormulaNote>,
    style_context: &StyleContext,
    mode: ValueRenderMode,
) {
    let formula_raw = std::mem::take(&mut cell.formula);
    let formula = formula_raw.trim();
    if !formula.is_empty() {
        formulas.push(FormulaNote {
            cell_ref: format!("{}{}", col_to_name(cell.col + 1), cell.row + 1),
            formula: formula.strip_prefix('=').unwrap_or(formula).to_string(),
        });
    }

    let text = style_context.render_cell_text(
        cell.cell_type.as_deref(),
        &cell.raw_value,
        &cell.inline_text,
        cell.style_index,
        mode,
    );

    let row = rows_by_index.entry(cell.row).or_default();
    row.cells.push((cell.col, text));
}

fn start_streaming_cell(e: &BytesStart<'_>, current_row: usize, next_col: usize) -> StreamingCell {
    let mut cell = StreamingCell {
        row: current_row,
        col: next_col,
        ..Default::default()
    };

    if let Some(cell_ref) = attr_string(e, b"r")
        && let Some((row, col)) = parse_cell_ref(&cell_ref)
    {
        cell.row = row;
        cell.col = col;
    }

    cell.style_index = attr_usize(e, b"s");
    cell.cell_type = attr_string(e, b"t");
    cell
}

fn local_name(name: &[u8]) -> &[u8] {
    if let Some(idx) = name.iter().rposition(|b| *b == b':') {
        &name[idx + 1..]
    } else if let Some(idx) = name.iter().rposition(|b| *b == b'}') {
        &name[idx + 1..]
    } else {
        name
    }
}

fn attr_string(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let attr_key = local_name(attr.key.as_ref());
        if attr_key == key {
            if let Ok(value) = attr.unescape_value() {
                return Some(value.into_owned());
            }
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn attr_usize(e: &BytesStart<'_>, key: &[u8]) -> Option<usize> {
    attr_string(e, key)?.parse::<usize>().ok()
}

fn col_to_name(mut n: usize) -> String {
    let mut reversed = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        #[allow(clippy::cast_possible_truncation)]
        reversed.push((b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    reversed.into_iter().rev().collect()
}

fn unescape_text(t: &BytesText<'_>) -> Result<String, XlsxError> {
    t.unescape()
        .map(std::borrow::Cow::into_owned)
        .map_err(|e| XlsxError::Xml(e.to_string()))
}
