//! Build a table-centric IR for XLSX using an in-house XML reader.
//!
//! - Always returns at least one table per sheet (synthetic headers).
//! - Collects formulas for footnotes.
//! - Uses shared OPC plumbing from `officemd_core`.

use std::collections::{HashMap, HashSet};

use crate::error::XlsxError;
use crate::sheet_reader::{SheetTextGrid, collect_sheet_text_grid};
use crate::style_format::{StyleContext, ValueRenderMode};
use officemd_core::ir::{
    DocumentKind, DocumentProperties, OoxmlDocument, Paragraph, Sheet, Table, TableCell,
};
use officemd_core::opc::{OpcPackage, load_relationships_for_part, relationship_target_map};
use quick_xml::Reader as XmlReader;
use quick_xml::events::Event;

/// XLSX extraction options. Defaults keep current behavior.
#[derive(Debug, Clone, Default)]
pub struct XlsxExtractOptions {
    pub text: XlsxTextOptions,
    pub sheet_filter: Option<SheetFilter>,
    pub include: XlsxIncludeOptions,
    pub trim: XlsxTrimOptions,
}

/// Text extraction behavior for XLSX content.
#[derive(Debug, Clone, Default)]
pub struct XlsxTextOptions {
    pub style_aware_values: bool,
    /// Kept for backward compatibility; extraction uses the in-house XML reader in all modes.
    pub streaming_rows: bool,
}

/// Optional XLSX content to include in the extracted IR.
#[derive(Debug, Clone, Default)]
pub struct XlsxIncludeOptions {
    pub document_properties: bool,
}

/// XLSX grid trimming behavior.
#[derive(Debug, Clone, Default)]
pub struct XlsxTrimOptions {
    /// Strip trailing all-empty rows (from the bottom) and trailing all-empty
    /// columns (from the right) of each sheet grid before building the IR.
    /// Useful for LLM-compact output to avoid wasting tokens on empty cells.
    pub empty_edges: bool,
}

/// Optional sheet selection applied during extraction.
#[derive(Debug, Clone, Default)]
pub struct SheetFilter {
    pub names: HashSet<String>,
    pub indices_1_based: HashSet<usize>,
}

impl SheetFilter {
    /// Return whether a 1-based sheet index and sheet name are selected by this filter.
    ///
    /// An empty filter matches every sheet.
    #[must_use]
    pub fn matches(&self, idx1: usize, name: &str) -> bool {
        if self.names.is_empty() && self.indices_1_based.is_empty() {
            return true;
        }
        self.indices_1_based.contains(&idx1) || self.names.contains(name)
    }

    /// Return the 0-based indices of sheets selected by this filter.
    ///
    /// The input contains `(sheet_name, relationship_id)` pairs in workbook order.
    /// An empty filter returns every sheet index.
    #[must_use]
    pub fn selected_indices(&self, sheets: &[(String, String)]) -> Vec<usize> {
        if self.names.is_empty() && self.indices_1_based.is_empty() {
            return (0..sheets.len()).collect();
        }

        sheets
            .iter()
            .enumerate()
            .filter_map(|(idx, (name, _))| self.matches(idx + 1, name).then_some(idx))
            .collect()
    }
}

use officemd_core::ir::synthetic_col_headers;

/// Extract a simple IR with one table per sheet (no chunking yet).
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn extract_tables_ir(content: &[u8]) -> Result<OoxmlDocument, XlsxError> {
    extract_tables_ir_with_options(content, &XlsxExtractOptions::default())
}

/// Extract table-centric IR with explicit options.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn extract_tables_ir_with_options(
    content: &[u8],
    options: &XlsxExtractOptions,
) -> Result<OoxmlDocument, XlsxError> {
    let mut package = OpcPackage::from_bytes(content).map_err(XlsxError::from)?;
    // `streaming_rows` is retained for backward-compatible options surface;
    // both modes use the same in-house sheet reader.
    let style_context = StyleContext::load(&mut package)?;
    let render_mode = if options.text.style_aware_values {
        ValueRenderMode::StyleAware
    } else {
        ValueRenderMode::LegacyDefault
    };

    let sheet_targets = resolve_sheet_targets(&mut package)?;
    let mut sheets = Vec::with_capacity(sheet_targets.len());
    let selected_sheet_indices = options.sheet_filter.as_ref().map_or_else(
        || (0..sheet_targets.len()).collect(),
        |filter| filter.selected_indices(&sheet_targets),
    );

    for sheet_idx in selected_sheet_indices {
        let (name, path) = &sheet_targets[sheet_idx];
        let sheet_name = name.clone();
        let mut grid = collect_sheet_text_grid(&mut package, path, &style_context, render_mode)?;

        if options.trim.empty_edges {
            trim_grid(&mut grid);
        }

        let SheetTextGrid {
            cols: grid_cols,
            rows: grid_rows,
            formulas,
        } = grid;

        let cols = grid_cols.max(1);
        let headers = synthetic_col_headers(cols);
        let mut rows = build_ir_rows(grid_rows, cols);

        if rows.is_empty() {
            // Ensure at least 1x1 table per spec.
            rows.push(vec![TableCell {
                content: vec![Paragraph {
                    inlines: vec![officemd_core::ir::Inline::Text(String::new())],
                }],
            }]);
        }

        let caption = Some(format!(
            "Table {} (rows 1–{}, cols A–{})",
            sheet_idx + 1,
            rows.len().max(1),
            col_to_name(cols)
        ));

        let table = Table {
            caption,
            headers,
            rows,
            synthetic_headers: true,
        };

        sheets.push(Sheet {
            name: sheet_name,
            tables: vec![table],
            formulas,
            hyperlinks: Vec::new(),
        });
    }

    let properties = if options.include.document_properties {
        Some(DocumentProperties {
            core: extract_props_map(&mut package, "docProps/core.xml")?,
            app: extract_props_map(&mut package, "docProps/app.xml")?,
            custom: HashMap::new(),
        })
    } else {
        None
    };

    Ok(OoxmlDocument {
        kind: DocumentKind::Xlsx,
        sheets,
        properties,
        ..Default::default()
    })
}

/// Strip trailing all-empty rows from the bottom and trailing all-empty
/// columns from the right of the grid.
fn trim_grid(grid: &mut SheetTextGrid) {
    // 1. Remove trailing all-empty rows.
    while grid
        .rows
        .last()
        .is_some_and(|r| r.iter().all(String::is_empty))
    {
        grid.rows.pop();
    }

    // 2. Find the rightmost column index that contains any non-empty value.
    let max_col = grid
        .rows
        .iter()
        .flat_map(|r| {
            r.iter()
                .enumerate()
                .filter(|(_, c)| !c.is_empty())
                .map(|(i, _)| i)
        })
        .max()
        .map_or(0, |m| m + 1);

    // 3. Truncate columns if the data ends before grid.cols.
    if max_col < grid.cols {
        if max_col == 0 {
            // All remaining rows are entirely empty; reset to 0 cols so the
            // caller's `max(1)` fallback produces a minimal 1×1 table.
            grid.cols = 0;
            grid.rows.clear();
        } else {
            for row in &mut grid.rows {
                row.truncate(max_col);
            }
            grid.cols = max_col;
        }
    }
}

fn build_ir_rows(rows: Vec<Vec<String>>, cols: usize) -> Vec<Vec<TableCell>> {
    let mut out = Vec::with_capacity(rows.len());
    for mut row in rows {
        row.resize(cols, String::new());
        let mut cells = Vec::with_capacity(cols);
        for text in row.into_iter().take(cols) {
            let para = Paragraph {
                inlines: vec![officemd_core::ir::Inline::Text(text)],
            };
            cells.push(TableCell {
                content: vec![para],
            });
        }
        out.push(cells);
    }
    out
}

/// Extract table-centric IR JSON with default options.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed or serialized to JSON.
pub fn extract_tables_ir_json(content: &[u8]) -> Result<String, XlsxError> {
    let doc = extract_tables_ir(content)?;
    serde_json::to_string(&doc).map_err(|e| XlsxError::Xml(e.to_string()))
}

/// Extract table-centric IR JSON with explicit options.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed or serialized to JSON.
pub fn extract_tables_ir_json_with_options(
    content: &[u8],
    style_aware_values: bool,
    streaming_rows: bool,
    include_document_properties: bool,
) -> Result<String, XlsxError> {
    let doc = extract_tables_ir_with_options(
        content,
        &XlsxExtractOptions {
            text: XlsxTextOptions {
                style_aware_values,
                streaming_rows,
            },
            sheet_filter: None,
            include: XlsxIncludeOptions {
                document_properties: include_document_properties,
            },
            trim: XlsxTrimOptions { empty_edges: false },
        },
    )?;
    serde_json::to_string(&doc).map_err(|e| XlsxError::Xml(e.to_string()))
}

/// Convert column number (1-based) to Excel column name (A, B, ..., AA).
fn col_to_name(mut n: usize) -> String {
    let mut reversed = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        // rem is always 0..25, safe to truncate to u8
        #[allow(clippy::cast_possible_truncation)]
        reversed.push((b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    reversed.into_iter().rev().collect()
}

fn extract_props_map(
    package: &mut OpcPackage<'_>,
    path: &str,
) -> Result<std::collections::HashMap<String, String>, XlsxError> {
    let mut map = std::collections::HashMap::new();
    let Some(xml) = package.read_part_string(path).map_err(XlsxError::from)? else {
        return Ok(map);
    };

    let mut reader = XmlReader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut current_tag: Option<String> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                current_tag = Some(
                    std::str::from_utf8(e.name().as_ref())
                        .unwrap_or_default()
                        .to_string(),
                );
            }
            Ok(Event::Text(t)) => {
                if let Some(tag) = &current_tag {
                    let val = t
                        .unescape()
                        .map_err(|e| XlsxError::Xml(e.to_string()))?
                        .to_string();
                    map.insert(tag.clone(), val);
                }
            }
            Ok(Event::End(_)) => {
                current_tag = None;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(XlsxError::Xml(e.to_string())),
        }
    }
    Ok(map)
}

/// Resolve sheet names to their underlying XML part path using workbook relationships.
pub(crate) fn resolve_sheet_targets(
    package: &mut OpcPackage<'_>,
) -> Result<Vec<(String, String)>, XlsxError> {
    let workbook_xml = package
        .read_required_part_string("xl/workbook.xml")
        .map_err(XlsxError::from)?;

    let mut sheet_entries: Vec<(String, String)> = Vec::new();
    let mut reader = XmlReader::from_str(&workbook_xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) if e.name().as_ref() == b"sheet" => {
                let mut name: Option<String> = None;
                let mut rel_id: Option<String> = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => {
                            name = Some(
                                attr.unescape_value()
                                    .map_err(|e| XlsxError::Xml(e.to_string()))?
                                    .into_owned(),
                            );
                        }
                        b"r:id" => {
                            rel_id = Some(
                                attr.unescape_value()
                                    .map_err(|e| XlsxError::Xml(e.to_string()))?
                                    .into_owned(),
                            );
                        }
                        _ => {}
                    }
                }
                if let (Some(name), Some(rel_id)) = (name, rel_id) {
                    sheet_entries.push((name, rel_id));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(XlsxError::Xml(e.to_string())),
        }
    }

    let rels = load_relationships_for_part(package, "xl/workbook.xml").map_err(XlsxError::from)?;
    let rel_map = relationship_target_map(&rels, "xl/workbook.xml", None);

    let mut resolved = Vec::new();
    for (name, rel_id) in sheet_entries {
        let target = rel_map
            .get(&rel_id)
            .ok_or_else(|| XlsxError::Xml(format!("Missing target for {rel_id}")))?;
        resolved.push((name, target.clone()));
    }

    Ok(resolved)
}
