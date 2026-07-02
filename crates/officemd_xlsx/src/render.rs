//! Glue to render XLSX bytes to Markdown via the shared renderer.

use crate::error::XlsxError;
use crate::table_ir::{
    XlsxExtractOptions, XlsxIncludeOptions, XlsxTrimOptions, col_to_name, extract_props_map,
    resolve_sheet_targets, trim_grid,
};
use crate::{
    sheet_reader::{SheetTextGrid, collect_sheet_text_grid},
    style_format::{StyleContext, ValueRenderMode},
};
use officemd_markdown::{MarkdownProfile, RenderOptions};
use std::collections::HashMap;
use std::fmt::Write as _;

use officemd_core::opc::OpcPackage;

/// Render XLSX bytes to Markdown using table IR (single-table per sheet for now).
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn markdown_from_bytes(content: &[u8]) -> Result<String, XlsxError> {
    markdown_from_bytes_with_options(content, RenderOptions::default())
}

/// Render XLSX bytes to Markdown with rendering options.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn markdown_from_bytes_with_options(
    content: &[u8],
    options: RenderOptions,
) -> Result<String, XlsxError> {
    markdown_from_bytes_with_extract_options(
        content,
        options,
        &XlsxExtractOptions {
            include: XlsxIncludeOptions {
                document_properties: options.include.document_properties,
            },
            trim: XlsxTrimOptions {
                empty_edges: matches!(options.markdown_profile, MarkdownProfile::LlmCompact),
            },
            ..Default::default()
        },
    )
}

/// Render XLSX bytes to Markdown with rendering and extraction options.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn markdown_from_bytes_with_extract_options(
    content: &[u8],
    render: RenderOptions,
    extract: &XlsxExtractOptions,
) -> Result<String, XlsxError> {
    let mut package = OpcPackage::from_bytes(content).map_err(XlsxError::from)?;
    let style_context = StyleContext::load(&mut package)?;
    let render_mode = if extract.text.style_aware_values {
        ValueRenderMode::StyleAware
    } else {
        ValueRenderMode::LegacyDefault
    };

    let sheet_targets = resolve_sheet_targets(&mut package)?;
    let selected_sheet_indices = extract.sheet_filter.as_ref().map_or_else(
        || (0..sheet_targets.len()).collect(),
        |filter| filter.selected_indices(&sheet_targets),
    );

    let properties = if extract.include.document_properties {
        Some((
            extract_props_map(&mut package, "docProps/core.xml")?,
            extract_props_map(&mut package, "docProps/app.xml")?,
            HashMap::new(),
        ))
    } else {
        None
    };

    let mut out = String::new();
    if render.include.frontmatter {
        write_xlsx_frontmatter(&mut out, render);
    }
    if let Some((core, app, custom)) = properties.as_ref() {
        write_properties(&mut out, core, app, custom, render);
    }

    for sheet_idx in selected_sheet_indices {
        let (name, path) = &sheet_targets[sheet_idx];
        let mut grid = collect_sheet_text_grid(&mut package, path, &style_context, render_mode)?;
        if extract.trim.empty_edges {
            trim_grid(&mut grid);
        }

        write_sheet(&mut out, name, sheet_idx, grid, render);
    }

    Ok(out)
}

fn write_xlsx_frontmatter(out: &mut String, options: RenderOptions) {
    let profile = match options.markdown_profile {
        MarkdownProfile::LlmCompact => "compact",
        MarkdownProfile::Human => "human",
    };
    let _ = write!(
        out,
        "<!-- officemd: kind=xlsx profile={profile} first_row_as_header={} formulas={} headers_footers={} properties={} -->\n\n",
        options.table.first_row_as_header,
        options.include.formulas,
        options.include.headers_footers,
        options.include.document_properties,
    );
}

fn write_sheet(
    out: &mut String,
    name: &str,
    sheet_idx: usize,
    grid: SheetTextGrid,
    options: RenderOptions,
) {
    let SheetTextGrid {
        cols: grid_cols,
        rows,
        formulas,
    } = grid;
    let cols = grid_cols.max(1);
    let row_count = rows.len().max(1);

    let _ = write!(out, "## Sheet: {name}\n\n");
    let _ = writeln!(
        out,
        "### Table {} (rows 1–{}, cols A–{})",
        sheet_idx + 1,
        row_count,
        col_to_name(cols)
    );

    if rows.is_empty() {
        write_table_row(out, std::iter::once(""));
        write_table_separator(out, 1);
    } else if options.table.first_row_as_header {
        write_table_row(out, row_values(&rows[0], cols));
        write_table_separator(out, cols);
        for row in &rows[1..] {
            write_table_row(out, row_values(row, cols));
        }
    } else {
        write_table_row(out, (1..=cols).map(|idx| format!("Col{idx}")));
        write_table_separator(out, cols);
        for row in &rows {
            write_table_row(out, row_values(row, cols));
        }
    }
    out.push('\n');

    if options.include.formulas && !formulas.is_empty() {
        if matches!(options.markdown_profile, MarkdownProfile::Human) {
            out.push_str("### Formulas\n");
        }
        for (i, note) in formulas.iter().enumerate() {
            let formula_body = note
                .formula
                .strip_prefix('=')
                .unwrap_or(note.formula.as_str());
            if matches!(options.markdown_profile, MarkdownProfile::Human) {
                let _ = writeln!(
                    out,
                    "[^f{}]: {} = `={}`",
                    i + 1,
                    note.cell_ref,
                    formula_body
                );
            } else {
                let _ = writeln!(out, "{}=`={}`", note.cell_ref, formula_body);
            }
        }
        out.push('\n');
    }
}

fn row_values(row: &[String], cols: usize) -> impl Iterator<Item = &str> {
    (0..cols).map(|idx| row.get(idx).map_or("", String::as_str))
}

fn write_table_row<'a, I, S>(out: &mut String, cells: I)
where
    I: IntoIterator<Item = S>,
    S: AsRef<str> + 'a,
{
    out.push('|');
    for cell in cells {
        out.push(' ');
        write_escaped_pipes(out, cell.as_ref());
        out.push(' ');
        out.push('|');
    }
    out.push('\n');
}

fn write_table_separator(out: &mut String, cols: usize) {
    out.push('|');
    for _ in 0..cols {
        out.push_str(" --- |");
    }
    out.push('\n');
}

fn write_properties(
    out: &mut String,
    core: &HashMap<String, String>,
    app: &HashMap<String, String>,
    custom: &HashMap<String, String>,
    options: RenderOptions,
) {
    if core.is_empty() && app.is_empty() && custom.is_empty() {
        return;
    }

    let mut entries = core
        .iter()
        .chain(app.iter())
        .chain(custom.iter())
        .map(|(k, v)| (k.as_str(), v.as_str()))
        .collect::<Vec<_>>();
    entries.sort_unstable_by(|(ka, va), (kb, vb)| ka.cmp(kb).then_with(|| va.cmp(vb)));

    if matches!(options.markdown_profile, MarkdownProfile::Human) {
        out.push_str("### Document Properties\n");
        for (key, value) in entries {
            let _ = write!(out, "- {key}: ");
            write_escaped_pipes(out, value);
            out.push('\n');
        }
        out.push_str("\n---\n\n");
    } else {
        out.push_str("properties: ");
        for (idx, (key, value)) in entries.iter().enumerate() {
            if idx > 0 {
                out.push_str("; ");
            }
            let _ = write!(out, "{key}=");
            write_escaped_pipes(out, value);
        }
        out.push_str("\n\n");
    }
}

fn write_escaped_pipes(out: &mut String, value: &str) {
    if !value.contains('|') {
        out.push_str(value);
        return;
    }
    for ch in value.chars() {
        if ch == '|' {
            out.push('\\');
        }
        out.push(ch);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    fn build_xlsx(parts: Vec<(&str, &str)>) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut writer = ZipWriter::new(std::io::Cursor::new(&mut buffer));
        let options: FileOptions<'_, ()> = FileOptions::default();

        for (path, contents) in parts {
            writer.start_file(path, options).expect("start file");
            writer
                .write_all(contents.as_bytes())
                .expect("write contents");
        }

        writer.finish().expect("finish zip");
        buffer
    }

    fn minimal_xlsx_with_doc_props() -> Vec<u8> {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1"><c r="A1"><v>1</v></c></row>
    </sheetData>
</worksheet>"#;

        let core = r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Quarterly Results</dc:title>
</cp:coreProperties>"#;

        build_xlsx(vec![
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", sheet),
            ("docProps/core.xml", core),
        ])
    }

    #[test]
    fn markdown_includes_document_properties_when_requested() {
        let bytes = minimal_xlsx_with_doc_props();
        let markdown = markdown_from_bytes_with_options(
            &bytes,
            RenderOptions {
                include: officemd_markdown::RenderIncludeOptions {
                    document_properties: true,
                    ..Default::default()
                },
                markdown_profile: officemd_markdown::MarkdownProfile::Human,
                ..Default::default()
            },
        )
        .expect("render markdown");

        assert!(markdown.contains("### Document Properties"));
        assert!(markdown.contains("Quarterly Results"));
    }

    #[test]
    fn markdown_omits_document_properties_by_default() {
        let bytes = minimal_xlsx_with_doc_props();
        let markdown = markdown_from_bytes(&bytes).expect("render markdown");

        assert!(!markdown.contains("### Document Properties"));
        assert!(!markdown.contains("Quarterly Results"));
    }
}
