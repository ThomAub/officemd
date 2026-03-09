use officemd_core::opc::OpcPackage;
use quick_xml::Reader as XmlReader;
use quick_xml::events::{BytesStart, Event};

use crate::error::XlsxError;
use crate::style_format::parse_cell_ref;
use crate::table_ir::{SheetFilter, resolve_sheet_targets};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsxSheetSummary {
    pub name: String,
    pub rows: usize,
    pub cols: usize,
}

/// Inspect sheet dimensions (row and column counts) for each sheet.
///
/// # Errors
///
/// Returns an error if the XLSX archive cannot be read or sheet XML is malformed.
pub fn inspect_sheet_summaries(
    content: &[u8],
    sheet_filter: Option<&SheetFilter>,
) -> Result<Vec<XlsxSheetSummary>, XlsxError> {
    let mut package = OpcPackage::from_bytes(content).map_err(XlsxError::from)?;
    let sheet_targets = resolve_sheet_targets(&mut package)?;
    let selected_sheet_indices = sheet_filter.map_or_else(
        || (0..sheet_targets.len()).collect(),
        |filter| filter.selected_indices(&sheet_targets),
    );

    let mut summaries = Vec::with_capacity(selected_sheet_indices.len());
    for sheet_idx in selected_sheet_indices {
        let (sheet_name, sheet_path) = &sheet_targets[sheet_idx];

        let sheet_xml = package
            .read_part_bytes(sheet_path)
            .map_err(XlsxError::from)?
            .ok_or_else(|| XlsxError::Xml(format!("Missing required part: {sheet_path}")))?;
        let (rows, cols) = sheet_size_from_dimension_or_scan(sheet_xml.as_ref());
        summaries.push(XlsxSheetSummary {
            name: sheet_name.clone(),
            rows,
            cols,
        });
    }

    Ok(summaries)
}

fn sheet_size_from_dimension_or_scan(xml: &[u8]) -> (usize, usize) {
    parse_dimension_ref(xml).unwrap_or_else(|| scan_max_row_col(xml))
}

fn parse_dimension_ref(xml: &[u8]) -> Option<(usize, usize)> {
    let mut reader = XmlReader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                if local_name(e.name().as_ref()) == b"dimension" {
                    return attr_string(e, b"ref").and_then(|r| dimension_ref_to_size(&r));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {}
        }
        buf.clear();
    }

    None
}

fn dimension_ref_to_size(reference: &str) -> Option<(usize, usize)> {
    let trimmed = reference.trim();
    if trimmed.is_empty() {
        return None;
    }

    let mut bounds = trimmed.splitn(2, ':');
    let start_ref = bounds.next()?;
    let end_ref = bounds.next().unwrap_or(start_ref);
    let start = parse_a1_ref(start_ref)?;
    let end = parse_a1_ref(end_ref)?;

    let rows = end.0.abs_diff(start.0) + 1;
    let cols = end.1.abs_diff(start.1) + 1;
    Some((rows, cols))
}

fn parse_a1_ref(value: &str) -> Option<(usize, usize)> {
    if !value.contains('$') {
        return parse_cell_ref(value);
    }

    let mut normalized = String::with_capacity(value.len());
    for ch in value.chars() {
        if ch != '$' {
            normalized.push(ch);
        }
    }
    parse_cell_ref(&normalized)
}

fn scan_max_row_col(xml: &[u8]) -> (usize, usize) {
    let mut reader = XmlReader::from_reader(std::io::Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut current_row_1 = 0usize;
    let mut next_row_1 = 1usize;
    let mut next_col_1 = 1usize;
    let mut max_row_1 = 0usize;
    let mut max_col_1 = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match local_name(e.name().as_ref()) {
                b"row" => {
                    let row_1 = attr_usize(e, b"r").unwrap_or(next_row_1);
                    current_row_1 = row_1;
                    next_row_1 = row_1 + 1;
                    next_col_1 = 1;
                    max_row_1 = max_row_1.max(row_1);
                }
                b"c" => {
                    let (row_1, col_1) = attr_string(e, b"r")
                        .and_then(|r| parse_a1_ref(&r))
                        .map_or((current_row_1.max(1), next_col_1), |(row_0, col_0)| {
                            (row_0 + 1, col_0 + 1)
                        });
                    next_col_1 = col_1 + 1;
                    max_row_1 = max_row_1.max(row_1);
                    max_col_1 = max_col_1.max(col_1);
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => {
                if local_name(e.name().as_ref()) == b"row" {
                    current_row_1 = 0;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {}
        }
        buf.clear();
    }

    (max_row_1, max_col_1)
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
        if local_name(attr.key.as_ref()) == key {
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
            writer.start_file(path, options).unwrap();
            writer.write_all(contents.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
        buffer
    }

    #[test]
    fn inspects_sheet_dimensions_with_filter() {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#;
        let sheet1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;
        let sheet2 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
  </sheetData>
</worksheet>"#;
        let content = build_xlsx(vec![
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", sheet1),
            ("xl/worksheets/sheet2.xml", sheet2),
        ]);

        let mut filter = SheetFilter::default();
        filter.indices_1_based.insert(2);
        let summaries =
            inspect_sheet_summaries(&content, Some(&filter)).expect("inspect with sheet filter");
        assert_eq!(summaries.len(), 1);
        assert_eq!(summaries[0].name, "Data");
        assert_eq!(summaries[0].rows, 2);
        assert_eq!(summaries[0].cols, 2);
    }

    #[test]
    fn dimension_reference_takes_fast_path() {
        let sheet = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B3:D10"/>
</worksheet>"#;
        assert_eq!(sheet_size_from_dimension_or_scan(sheet), (8, 3));
    }
}
