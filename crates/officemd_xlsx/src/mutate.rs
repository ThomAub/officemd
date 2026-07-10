use std::{
    collections::BTreeMap,
    io::{Cursor, Read, Write},
};

use officemd_core::opc::OpcPackage;
use quick_xml::{
    Reader, Writer,
    events::{BytesEnd, BytesStart, BytesText, Event},
};

use crate::{XlsxError, style_format::parse_cell_ref, table_ir::resolve_sheet_targets};

#[derive(Debug, Clone, PartialEq)]
pub enum XlsxCellUpdate {
    Text(String),
    Number(f64),
    Bool(bool),
    Blank,
    Formula(String),
}

/// Set one absolute A1-addressed cell while preserving its existing style.
///
/// # Errors
///
/// Returns an error when the package, sheet, address, or worksheet XML is invalid.
pub fn set_cell(
    content: &[u8],
    sheet: &str,
    address: &str,
    update: &XlsxCellUpdate,
) -> Result<Vec<u8>, XlsxError> {
    let (target_row, target_col) = parse_cell_ref(address)
        .ok_or_else(|| XlsxError::Xml(format!("invalid cell address: {address}")))?;
    if matches!(update, XlsxCellUpdate::Number(value) if !value.is_finite()) {
        return Err(XlsxError::Xml("cell number must be finite".to_string()));
    }
    if matches!(update, XlsxCellUpdate::Formula(formula) if formula.trim().trim_start_matches('=').is_empty())
    {
        return Err(XlsxError::Xml("cell formula must not be empty".to_string()));
    }

    let sheet_path = {
        let mut package = OpcPackage::from_bytes(content)?;
        resolve_sheet_targets(&mut package)?
            .into_iter()
            .find(|(candidate, _)| candidate == sheet)
            .map(|(_, path)| path)
            .ok_or_else(|| XlsxError::Xml(format!("sheet not found: {sheet}")))?
    };
    let mut parts = read_zip_parts(content)?;
    let worksheet = parts
        .get(&sheet_path)
        .ok_or_else(|| XlsxError::MissingPart(sheet_path.clone()))?;
    let rewritten = rewrite_worksheet(worksheet, address, target_row + 1, target_col + 1, update)?;
    parts.insert(sheet_path, rewritten);
    write_zip_parts(parts)
}

#[allow(clippy::too_many_lines)]
fn rewrite_worksheet(
    xml: &[u8],
    address: &str,
    target_row: usize,
    target_col: usize,
    update: &XlsxCellUpdate,
) -> Result<Vec<u8>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::with_capacity(xml.len() + 128)));
    let mut buffer = Vec::new();
    let mut in_sheet_data = false;
    let mut target_row_seen = false;
    let mut cell_written = false;
    let mut current_row = None;
    let mut next_row = 1usize;
    let mut next_col = 1usize;
    let mut skipped_cell_depth = 0usize;

    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| XlsxError::Xml(error.to_string()))?;

        if skipped_cell_depth > 0 {
            match &event {
                Event::Start(_) => skipped_cell_depth += 1,
                Event::End(_) => skipped_cell_depth -= 1,
                Event::Eof => {
                    return Err(XlsxError::Xml(
                        "unexpected end of worksheet while replacing cell".to_string(),
                    ));
                }
                _ => {}
            }
            buffer.clear();
            continue;
        }

        match event {
            Event::Empty(ref sheet_data)
                if local_name(sheet_data.name().as_ref()) == b"sheetData" =>
            {
                writer.write_event(Event::Start(sheet_data.to_owned()))?;
                write_new_row(&mut writer, address, target_row, update)?;
                writer.write_event(Event::End(BytesEnd::new("sheetData")))?;
                target_row_seen = true;
                cell_written = true;
            }
            Event::Start(ref start) if local_name(start.name().as_ref()) == b"sheetData" => {
                in_sheet_data = true;
                writer.write_event(event.into_owned())?;
            }
            Event::End(ref end) if local_name(end.name().as_ref()) == b"sheetData" => {
                if !target_row_seen {
                    write_new_row(&mut writer, address, target_row, update)?;
                    target_row_seen = true;
                    cell_written = true;
                }
                in_sheet_data = false;
                writer.write_event(event.into_owned())?;
            }
            Event::Start(ref row) if in_sheet_data && local_name(row.name().as_ref()) == b"row" => {
                let row_number = attr_usize(row, b"r").unwrap_or(next_row);
                if !target_row_seen && row_number > target_row {
                    write_new_row(&mut writer, address, target_row, update)?;
                    target_row_seen = true;
                    cell_written = true;
                }
                if row_number == target_row {
                    target_row_seen = true;
                }
                current_row = Some(row_number);
                next_row = row_number + 1;
                next_col = 1;
                writer.write_event(event.into_owned())?;
            }
            Event::Empty(ref row) if in_sheet_data && local_name(row.name().as_ref()) == b"row" => {
                let row_number = attr_usize(row, b"r").unwrap_or(next_row);
                if !target_row_seen && row_number > target_row {
                    write_new_row(&mut writer, address, target_row, update)?;
                    target_row_seen = true;
                    cell_written = true;
                }
                if row_number == target_row {
                    write_row_start(&mut writer, row)?;
                    write_cell(&mut writer, address, None, update)?;
                    writer.write_event(Event::End(BytesEnd::new("row")))?;
                    target_row_seen = true;
                    cell_written = true;
                } else {
                    writer.write_event(event.into_owned())?;
                }
                next_row = row_number + 1;
            }
            Event::Start(ref cell)
                if current_row == Some(target_row) && local_name(cell.name().as_ref()) == b"c" =>
            {
                let (cell_row, cell_col) = cell_position(cell, current_row.unwrap_or(1), next_col);
                next_col = cell_col + 1;
                if !cell_written && cell_col > target_col {
                    write_cell(&mut writer, address, None, update)?;
                    cell_written = true;
                }
                if cell_row == target_row && cell_col == target_col {
                    write_cell(&mut writer, address, Some(cell), update)?;
                    cell_written = true;
                    skipped_cell_depth = 1;
                } else {
                    writer.write_event(event.into_owned())?;
                }
            }
            Event::Empty(ref cell)
                if current_row == Some(target_row) && local_name(cell.name().as_ref()) == b"c" =>
            {
                let (cell_row, cell_col) = cell_position(cell, current_row.unwrap_or(1), next_col);
                next_col = cell_col + 1;
                if !cell_written && cell_col > target_col {
                    write_cell(&mut writer, address, None, update)?;
                    cell_written = true;
                }
                if cell_row == target_row && cell_col == target_col {
                    write_cell(&mut writer, address, Some(cell), update)?;
                    cell_written = true;
                } else {
                    writer.write_event(event.into_owned())?;
                }
            }
            Event::End(ref row)
                if current_row == Some(target_row) && local_name(row.name().as_ref()) == b"row" =>
            {
                if !cell_written {
                    write_cell(&mut writer, address, None, update)?;
                    cell_written = true;
                }
                current_row = None;
                writer.write_event(event.into_owned())?;
            }
            Event::End(ref row) if local_name(row.name().as_ref()) == b"row" => {
                current_row = None;
                writer.write_event(event.into_owned())?;
            }
            Event::Eof => break,
            _ => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }

    if !in_sheet_data && !target_row_seen {
        return Err(XlsxError::Xml(
            "worksheet does not contain sheetData".to_string(),
        ));
    }
    Ok(writer.into_inner().into_inner())
}

fn write_new_row(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    address: &str,
    row: usize,
    update: &XlsxCellUpdate,
) -> Result<(), XlsxError> {
    let mut start = BytesStart::new("row");
    let row_value = row.to_string();
    start.push_attribute(("r", row_value.as_str()));
    writer.write_event(Event::Start(start))?;
    write_cell(writer, address, None, update)?;
    writer.write_event(Event::End(BytesEnd::new("row")))?;
    Ok(())
}

fn write_row_start(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    existing: &BytesStart<'_>,
) -> Result<(), XlsxError> {
    writer.write_event(Event::Start(existing.to_owned()))?;
    Ok(())
}

fn write_cell(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    address: &str,
    existing: Option<&BytesStart<'_>>,
    update: &XlsxCellUpdate,
) -> Result<(), XlsxError> {
    let mut attributes = Vec::new();
    if let Some(existing) = existing {
        for attribute in existing.attributes().with_checks(false).flatten() {
            let Ok(key) = String::from_utf8(attribute.key.as_ref().to_vec()) else {
                continue;
            };
            if matches!(local_name(key.as_bytes()), b"r" | b"t") {
                continue;
            }
            let value = attribute.unescape_value().map_or_else(
                |_| String::from_utf8_lossy(&attribute.value).into_owned(),
                |value| value.into_owned(),
            );
            attributes.push((key, value));
        }
    }
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", address));
    for (key, value) in &attributes {
        cell.push_attribute((key.as_str(), value.as_str()));
    }
    match update {
        XlsxCellUpdate::Text(_) => cell.push_attribute(("t", "inlineStr")),
        XlsxCellUpdate::Bool(_) => cell.push_attribute(("t", "b")),
        XlsxCellUpdate::Number(_) | XlsxCellUpdate::Blank | XlsxCellUpdate::Formula(_) => {}
    }
    writer.write_event(Event::Start(cell))?;
    match update {
        XlsxCellUpdate::Text(value) => {
            writer.write_event(Event::Start(BytesStart::new("is")))?;
            writer.write_event(Event::Start(BytesStart::new("t")))?;
            writer.write_event(Event::Text(BytesText::new(value)))?;
            writer.write_event(Event::End(BytesEnd::new("t")))?;
            writer.write_event(Event::End(BytesEnd::new("is")))?;
        }
        XlsxCellUpdate::Number(value) => write_text_element(writer, "v", &value.to_string())?,
        XlsxCellUpdate::Bool(value) => {
            write_text_element(writer, "v", if *value { "1" } else { "0" })?
        }
        XlsxCellUpdate::Blank => {}
        XlsxCellUpdate::Formula(formula) => {
            write_text_element(writer, "f", formula.trim().trim_start_matches('='))?;
        }
    }
    writer.write_event(Event::End(BytesEnd::new("c")))?;
    Ok(())
}

fn write_text_element(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    name: &str,
    value: &str,
) -> Result<(), XlsxError> {
    writer.write_event(Event::Start(BytesStart::new(name)))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new(name)))?;
    Ok(())
}

fn cell_position(cell: &BytesStart<'_>, current_row: usize, next_col: usize) -> (usize, usize) {
    attr_string(cell, b"r")
        .and_then(|value| parse_cell_ref(&value))
        .map_or((current_row, next_col), |(row, col)| (row + 1, col + 1))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .rposition(|byte| matches!(*byte, b':' | b'}'))
        .map_or(name, |index| &name[index + 1..])
}

fn attr_string(element: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    element
        .attributes()
        .with_checks(false)
        .flatten()
        .find_map(|attribute| {
            (local_name(attribute.key.as_ref()) == key).then(|| {
                attribute.unescape_value().map_or_else(
                    |_| String::from_utf8_lossy(&attribute.value).into_owned(),
                    |value| value.into_owned(),
                )
            })
        })
}

fn attr_usize(element: &BytesStart<'_>, key: &[u8]) -> Option<usize> {
    attr_string(element, key)?.parse().ok()
}

fn read_zip_parts(content: &[u8]) -> Result<BTreeMap<String, Vec<u8>>, XlsxError> {
    let mut archive = zip::ZipArchive::new(Cursor::new(content))?;
    let mut parts = BTreeMap::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index)?;
        if file.is_dir() {
            continue;
        }
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)?;
        parts.insert(file.name().to_string(), bytes);
    }
    Ok(parts)
}

fn write_zip_parts(parts: BTreeMap<String, Vec<u8>>) -> Result<Vec<u8>, XlsxError> {
    let mut output = Vec::new();
    {
        let mut archive = zip::ZipWriter::new(Cursor::new(&mut output));
        let options: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();
        for (name, bytes) in parts {
            archive.start_file(name, options)?;
            archive.write_all(&bytes)?;
        }
        archive.finish()?;
    }
    Ok(output)
}
