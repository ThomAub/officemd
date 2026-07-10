use std::{
    collections::BTreeMap,
    io::{Cursor, Read, Write},
    sync::LazyLock,
};

use regex::Regex;

use crate::DocxError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DocxLocator {
    Paragraph {
        part: String,
        paragraph_index: u32,
    },
    TableCell {
        part: String,
        table_index: u32,
        row_index: u32,
        column_index: u32,
    },
}

/// Replace text only inside one stable DOCX locator.
///
/// Returns the rewritten package and replacement count.
///
/// # Errors
///
/// Returns an error when the package, part, or locator cannot be resolved.
pub fn replace_locator_text(
    content: &[u8],
    locator: &DocxLocator,
    expected_text: &str,
    replacement: &str,
    preserve_formatting: bool,
) -> Result<(Vec<u8>, usize), DocxError> {
    let mut parts = read_zip_parts(content)?;
    let part = locator_part(locator);
    let xml = parts
        .get(&part)
        .ok_or_else(|| DocxError::MissingPart(part.clone()))?;
    let xml = String::from_utf8_lossy(xml);
    let (updated, replacements) = match locator {
        DocxLocator::Paragraph {
            paragraph_index, ..
        } => replace_paragraph(
            &xml,
            *paragraph_index,
            expected_text,
            replacement,
            preserve_formatting,
        ),
        DocxLocator::TableCell {
            table_index,
            row_index,
            column_index,
            ..
        } => replace_table_cell(
            &xml,
            *table_index,
            *row_index,
            *column_index,
            expected_text,
            replacement,
            preserve_formatting,
        ),
    };
    if replacements == 0 {
        return Ok((content.to_vec(), 0));
    }
    parts.insert(part, updated.into_bytes());
    Ok((write_zip_parts(parts)?, replacements))
}

fn locator_part(locator: &DocxLocator) -> String {
    let part = match locator {
        DocxLocator::Paragraph { part, .. } | DocxLocator::TableCell { part, .. } => part,
    };
    let normalized = part.to_ascii_lowercase();
    match normalized.as_str() {
        "body" | "document" => "word/document.xml".to_string(),
        "footnotes" => "word/footnotes.xml".to_string(),
        "endnotes" => "word/endnotes.xml".to_string(),
        path if path.starts_with("word/") => part.to_string(),
        _ => format!("word/{part}.xml"),
    }
}

fn replace_paragraph(
    xml: &str,
    paragraph_index: u32,
    expected_text: &str,
    replacement: &str,
    preserve_formatting: bool,
) -> (String, usize) {
    let table_ranges = WORD_TABLE_RE
        .find_iter(xml)
        .map(|table| (table.start(), table.end()))
        .collect::<Vec<_>>();
    let mut visible_index = 0u32;
    let mut output = String::with_capacity(xml.len());
    let mut last = 0usize;
    let mut replacements = 0usize;
    for paragraph in WORD_PARAGRAPH_RE.find_iter(xml) {
        output.push_str(&xml[last..paragraph.start()]);
        let block = paragraph.as_str();
        let inside_table = table_ranges
            .iter()
            .any(|(start, end)| paragraph.start() >= *start && paragraph.end() <= *end);
        if inside_table || docx_text(block).trim().is_empty() {
            output.push_str(block);
            last = paragraph.end();
            continue;
        }
        let is_target = visible_index == paragraph_index;
        visible_index = visible_index.saturating_add(1);
        if is_target {
            let (rewritten, count) =
                replace_text_nodes(block, expected_text, replacement, preserve_formatting);
            output.push_str(&rewritten);
            replacements = count;
        } else {
            output.push_str(block);
        }
        last = paragraph.end();
    }
    output.push_str(&xml[last..]);
    (output, replacements)
}

fn replace_table_cell(
    xml: &str,
    table_index: u32,
    row_index: u32,
    column_index: u32,
    expected_text: &str,
    replacement: &str,
    preserve_formatting: bool,
) -> (String, usize) {
    let mut current_table = 0u32;
    replace_indexed_match(xml, &WORD_TABLE_RE, |table| {
        let is_target_table = current_table == table_index;
        current_table = current_table.saturating_add(1);
        if !is_target_table {
            return None;
        }
        let mut current_row = 0u32;
        Some(replace_indexed_match(table, &WORD_ROW_RE, |row| {
            let is_target_row = current_row == row_index;
            current_row = current_row.saturating_add(1);
            if !is_target_row {
                return None;
            }
            let mut current_col = 0u32;
            Some(replace_indexed_match(row, &WORD_CELL_RE, |cell| {
                let is_target_col = current_col == column_index;
                current_col = current_col.saturating_add(1);
                is_target_col.then(|| {
                    replace_text_nodes(cell, expected_text, replacement, preserve_formatting)
                })
            }))
        }))
    })
}

fn replace_indexed_match<F>(xml: &str, regex: &Regex, mut replacer: F) -> (String, usize)
where
    F: FnMut(&str) -> Option<(String, usize)>,
{
    let mut output = String::with_capacity(xml.len());
    let mut last = 0usize;
    let mut replacements = 0usize;
    for matched in regex.find_iter(xml) {
        output.push_str(&xml[last..matched.start()]);
        if replacements == 0
            && let Some((updated, count)) = replacer(matched.as_str())
        {
            output.push_str(&updated);
            replacements = replacements.saturating_add(count);
        } else {
            output.push_str(matched.as_str());
        }
        last = matched.end();
    }
    output.push_str(&xml[last..]);
    (output, replacements)
}

fn replace_text_nodes(
    xml: &str,
    expected_text: &str,
    replacement: &str,
    preserve_formatting: bool,
) -> (String, usize) {
    let nodes = WORD_TEXT_NODE_RE
        .captures_iter(xml)
        .filter_map(|captures| captures.get(1))
        .map(|text| {
            let decoded = quick_xml::escape::unescape(text.as_str())
                .map_or_else(|_| text.as_str().to_string(), |value| value.into_owned());
            (text.start(), text.end(), decoded)
        })
        .collect::<Vec<_>>();
    if nodes.is_empty() || expected_text.is_empty() {
        return (xml.to_string(), 0);
    }
    let combined = nodes
        .iter()
        .map(|(_, _, text)| text.as_str())
        .collect::<String>();
    let replacements = combined.match_indices(expected_text).count();
    if replacements == 0 {
        return (xml.to_string(), 0);
    }
    let updated = combined.replace(expected_text, replacement);
    let rewritten_nodes = distribute_text(&nodes, &updated, preserve_formatting);
    let mut output = String::with_capacity(xml.len() + updated.len());
    let mut last = 0usize;
    for ((start, end, _), rewritten) in nodes.iter().zip(rewritten_nodes) {
        output.push_str(&xml[last..*start]);
        output.push_str(&officemd_core::opc::xml_escape_text(&rewritten));
        last = *end;
    }
    output.push_str(&xml[last..]);
    (output, replacements)
}

fn distribute_text(
    nodes: &[(usize, usize, String)],
    updated: &str,
    preserve_formatting: bool,
) -> Vec<String> {
    if !preserve_formatting {
        let mut values = vec![String::new(); nodes.len()];
        values[0] = updated.to_string();
        return values;
    }
    let chars = updated.chars().collect::<Vec<_>>();
    let mut offset = 0usize;
    nodes
        .iter()
        .enumerate()
        .map(|(index, (_, _, original))| {
            let end = if index + 1 == nodes.len() {
                chars.len()
            } else {
                offset
                    .saturating_add(original.chars().count())
                    .min(chars.len())
            };
            let value = chars[offset..end].iter().collect::<String>();
            offset = end;
            value
        })
        .collect()
}

fn docx_text(xml: &str) -> String {
    WORD_TEXT_NODE_RE
        .captures_iter(xml)
        .filter_map(|captures| captures.get(1).map(|text| text.as_str()))
        .collect()
}

fn read_zip_parts(content: &[u8]) -> Result<BTreeMap<String, Vec<u8>>, DocxError> {
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

fn write_zip_parts(parts: BTreeMap<String, Vec<u8>>) -> Result<Vec<u8>, DocxError> {
    let mut output = Vec::new();
    {
        let mut writer = zip::ZipWriter::new(Cursor::new(&mut output));
        let options: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();
        for (name, bytes) in parts {
            writer.start_file(name, options)?;
            writer.write_all(&bytes)?;
        }
        writer.finish()?;
    }
    Ok(output)
}

static WORD_PARAGRAPH_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:p(?:\s+[^>]*)?>.*?</w:p>").unwrap());
static WORD_TABLE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:tbl(?:\s+[^>]*)?>.*?</w:tbl>").unwrap());
static WORD_ROW_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:tr(?:\s+[^>]*)?>.*?</w:tr>").unwrap());
static WORD_CELL_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:tc(?:\s+[^>]*)?>.*?</w:tc>").unwrap());
static WORD_TEXT_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:t(?:\s+[^>]*)?>(.*?)</w:t>").unwrap());
