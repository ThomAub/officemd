//! PPTX IR extraction helpers.

use std::collections::{HashMap, HashSet};

use officemd_core::ir::{
    Block, CommentNote, DocumentKind, DocumentProperties, Hyperlink, Inline, OoxmlDocument,
    Paragraph, Slide, Table, TableCell,
};
use officemd_core::opc::{
    OpcPackage, load_relationships_for_part, relationship_target_map, resolve_relationship_target,
};
use officemd_core::rels::Relationship;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::error::PptxError;

const REL_NOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_COMMENT_AUTHORS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";

#[derive(Debug, Clone, Default)]
pub struct PptxExtractOptions {
    pub slide_numbers: Option<HashSet<usize>>,
}

#[derive(Debug, Clone)]
struct SlideSelection {
    allowed: HashSet<usize>,
    max: usize,
}

impl SlideSelection {
    fn new(allowed: HashSet<usize>) -> Self {
        let max = allowed.iter().copied().max().unwrap_or(0);
        Self { allowed, max }
    }

    fn contains(&self, slide_number: usize) -> bool {
        self.allowed.contains(&slide_number)
    }
}

/// Extract PPTX IR.
///
/// # Errors
///
/// Returns `PptxError` if the ZIP archive is invalid, required parts are
/// missing, or XML within the presentation is malformed.
pub fn extract_ir(content: &[u8]) -> Result<OoxmlDocument, PptxError> {
    extract_ir_with_options(content, PptxExtractOptions::default())
}

/// Extract PPTX IR with explicit options.
///
/// # Errors
///
/// Returns `PptxError` if the ZIP archive is invalid, required parts are
/// missing, or XML within the presentation is malformed.
pub fn extract_ir_with_options(
    content: &[u8],
    options: PptxExtractOptions,
) -> Result<OoxmlDocument, PptxError> {
    let mut package = OpcPackage::from_bytes(content).map_err(PptxError::from)?;
    let slide_selection = options.slide_numbers.map(SlideSelection::new);

    let presentation_xml = read_required_part(&mut package, "ppt/presentation.xml")?;
    let slide_rids = collect_slide_rids(&presentation_xml, slide_selection.as_ref())?;

    let presentation_rels = load_relationships_for_part(&mut package, "ppt/presentation.xml")
        .map_err(PptxError::from)?;
    let presentation_rels_map = build_rel_map(&presentation_rels, "ppt/presentation.xml");

    let authors_path = find_rel_target(
        &presentation_rels,
        REL_COMMENT_AUTHORS,
        "ppt/presentation.xml",
    );
    let authors = if let Some(path) = authors_path {
        read_part(&mut package, &path)?
            .map(|xml| parse_comment_authors(&xml))
            .unwrap_or_default()
    } else {
        HashMap::new()
    };

    let properties = extract_properties(&mut package)?;
    let mut slides = Vec::with_capacity(slide_rids.len());

    for (slide_number, rid) in slide_rids {
        let Some(slide_path) = presentation_rels_map.get(&rid).cloned() else {
            return Err(PptxError::MissingPart(format!(
                "presentation rel for {rid}"
            )));
        };
        let slide_xml = read_required_part(&mut package, &slide_path)?;
        let slide_rels =
            load_relationships_for_part(&mut package, &slide_path).map_err(PptxError::from)?;
        let slide_rels_map = build_rel_map(&slide_rels, &slide_path);

        let (title, blocks) = parse_blocks(&slide_xml, &slide_rels_map, true, true)?;

        let notes = extract_notes(&mut package, &slide_rels, &slide_path)?;

        let comments = extract_comments(&mut package, &slide_rels, &slide_path, &authors)?;

        slides.push(Slide {
            number: slide_number,
            title,
            blocks,
            notes,
            comments,
        });
    }

    Ok(OoxmlDocument {
        kind: DocumentKind::Pptx,
        properties,
        slides,
        ..Default::default()
    })
}

/// Extract PPTX IR JSON (slides, notes, comments, tables).
///
/// # Errors
///
/// Returns `PptxError` if extraction fails or the document cannot be
/// serialised to JSON.
pub fn extract_ir_json(content: &[u8]) -> Result<String, PptxError> {
    let doc = extract_ir(content)?;
    serde_json::to_string(&doc).map_err(|e| PptxError::Xml(e.to_string()))
}

fn extract_notes(
    package: &mut OpcPackage<'_>,
    slide_rels: &[Relationship],
    slide_part: &str,
) -> Result<Option<Vec<Paragraph>>, PptxError> {
    let Some(notes_path) = find_rel_target(slide_rels, REL_NOTES, slide_part) else {
        return Ok(None);
    };
    let notes_xml = read_required_part(package, &notes_path)?;
    let notes_rels = load_relationships_for_part(package, &notes_path).map_err(PptxError::from)?;
    let notes_rels_map = build_rel_map(&notes_rels, &notes_path);

    let (_title, blocks) = parse_blocks(&notes_xml, &notes_rels_map, false, false)?;
    let paragraphs = blocks
        .into_iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect::<Vec<_>>();

    if paragraphs.is_empty() {
        Ok(None)
    } else {
        Ok(Some(paragraphs))
    }
}

fn extract_comments(
    package: &mut OpcPackage<'_>,
    slide_rels: &[Relationship],
    slide_part: &str,
    authors: &HashMap<String, String>,
) -> Result<Vec<CommentNote>, PptxError> {
    let Some(comments_path) = find_rel_target(slide_rels, REL_COMMENTS, slide_part) else {
        return Ok(Vec::new());
    };
    let comments_xml = read_required_part(package, &comments_path)?;
    Ok(parse_comments(&comments_xml, authors))
}

fn read_required_part(package: &mut OpcPackage<'_>, path: &str) -> Result<String, PptxError> {
    read_part(package, path)?.ok_or_else(|| PptxError::MissingPart(path.to_string()))
}

fn extract_properties(
    package: &mut OpcPackage<'_>,
) -> Result<Option<DocumentProperties>, PptxError> {
    officemd_core::opc::extract_properties(package).map_err(PptxError::from)
}

fn read_part(package: &mut OpcPackage<'_>, path: &str) -> Result<Option<String>, PptxError> {
    package.read_part_string(path).map_err(PptxError::from)
}

fn collect_slide_rids(
    xml: &str,
    selection: Option<&SlideSelection>,
) -> Result<Vec<(usize, String)>, PptxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut slide_entries = Vec::new();
    let mut slide_number = 0usize;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"sldId" {
                    slide_number += 1;
                    if let Some(sel) = selection {
                        if sel.max > 0 && slide_number > sel.max {
                            break;
                        }
                        if !sel.contains(slide_number) {
                            continue;
                        }
                    }
                    if let Some(rid) = attr_value_rid(e) {
                        slide_entries.push((slide_number, rid));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PptxError::Xml(e.to_string())),
        }
    }

    Ok(slide_entries)
}

#[allow(clippy::too_many_lines)]
fn parse_blocks(
    xml: &str,
    rels_map: &HashMap<String, String>,
    detect_title: bool,
    include_tables: bool,
) -> Result<(Option<String>, Vec<Block>), PptxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut blocks = Vec::new();
    let mut table_rows: Vec<Vec<TableCell>> = Vec::new();
    let mut current_row: Vec<TableCell> = Vec::new();
    let mut current_cell_paragraphs: Vec<Paragraph> = Vec::new();

    let mut current_paragraph: Option<Paragraph> = None;
    let mut in_table = false;
    let mut in_row = false;
    let mut in_cell = false;
    let mut in_run = false;
    let mut in_text = false;
    let mut current_run_text = String::new();
    let mut current_run_link: Option<String> = None;

    let mut shape_depth = 0usize;
    let mut shape_is_title = false;
    let mut title: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                handle_start(
                    e,
                    false,
                    include_tables,
                    &mut in_table,
                    &mut in_row,
                    &mut in_cell,
                    &mut table_rows,
                    &mut current_row,
                    &mut current_cell_paragraphs,
                    &mut current_paragraph,
                    &mut in_run,
                    &mut in_text,
                    &mut current_run_text,
                    &mut current_run_link,
                    &mut shape_depth,
                    &mut shape_is_title,
                );
            }
            Ok(Event::Empty(ref e)) => {
                handle_start(
                    e,
                    true,
                    include_tables,
                    &mut in_table,
                    &mut in_row,
                    &mut in_cell,
                    &mut table_rows,
                    &mut current_row,
                    &mut current_cell_paragraphs,
                    &mut current_paragraph,
                    &mut in_run,
                    &mut in_text,
                    &mut current_run_text,
                    &mut current_run_link,
                    &mut shape_depth,
                    &mut shape_is_title,
                );
                handle_end(
                    e.name().as_ref(),
                    include_tables,
                    &mut in_table,
                    &mut in_row,
                    &mut in_cell,
                    &mut table_rows,
                    &mut current_row,
                    &mut current_cell_paragraphs,
                    &mut current_paragraph,
                    &mut in_run,
                    &mut in_text,
                    &mut current_run_text,
                    &mut current_run_link,
                    &mut blocks,
                    rels_map,
                    detect_title,
                    &mut title,
                    &mut shape_depth,
                    &mut shape_is_title,
                );
            }
            Ok(Event::End(ref e)) => {
                handle_end(
                    e.name().as_ref(),
                    include_tables,
                    &mut in_table,
                    &mut in_row,
                    &mut in_cell,
                    &mut table_rows,
                    &mut current_row,
                    &mut current_cell_paragraphs,
                    &mut current_paragraph,
                    &mut in_run,
                    &mut in_text,
                    &mut current_run_text,
                    &mut current_run_link,
                    &mut blocks,
                    rels_map,
                    detect_title,
                    &mut title,
                    &mut shape_depth,
                    &mut shape_is_title,
                );
            }
            Ok(Event::Text(t)) => {
                if in_text {
                    let text = t.unescape().map_err(|e| PptxError::Xml(e.to_string()))?;
                    if in_run {
                        current_run_text.push_str(&text);
                    } else if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.inlines.push(Inline::Text(text.into_owned()));
                    }
                }
            }
            Ok(Event::CData(t)) => {
                if in_text {
                    let text = decode_cdata(&t)?;
                    if in_run {
                        current_run_text.push_str(&text);
                    } else if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.inlines.push(Inline::Text(text));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PptxError::Xml(e.to_string())),
        }
    }

    if detect_title && title.is_none() {
        for block in &blocks {
            if let Block::Paragraph(paragraph) = block {
                let text = paragraph_plain_text(paragraph);
                if !text.is_empty() {
                    title = Some(text);
                    break;
                }
            }
        }
    }

    Ok((title, blocks))
}

#[allow(clippy::too_many_arguments)]
fn handle_start(
    e: &BytesStart<'_>,
    is_empty: bool,
    include_tables: bool,
    in_table: &mut bool,
    in_row: &mut bool,
    in_cell: &mut bool,
    table_rows: &mut Vec<Vec<TableCell>>,
    current_row: &mut Vec<TableCell>,
    current_cell_paragraphs: &mut Vec<Paragraph>,
    current_paragraph: &mut Option<Paragraph>,
    in_run: &mut bool,
    in_text: &mut bool,
    current_run_text: &mut String,
    current_run_link: &mut Option<String>,
    shape_depth: &mut usize,
    shape_is_title: &mut bool,
) {
    let name = e.name();
    let name = local_name(name.as_ref());
    match name {
        b"sp" => {
            *shape_depth += 1;
            if *shape_depth == 1 {
                *shape_is_title = false;
            }
        }
        b"ph" => {
            if *shape_depth > 0
                && let Some(value) = attr_value_exact(e, b"type")
                && (value == "title" || value == "ctrTitle")
            {
                *shape_is_title = true;
            }
        }
        b"tbl" if include_tables => {
            *in_table = true;
            table_rows.clear();
        }
        b"tr" if *in_table && include_tables => {
            *in_row = true;
            current_row.clear();
        }
        b"tc" if *in_table && include_tables => {
            *in_cell = true;
            current_cell_paragraphs.clear();
        }
        b"p" => {
            *current_paragraph = Some(Paragraph {
                inlines: Vec::new(),
            });
        }
        b"r" => {
            *in_run = true;
            current_run_text.clear();
            *current_run_link = None;
        }
        b"t" => {
            *in_text = true;
            if is_empty {
                *in_text = false;
            }
        }
        b"br" => {
            if let Some(paragraph) = current_paragraph.as_mut() {
                if *in_run {
                    current_run_text.push('\n');
                } else {
                    paragraph.inlines.push(Inline::Text("\n".to_string()));
                }
            }
        }
        b"hlinkClick" => {
            if *in_run && let Some(rid) = attr_value_rid(e) {
                *current_run_link = Some(rid);
            }
        }
        _ => {}
    }
}

#[allow(clippy::too_many_arguments)]
fn handle_end(
    name: &[u8],
    include_tables: bool,
    in_table: &mut bool,
    in_row: &mut bool,
    in_cell: &mut bool,
    table_rows: &mut Vec<Vec<TableCell>>,
    current_row: &mut Vec<TableCell>,
    current_cell_paragraphs: &mut Vec<Paragraph>,
    current_paragraph: &mut Option<Paragraph>,
    in_run: &mut bool,
    in_text: &mut bool,
    current_run_text: &mut String,
    current_run_link: &mut Option<String>,
    blocks: &mut Vec<Block>,
    rels_map: &HashMap<String, String>,
    detect_title: bool,
    title: &mut Option<String>,
    shape_depth: &mut usize,
    shape_is_title: &mut bool,
) {
    let name = local_name(name);
    match name {
        b"t" => {
            *in_text = false;
        }
        b"r" => {
            if let Some(paragraph) = current_paragraph.as_mut() {
                finish_run(paragraph, rels_map, current_run_text, current_run_link);
            }
            *in_run = false;
        }
        b"p" => {
            if let Some(mut paragraph) = current_paragraph.take() {
                if *in_run {
                    finish_run(&mut paragraph, rels_map, current_run_text, current_run_link);
                    *in_run = false;
                }
                if !paragraph_is_empty(&paragraph) {
                    if detect_title && *shape_is_title && title.is_none() {
                        let text = paragraph_plain_text(&paragraph);
                        if !text.is_empty() {
                            *title = Some(text);
                        }
                    }
                    if *in_table && *in_cell && include_tables {
                        current_cell_paragraphs.push(paragraph);
                    } else {
                        blocks.push(Block::Paragraph(paragraph));
                    }
                }
            }
        }
        b"tc" if *in_table && *in_cell && include_tables => {
            if current_cell_paragraphs.is_empty() {
                current_cell_paragraphs.push(empty_paragraph());
            }
            let cell_paragraphs = std::mem::take(current_cell_paragraphs);
            current_row.push(TableCell {
                content: cell_paragraphs,
            });
            *in_cell = false;
        }
        b"tr" if *in_table && *in_row && include_tables => {
            let row = std::mem::take(current_row);
            table_rows.push(row);
            *in_row = false;
        }
        b"tbl" if *in_table && include_tables => {
            let table = build_table(std::mem::take(table_rows));
            blocks.push(Block::Table(table));
            *in_table = false;
        }
        b"sp" if *shape_depth > 0 => {
            *shape_depth -= 1;
            if *shape_depth == 0 {
                *shape_is_title = false;
            }
        }
        _ => {}
    }
}

fn finish_run(
    paragraph: &mut Paragraph,
    rels_map: &HashMap<String, String>,
    current_run_text: &mut String,
    current_run_link: &mut Option<String>,
) {
    if current_run_text.is_empty() && current_run_link.is_none() {
        return;
    }

    let display = std::mem::take(current_run_text);
    if let Some(rel_id) = current_run_link.take() {
        if let Some(target_raw) = rels_map.get(&rel_id) {
            let target = target_raw.clone();
            paragraph.inlines.push(Inline::Link(Hyperlink {
                display,
                target,
                rel_id: Some(rel_id),
            }));
        } else {
            paragraph.inlines.push(Inline::Text(display));
        }
    } else {
        paragraph.inlines.push(Inline::Text(display));
    }
}

fn build_table(rows: Vec<Vec<TableCell>>) -> Table {
    let mut max_cols = rows.iter().map(std::vec::Vec::len).max().unwrap_or(0);
    if max_cols == 0 {
        max_cols = 1;
    }

    let headers = officemd_core::ir::synthetic_col_headers(max_cols);
    let mut fixed_rows = Vec::with_capacity(rows.len().max(1));

    if rows.is_empty() {
        fixed_rows.push(vec![TableCell {
            content: vec![empty_paragraph()],
        }]);
    } else {
        for mut row in rows {
            while row.len() < max_cols {
                row.push(TableCell {
                    content: vec![empty_paragraph()],
                });
            }
            fixed_rows.push(row);
        }
    }

    Table {
        caption: None,
        headers,
        rows: fixed_rows,
        synthetic_headers: true,
    }
}

fn empty_paragraph() -> Paragraph {
    Paragraph {
        inlines: vec![Inline::Text(String::new())],
    }
}

fn paragraph_plain_text(paragraph: &Paragraph) -> String {
    let mut out = String::new();
    for inline in &paragraph.inlines {
        match inline {
            Inline::Text(t) => out.push_str(t),
            Inline::Link(link) => {
                if link.display.is_empty() {
                    out.push_str(&link.target);
                } else {
                    out.push_str(&link.display);
                }
            }
        }
    }
    out.trim().to_string()
}

fn paragraph_is_empty(paragraph: &Paragraph) -> bool {
    paragraph.inlines.iter().all(|inline| match inline {
        Inline::Text(t) => t.trim().is_empty(),
        Inline::Link(link) => link.display.trim().is_empty() && link.target.trim().is_empty(),
    })
}

fn parse_comment_authors(xml: &str) -> HashMap<String, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut authors = HashMap::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"cmAuthor"
                    && let Some(id) = attr_value_exact(e, b"id")
                {
                    let name = attr_value_exact(e, b"name").unwrap_or_default();
                    authors.insert(id, name);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {}
        }
    }

    authors
}

fn parse_comments(xml: &str, authors: &HashMap<String, String>) -> Vec<CommentNote> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut comments = Vec::new();
    let mut in_comment = false;
    let mut in_text = false;
    let mut current_text = String::new();
    let mut current_author_id: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let name = local_name(name.as_ref());
                if name == b"cm" {
                    in_comment = true;
                    current_text.clear();
                    current_author_id = attr_value_exact(e, b"authorId");
                } else if in_comment && (name == b"t" || name == b"text") {
                    in_text = true;
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let name = local_name(name.as_ref());
                if name == b"cm" {
                    current_text.clear();
                    current_author_id = attr_value_exact(e, b"authorId");
                    finalize_comment(
                        &mut comments,
                        authors,
                        &current_text,
                        current_author_id.as_deref(),
                    );
                    current_author_id = None;
                } else if in_comment && (name == b"t" || name == b"text") {
                    in_text = false;
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let name = local_name(name.as_ref());
                if name == b"cm" {
                    finalize_comment(
                        &mut comments,
                        authors,
                        &current_text,
                        current_author_id.as_deref(),
                    );
                    in_comment = false;
                    current_author_id = None;
                } else if name == b"t" || name == b"text" {
                    in_text = false;
                }
            }
            Ok(Event::Text(t)) => {
                if in_comment
                    && in_text
                    && let Ok(text) = t.unescape()
                {
                    current_text.push_str(&text);
                }
            }
            Ok(Event::CData(t)) => {
                if in_comment
                    && in_text
                    && let Ok(text) = std::str::from_utf8(t.as_ref())
                {
                    current_text.push_str(text);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {}
        }
    }

    comments
}

fn finalize_comment(
    comments: &mut Vec<CommentNote>,
    authors: &HashMap<String, String>,
    text: &str,
    author_id: Option<&str>,
) {
    let author = author_id
        .and_then(|id| authors.get(id).cloned().or_else(|| Some(id.to_string())))
        .unwrap_or_default();
    let comment_text = text.trim().to_string();

    let id = format!("c{}", comments.len() + 1);
    comments.push(CommentNote {
        id,
        author,
        text: comment_text,
    });
}

fn build_rel_map(rels: &[Relationship], part_path: &str) -> HashMap<String, String> {
    relationship_target_map(rels, part_path, None)
}

fn find_rel_target(rels: &[Relationship], rel_type: &str, part_path: &str) -> Option<String> {
    rels.iter()
        .find(|rel| rel.rel_type == rel_type)
        .map(|rel| resolve_relationship_target(part_path, rel))
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

fn attr_value_exact(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let attr_key = attr.key.as_ref();
        if attr_key == key {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
        if !key.contains(&b':') && !key.contains(&b'}') && local_name(attr_key) == key {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn attr_value_rid(e: &BytesStart<'_>) -> Option<String> {
    for attr in e.attributes().flatten() {
        let attr_key = attr.key.as_ref();
        if attr_key == b"r:id" || attr_key.ends_with(b":id") {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
        if (attr_key.contains(&b':') || attr_key.contains(&b'}')) && local_name(attr_key) == b"id" {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn decode_cdata(cdata: &quick_xml::events::BytesCData<'_>) -> Result<String, PptxError> {
    std::str::from_utf8(cdata.as_ref())
        .map(std::string::ToString::to_string)
        .map_err(|e| PptxError::Xml(e.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    fn build_pptx(parts: Vec<(&str, &str)>) -> Vec<u8> {
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

    fn slide_order_fixture_parts() -> Vec<(&'static str, &'static str)> {
        let presentation = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>
"#;

        let presentation_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>
</Relationships>
"#;

        let slide1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr><a:hlinkClick r:id="rId2"/></a:rPr>
              <a:t>Welcome</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Body</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:graphicFrame>
        <a:graphic>
          <a:graphicData>
            <a:tbl>
              <a:tr>
                <a:tc><a:txBody><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;

        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>
</Relationships>
"#;

        let slide2 = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Second</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;

        vec![
            ("ppt/presentation.xml", presentation),
            ("ppt/_rels/presentation.xml.rels", presentation_rels),
            ("ppt/slides/slide1.xml", slide1),
            ("ppt/slides/_rels/slide1.xml.rels", slide1_rels),
            ("ppt/slides/slide2.xml", slide2),
            ("ppt/notesSlides/notesSlide1.xml", slide_order_notes_xml()),
            ("ppt/commentAuthors.xml", slide_order_comment_authors_xml()),
            ("ppt/comments/comment1.xml", slide_order_comments_xml()),
        ]
    }

    fn slide_order_notes_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Note text</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>
"#
    }

    fn slide_order_comment_authors_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="Alice"/>
</p:cmAuthorLst>
"#
    }

    fn slide_order_comments_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cm authorId="0"><p:text>Needs review</p:text></p:cm>
</p:cmLst>
"#
    }

    fn assert_slide_order_document(doc: &officemd_core::ir::OoxmlDocument) {
        assert_eq!(doc.slides.len(), 2);
        assert_eq!(doc.slides[0].number, 1);
        assert_eq!(doc.slides[0].title.as_deref(), Some("Welcome"));
        assert!(
            doc.slides[0]
                .blocks
                .iter()
                .any(|b| matches!(b, Block::Table(_)))
        );
        assert_eq!(
            doc.slides[0]
                .notes
                .as_ref()
                .and_then(|notes| notes.first())
                .map(paragraph_plain_text),
            Some("Note text".to_string())
        );
        assert_eq!(doc.slides[0].comments.len(), 1);
        assert_eq!(doc.slides[0].comments[0].author, "Alice");
        assert_eq!(doc.slides[0].comments[0].text, "Needs review");
        assert_eq!(
            first_link_target(doc).as_deref(),
            Some("https://example.com")
        );
    }

    fn first_link_target(doc: &officemd_core::ir::OoxmlDocument) -> Option<String> {
        let first_para = doc.slides[0]
            .blocks
            .iter()
            .find_map(|block| match block {
                Block::Paragraph(p) => Some(p),
                _ => None,
            })
            .unwrap();

        first_para.inlines.iter().find_map(|inline| match inline {
            Inline::Link(link) => Some(link.target.clone()),
            Inline::Text(_) => None,
        })
    }

    #[test]
    fn extracts_slide_order_and_content() {
        let content = build_pptx(slide_order_fixture_parts());
        let doc = extract_ir(&content).unwrap();
        let json = extract_ir_json(&content).unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed["kind"], "Pptx");
        assert_slide_order_document(&doc);
    }

    #[test]
    fn filters_slides_during_extraction() {
        let presentation = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>
"#;

        let presentation_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>
"#;

        let slide1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp><p:txBody><a:p><a:r><a:t>First</a:t></a:r></a:p></p:txBody></p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;
        let slide2 = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp><p:txBody><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;

        let content = build_pptx(vec![
            ("ppt/presentation.xml", presentation),
            ("ppt/_rels/presentation.xml.rels", presentation_rels),
            ("ppt/slides/slide1.xml", slide1),
            ("ppt/slides/slide2.xml", slide2),
        ]);

        let mut selected = HashSet::new();
        selected.insert(2);

        let doc = extract_ir_with_options(
            &content,
            PptxExtractOptions {
                slide_numbers: Some(selected),
            },
        )
        .expect("filtered extraction");

        assert_eq!(doc.slides.len(), 1);
        assert_eq!(doc.slides[0].number, 2);
        let body_text = doc.slides[0]
            .blocks
            .iter()
            .find_map(|block| match block {
                Block::Paragraph(p) => Some(paragraph_plain_text(p)),
                _ => None,
            })
            .unwrap_or_default();
        assert_eq!(body_text, "Second");
    }

    #[test]
    fn slide_rid_collection_stops_after_selection_max() {
        let presentation = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
  <broken
</p:presentation>
"#;

        let mut selected = HashSet::new();
        selected.insert(1);
        let selection = SlideSelection::new(selected);
        let entries =
            collect_slide_rids(presentation, Some(&selection)).expect("collect selected slides");

        assert_eq!(entries, vec![(1, "rId1".to_string())]);
    }

    #[test]
    fn errors_on_malformed_presentation_slide_id_xml() {
        let malformed_presentation = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1">
  </p:sldIdLst>
</p:presentation>
"#;

        let content = build_pptx(vec![("ppt/presentation.xml", malformed_presentation)]);
        let err = extract_ir(&content).expect_err("malformed presentation should error");
        assert!(matches!(err, PptxError::Xml(_)));
    }
}
