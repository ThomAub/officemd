//! Streaming DOCX extraction into the shared OOXML IR.

use std::collections::{HashMap, HashSet};

use officemd_core::ir::{
    Block, CommentNote, DocSection, DocumentKind, DocumentProperties, Hyperlink, Inline,
    OoxmlDocument, Paragraph, Table, TableCell,
};
use officemd_core::opc::{OpcPackage, load_relationships_for_part, resolve_relationship_target};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::error::DocxError;

const HYPERLINK_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/// Extract a DOCX into the shared IR.
///
/// # Errors
///
/// Returns [`DocxError`] if the content is not a valid ZIP archive, required
/// XML parts cannot be parsed, or relationship resolution fails.
pub fn extract_ir(content: &[u8]) -> Result<OoxmlDocument, DocxError> {
    let mut package = OpcPackage::from_bytes(content).map_err(DocxError::from)?;
    let comments = extract_comments(&mut package)?;

    let mut sections = Vec::new();

    if let Some(xml) = read_part_to_string(&mut package, "word/document.xml")? {
        let rels = load_rels_map(&mut package, "word/document.xml")?;
        let part = extract_part(&xml, &rels, &comments)?;
        sections.push(DocSection {
            name: "body".to_string(),
            blocks: part.blocks,
            comments: part.comments,
        });
    }

    for header in list_parts(&mut package, "word/header", ".xml") {
        if let Some(xml) = read_part_to_string(&mut package, &header)? {
            let rels = load_rels_map(&mut package, &header)?;
            let part = extract_part(&xml, &rels, &comments)?;
            sections.push(DocSection {
                name: part_name(&header),
                blocks: part.blocks,
                comments: part.comments,
            });
        }
    }

    for footer in list_parts(&mut package, "word/footer", ".xml") {
        if let Some(xml) = read_part_to_string(&mut package, &footer)? {
            let rels = load_rels_map(&mut package, &footer)?;
            let part = extract_part(&xml, &rels, &comments)?;
            sections.push(DocSection {
                name: part_name(&footer),
                blocks: part.blocks,
                comments: part.comments,
            });
        }
    }

    if let Some(xml) = read_part_to_string(&mut package, "word/footnotes.xml")? {
        let rels = load_rels_map(&mut package, "word/footnotes.xml")?;
        let part = extract_part(&xml, &rels, &comments)?;
        sections.push(DocSection {
            name: "footnotes".to_string(),
            blocks: part.blocks,
            comments: part.comments,
        });
    }

    if let Some(xml) = read_part_to_string(&mut package, "word/endnotes.xml")? {
        let rels = load_rels_map(&mut package, "word/endnotes.xml")?;
        let part = extract_part(&xml, &rels, &comments)?;
        sections.push(DocSection {
            name: "endnotes".to_string(),
            blocks: part.blocks,
            comments: part.comments,
        });
    }

    let properties = extract_properties(&mut package)?;

    Ok(OoxmlDocument {
        kind: DocumentKind::Docx,
        properties,
        sections,
        ..Default::default()
    })
}

/// Extract minimal IR as JSON string.
///
/// # Errors
///
/// Returns [`DocxError`] if extraction fails or the IR cannot be serialized
/// to JSON.
pub fn extract_ir_json(content: &[u8]) -> Result<String, DocxError> {
    let doc = extract_ir(content)?;
    serde_json::to_string(&doc).map_err(|e| DocxError::Json(e.to_string()))
}

struct PartOutput {
    blocks: Vec<Block>,
    comments: Vec<CommentNote>,
}

fn extract_part(
    xml: &str,
    rels: &HashMap<String, String>,
    comments: &HashMap<String, CommentNote>,
) -> Result<PartOutput, DocxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut state = PartState::new(rels, comments);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => state.handle_start(e),
            Ok(Event::Empty(ref e)) => state.handle_empty(e),
            Ok(Event::End(ref e)) => state.handle_end(e),
            Ok(Event::Text(ref t)) => state.handle_text(t)?,
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(DocxError::Xml(e.to_string())),
        }
    }

    state.finish_all();

    Ok(PartOutput {
        blocks: state.blocks,
        comments: state.section_comments,
    })
}

#[derive(Default)]
struct ParagraphBuilder {
    inlines: Vec<Inline>,
}

#[derive(Default)]
struct TableCellBuilder {
    paragraphs: Vec<Paragraph>,
}

impl TableCellBuilder {
    fn push_paragraph(&mut self, paragraph: Paragraph) {
        self.paragraphs.push(paragraph);
    }

    fn into_cell(self) -> TableCell {
        let mut paragraphs = self.paragraphs;
        if paragraphs.is_empty() {
            paragraphs.push(Paragraph {
                inlines: vec![Inline::Text(String::new())],
            });
        }
        TableCell {
            content: paragraphs,
        }
    }
}

#[derive(Default)]
struct TableBuilder {
    rows: Vec<Vec<TableCell>>,
    current_row: Option<Vec<TableCell>>,
    current_cell: Option<TableCellBuilder>,
}

impl TableBuilder {
    fn start_row(&mut self) {
        self.finish_cell();
        self.current_row = Some(Vec::new());
    }

    fn finish_row(&mut self) {
        self.finish_cell();
        if let Some(row) = self.current_row.take() {
            self.rows.push(row);
        }
    }

    fn start_cell(&mut self) {
        if self.current_row.is_none() {
            self.current_row = Some(Vec::new());
        }
        self.current_cell = Some(TableCellBuilder::default());
    }

    fn finish_cell(&mut self) {
        if let Some(cell) = self.current_cell.take()
            && let Some(row) = &mut self.current_row
        {
            row.push(cell.into_cell());
        }
    }

    fn push_paragraph(&mut self, paragraph: Paragraph) {
        if self.current_cell.is_none() {
            self.start_cell();
        }
        if let Some(cell) = &mut self.current_cell {
            cell.push_paragraph(paragraph);
        }
    }
}

struct LinkBuilder {
    target: String,
    rel_id: Option<String>,
    display: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum LinkEnd {
    Hyperlink,
    FieldSimple,
    FieldComplex,
}

struct LinkContext {
    builder: LinkBuilder,
    end_on: LinkEnd,
}

struct PartState<'a> {
    rels: &'a HashMap<String, String>,
    comment_map: &'a HashMap<String, CommentNote>,
    blocks: Vec<Block>,
    section_comments: Vec<CommentNote>,
    seen_comments: HashSet<String>,
    current_paragraph: Option<ParagraphBuilder>,
    current_table: Option<TableBuilder>,
    current_link: Option<LinkContext>,
    collect_instr_text: bool,
    in_text_run: bool,
    instr_buffer: String,
    pending_field_target: Option<String>,
}

impl<'a> PartState<'a> {
    fn new(
        rels: &'a HashMap<String, String>,
        comment_map: &'a HashMap<String, CommentNote>,
    ) -> Self {
        Self {
            rels,
            comment_map,
            blocks: Vec::new(),
            section_comments: Vec::new(),
            seen_comments: HashSet::new(),
            current_paragraph: None,
            current_table: None,
            current_link: None,
            collect_instr_text: false,
            in_text_run: false,
            instr_buffer: String::new(),
            pending_field_target: None,
        }
    }

    fn handle_start(&mut self, e: &BytesStart<'_>) {
        match local_name(e.name().as_ref()) {
            "p" => self.start_paragraph(),
            "tbl" => self.start_table(),
            "tr" => self.start_row(),
            "tc" => self.start_cell(),
            "t" => self.in_text_run = true,
            "hyperlink" => self.start_hyperlink(e),
            "fldSimple" => self.start_fld_simple(e),
            "instrText" => self.start_instr_text(),
            "fldChar" => self.handle_fld_char(e),
            "commentReference" => self.handle_comment_marker(e, true),
            "commentRangeStart" => self.handle_comment_marker(e, false),
            "tab" => self.push_text("\t"),
            "br" | "cr" => self.push_text("\n"),
            _ => {}
        }
    }

    fn handle_empty(&mut self, e: &BytesStart<'_>) {
        match local_name(e.name().as_ref()) {
            "p" => {
                self.start_paragraph();
                self.finish_paragraph();
            }
            "tbl" => {
                self.start_table();
                self.finish_table();
            }
            "tr" => {
                self.start_row();
                self.finish_row();
            }
            "tc" => {
                self.start_cell();
                self.finish_cell();
            }
            "hyperlink" => {
                self.start_hyperlink(e);
                self.end_link_on(LinkEnd::Hyperlink);
            }
            "fldSimple" => {
                self.start_fld_simple(e);
                self.end_link_on(LinkEnd::FieldSimple);
            }
            "commentReference" => self.handle_comment_marker(e, true),
            "commentRangeStart" => self.handle_comment_marker(e, false),
            "tab" => self.push_text("\t"),
            "br" | "cr" => self.push_text("\n"),
            _ => {}
        }
    }

    fn handle_end(&mut self, e: &quick_xml::events::BytesEnd<'_>) {
        match local_name(e.name().as_ref()) {
            "p" => self.finish_paragraph(),
            "tbl" => self.finish_table(),
            "tr" => self.finish_row(),
            "tc" => self.finish_cell(),
            "t" => self.in_text_run = false,
            "hyperlink" => self.end_link_on(LinkEnd::Hyperlink),
            "fldSimple" => self.end_link_on(LinkEnd::FieldSimple),
            "instrText" => self.finish_instr_text(),
            _ => {}
        }
    }

    fn handle_text(&mut self, t: &quick_xml::events::BytesText<'_>) -> Result<(), DocxError> {
        if !self.collect_instr_text && !self.in_text_run {
            return Ok(());
        }

        let text = t
            .unescape()
            .map_err(|e| DocxError::Xml(e.to_string()))?
            .to_string();
        if self.collect_instr_text {
            self.instr_buffer.push_str(&text);
        } else if self.in_text_run {
            self.push_text(&text);
        }
        Ok(())
    }

    fn finish_all(&mut self) {
        self.finish_paragraph();
        self.finish_table();
    }

    fn start_paragraph(&mut self) {
        self.finish_paragraph();
        self.current_paragraph = Some(ParagraphBuilder::default());
    }

    fn finish_paragraph(&mut self) {
        self.finish_link();
        if let Some(builder) = self.current_paragraph.take() {
            let paragraph = Paragraph {
                inlines: builder.inlines,
            };
            if let Some(table) = &mut self.current_table {
                table.push_paragraph(paragraph);
            } else if !paragraph.inlines.is_empty() {
                self.blocks.push(Block::Paragraph(paragraph));
            }
        }
    }

    fn start_table(&mut self) {
        self.finish_paragraph();
        self.current_table = Some(TableBuilder::default());
    }

    fn finish_table(&mut self) {
        if let Some(mut table) = self.current_table.take() {
            table.finish_cell();
            table.finish_row();
            let mut rows = table.rows;
            let mut max_cols = rows.iter().map(std::vec::Vec::len).max().unwrap_or(0);
            if max_cols == 0 {
                max_cols = 1;
                rows.push(vec![TableCell {
                    content: vec![Paragraph {
                        inlines: vec![Inline::Text(String::new())],
                    }],
                }]);
            }
            for row in &mut rows {
                while row.len() < max_cols {
                    row.push(TableCell {
                        content: vec![Paragraph {
                            inlines: vec![Inline::Text(String::new())],
                        }],
                    });
                }
            }
            let headers = officemd_core::ir::synthetic_col_headers(max_cols);
            self.blocks.push(Block::Table(Table {
                caption: None,
                headers,
                rows,
                synthetic_headers: true,
            }));
        }
    }

    fn start_row(&mut self) {
        if let Some(table) = &mut self.current_table {
            table.start_row();
        }
    }

    fn finish_row(&mut self) {
        if let Some(table) = &mut self.current_table {
            table.finish_row();
        }
    }

    fn start_cell(&mut self) {
        if let Some(table) = &mut self.current_table {
            table.start_cell();
        }
    }

    fn finish_cell(&mut self) {
        if let Some(table) = &mut self.current_table {
            table.finish_cell();
        }
    }

    fn start_hyperlink(&mut self, e: &BytesStart<'_>) {
        let rel_id = attr_value(e, "id");
        let target = rel_id
            .as_ref()
            .and_then(|id| self.rels.get(id).cloned())
            .unwrap_or_default();
        self.start_link(target, rel_id, LinkEnd::Hyperlink);
    }

    fn start_fld_simple(&mut self, e: &BytesStart<'_>) {
        if let Some(instr) = attr_value(e, "instr")
            && let Some(target) = parse_hyperlink_instr(&instr)
        {
            self.start_link(target, None, LinkEnd::FieldSimple);
        }
    }

    fn start_instr_text(&mut self) {
        self.collect_instr_text = true;
        self.instr_buffer.clear();
    }

    fn finish_instr_text(&mut self) {
        self.collect_instr_text = false;
        if let Some(target) = parse_hyperlink_instr(&self.instr_buffer) {
            self.pending_field_target = Some(target);
        }
        self.instr_buffer.clear();
    }

    fn handle_fld_char(&mut self, e: &BytesStart<'_>) {
        if let Some(kind) = attr_value(e, "fldCharType") {
            match kind.as_str() {
                "separate" => {
                    if let Some(target) = self.pending_field_target.take() {
                        self.start_link(target, None, LinkEnd::FieldComplex);
                    }
                }
                "end"
                    if self.current_link.as_ref().map(|c| c.end_on)
                        == Some(LinkEnd::FieldComplex) =>
                {
                    self.end_link_on(LinkEnd::FieldComplex);
                }
                _ => {}
            }
        }
    }

    fn handle_comment_marker(&mut self, e: &BytesStart<'_>, emit_anchor: bool) {
        if let Some(raw_id) = attr_value(e, "id")
            && let Some(note) = self.comment_map.get(&raw_id)
        {
            if self.seen_comments.insert(note.id.clone()) {
                self.section_comments.push(note.clone());
            }
            if emit_anchor {
                self.finish_link();
                self.push_text_raw(&format!("[^{}]", note.id));
            }
        }
    }

    fn start_link(&mut self, target: String, rel_id: Option<String>, end_on: LinkEnd) {
        self.finish_link();
        self.current_link = Some(LinkContext {
            builder: LinkBuilder {
                target,
                rel_id,
                display: String::new(),
            },
            end_on,
        });
    }

    fn finish_link(&mut self) {
        if let Some(link) = self.current_link.take()
            && let Some(paragraph) = &mut self.current_paragraph
        {
            let target = link.builder.target;
            if target.is_empty() {
                if !link.builder.display.is_empty() {
                    paragraph.inlines.push(Inline::Text(link.builder.display));
                }
            } else {
                paragraph.inlines.push(Inline::Link(Hyperlink {
                    display: link.builder.display,
                    target,
                    rel_id: link.builder.rel_id,
                }));
            }
        }
    }

    fn end_link_on(&mut self, end_on: LinkEnd) {
        if self.current_link.as_ref().map(|c| c.end_on) == Some(end_on) {
            self.finish_link();
        }
    }

    fn push_text(&mut self, text: &str) {
        if text.is_empty() {
            return;
        }
        if self.current_paragraph.is_none() {
            return;
        }
        self.push_text_raw(text);
    }

    fn push_text_raw(&mut self, text: &str) {
        if text.is_empty() {
            return;
        }
        if let Some(link) = &mut self.current_link {
            link.builder.display.push_str(text);
        } else if let Some(paragraph) = &mut self.current_paragraph {
            if let Some(Inline::Text(last)) = paragraph.inlines.last_mut() {
                last.push_str(text);
            } else {
                paragraph.inlines.push(Inline::Text(text.to_string()));
            }
        }
    }
}

fn local_name(name: &[u8]) -> &str {
    let s = std::str::from_utf8(name).unwrap_or("");
    if let Some(idx) = s.rfind(':') {
        &s[idx + 1..]
    } else if let Some(idx) = s.rfind('}') {
        &s[idx + 1..]
    } else {
        s
    }
}

fn attr_value(e: &BytesStart<'_>, key: &str) -> Option<String> {
    for attr in e.attributes().flatten() {
        if local_name(attr.key.as_ref()) == key {
            if let Ok(value) = attr.unescape_value() {
                return Some(value.to_string());
            }
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn parse_hyperlink_instr(instr: &str) -> Option<String> {
    let upper = instr.to_ascii_uppercase();
    let idx = upper.find("HYPERLINK")?;
    let tail = instr[idx + "HYPERLINK".len()..].trim();
    if let Some(start) = tail.find('"') {
        let rest = &tail[start + 1..];
        if let Some(end) = rest.find('"') {
            return Some(rest[..end].to_string());
        }
    }
    let token = tail.split_whitespace().next()?;
    if token.starts_with("http") || token.starts_with("mailto:") {
        return Some(token.to_string());
    }
    None
}

fn part_name(path: &str) -> String {
    path.rsplit('/')
        .next()
        .unwrap_or(path)
        .trim_end_matches(".xml")
        .to_string()
}

fn list_parts(package: &mut OpcPackage<'_>, prefix: &str, suffix: &str) -> Vec<String> {
    package.list_parts(prefix, suffix)
}

fn read_part_to_string(
    package: &mut OpcPackage<'_>,
    path: &str,
) -> Result<Option<String>, DocxError> {
    package.read_part_string(path).map_err(DocxError::from)
}

fn load_rels_map(
    package: &mut OpcPackage<'_>,
    part: &str,
) -> Result<HashMap<String, String>, DocxError> {
    let rels = load_relationships_for_part(package, part).map_err(DocxError::from)?;
    let mut map = HashMap::new();
    for rel in rels {
        let is_external = rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.eq_ignore_ascii_case("external"));
        if rel.rel_type == HYPERLINK_REL_TYPE && is_external {
            let target = resolve_relationship_target(part, &rel);
            map.insert(rel.id, target);
        }
    }
    Ok(map)
}

fn extract_comments(
    package: &mut OpcPackage<'_>,
) -> Result<HashMap<String, CommentNote>, DocxError> {
    let Some(xml) = read_part_to_string(package, "word/comments.xml")? else {
        return Ok(HashMap::new());
    };

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(false);

    let mut comments = HashMap::new();
    let mut current_id: Option<String> = None;
    let mut current_author = String::new();
    let mut current_text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match local_name(e.name().as_ref()) {
                "comment" => {
                    current_id = attr_value(e, "id");
                    current_author = attr_value(e, "author").unwrap_or_default();
                    current_text.clear();
                }
                "tab" if current_id.is_some() => {
                    current_text.push('\t');
                }
                "br" | "cr" if current_id.is_some() => {
                    current_text.push('\n');
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match local_name(e.name().as_ref()) {
                "comment" => {
                    current_id = attr_value(e, "id");
                    current_author = attr_value(e, "author").unwrap_or_default();
                    current_text.clear();
                    if let Some(raw_id) = current_id.take() {
                        let note = build_comment_note(raw_id, &current_author, &current_text);
                        comments.insert(note.0, note.1);
                    }
                }
                "tab" if current_id.is_some() => {
                    current_text.push('\t');
                }
                "br" | "cr" if current_id.is_some() => {
                    current_text.push('\n');
                }
                _ => {}
            },
            Ok(Event::Text(ref t)) => {
                if current_id.is_some() {
                    let text = t
                        .unescape()
                        .map_err(|e| DocxError::Xml(e.to_string()))?
                        .to_string();
                    current_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => match local_name(e.name().as_ref()) {
                "comment" => {
                    if let Some(raw_id) = current_id.take() {
                        let note = build_comment_note(raw_id, &current_author, &current_text);
                        comments.insert(note.0, note.1);
                    }
                }
                "p" if current_id.is_some() && !current_text.ends_with('\n') => {
                    current_text.push('\n');
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(DocxError::Xml(e.to_string())),
        }
    }

    Ok(comments)
}

fn build_comment_note(raw_id: String, author: &str, text: &str) -> (String, CommentNote) {
    let author = author.trim().to_string();
    let text = text.trim().to_string();
    let note_id = format!("c{raw_id}");
    let note = CommentNote {
        id: note_id.clone(),
        author,
        text,
    };
    (raw_id, note)
}

fn extract_properties(
    package: &mut OpcPackage<'_>,
) -> Result<Option<DocumentProperties>, DocxError> {
    let props = DocumentProperties {
        core: extract_props_map(package, "docProps/core.xml")?,
        app: extract_props_map(package, "docProps/app.xml")?,
        custom: extract_custom_props_map(package, "docProps/custom.xml")?,
    };

    if props.core.is_empty() && props.app.is_empty() && props.custom.is_empty() {
        Ok(None)
    } else {
        Ok(Some(props))
    }
}

fn extract_props_map(
    package: &mut OpcPackage<'_>,
    path: &str,
) -> Result<HashMap<String, String>, DocxError> {
    let Some(xml) = read_part_to_string(package, path)? else {
        return Ok(HashMap::new());
    };

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut current_tag: Option<String> = None;
    let mut map = HashMap::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                current_tag = Some(local_name(e.name().as_ref()).to_string());
            }
            Ok(Event::Text(ref t)) => {
                if let Some(tag) = &current_tag {
                    let val = t
                        .unescape()
                        .map_err(|e| DocxError::Xml(e.to_string()))?
                        .to_string();
                    map.insert(tag.clone(), val);
                }
            }
            Ok(Event::End(_)) => {
                current_tag = None;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(DocxError::Xml(e.to_string())),
        }
    }

    Ok(map)
}

fn extract_custom_props_map(
    package: &mut OpcPackage<'_>,
    path: &str,
) -> Result<HashMap<String, String>, DocxError> {
    let Some(xml) = read_part_to_string(package, path)? else {
        return Ok(HashMap::new());
    };

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut map = HashMap::new();
    let mut current_name: Option<String> = None;
    let mut current_value = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if local_name(e.name().as_ref()) == "property" {
                    current_name = attr_value(e, "name");
                    current_value.clear();
                }
            }
            Ok(Event::Text(ref t)) => {
                if current_name.is_some() {
                    let text = t
                        .unescape()
                        .map_err(|e| DocxError::Xml(e.to_string()))?
                        .to_string();
                    current_value.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                if local_name(e.name().as_ref()) == "property"
                    && let Some(name) = current_name.take()
                {
                    let value = current_value.trim().to_string();
                    map.insert(name, value);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(DocxError::Xml(e.to_string())),
        }
    }

    Ok(map)
}
