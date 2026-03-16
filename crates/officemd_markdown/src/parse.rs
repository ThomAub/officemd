//! Markdown parser that reconstructs OOXML IR from rendered markdown.
//!
//! Guarantees render-parse-render stability: `render(parse(render(ir))) == render(ir)`.

use std::collections::HashMap;
use std::sync::LazyLock;

use regex::Regex;

use officemd_core::ir::{
    self, Block, CommentNote, DocSection, DocumentKind, DocumentProperties, FormulaNote, Hyperlink,
    Inline, OoxmlDocument, Paragraph, PdfDocument, PdfPage, Sheet, Slide, Table, TableCell,
};

use crate::{MarkdownProfile, RenderOptions};

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Options for parsing markdown back into IR.
///
/// With frontmatter present, most options are auto-detected.
/// Manual overrides for when frontmatter is absent (e.g. hand-written markdown).
#[derive(Debug, Clone, Default)]
pub struct ParseOptions {
    pub markdown_profile: Option<MarkdownProfile>,
    pub assume_kind: Option<DocumentKind>,
}

/// Errors returned by the parser.
#[derive(Debug, thiserror::Error)]
pub enum ParseError {
    #[error("cannot determine document kind from markdown content")]
    UnknownKind,
    #[error("parse error: {0}")]
    Format(String),
}

/// Parse officemd-rendered markdown back into an `OoxmlDocument`.
///
/// Reads options from the `<!-- officemd: ... -->` frontmatter comment.
///
/// # Errors
///
/// Returns `ParseError` if the document kind cannot be determined.
pub fn parse_document(input: &str) -> Result<OoxmlDocument, ParseError> {
    parse_document_with_options(input, ParseOptions::default())
}

/// Parse with explicit option overrides.
///
/// # Errors
///
/// Returns `ParseError` if the document kind cannot be determined.
pub fn parse_document_with_options(
    input: &str,
    options: ParseOptions,
) -> Result<OoxmlDocument, ParseError> {
    let fm = parse_frontmatter(input);

    let render_opts = fm.as_ref().map(|f| f.render_options);

    let kind = options
        .assume_kind
        .or(fm.as_ref().map(|f| f.kind))
        .or_else(|| detect_kind(input))
        .ok_or(ParseError::UnknownKind)?;

    let profile = options
        .markdown_profile
        .or(render_opts.map(|o| o.markdown_profile))
        .unwrap_or(MarkdownProfile::LlmCompact);

    let use_first_row_as_header = render_opts.is_none_or(|o| o.use_first_row_as_header);

    // Strip frontmatter line from input
    let body = if fm.is_some() {
        strip_frontmatter(input)
    } else {
        input
    };

    let mut doc = match kind {
        DocumentKind::Pdf => parse_pdf(body),
        DocumentKind::Xlsx => parse_xlsx(body, profile, use_first_row_as_header),
        DocumentKind::Docx => parse_docx(body, profile, use_first_row_as_header),
        DocumentKind::Pptx => parse_pptx(body, profile, use_first_row_as_header),
    };
    doc.kind = kind;

    // Parse properties if present
    doc.properties = parse_properties_from_body(body, profile);

    Ok(doc)
}

// ---------------------------------------------------------------------------
// Frontmatter parsing
// ---------------------------------------------------------------------------

struct FrontmatterOptions {
    kind: DocumentKind,
    render_options: RenderOptions,
}

static FRONTMATTER_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^<!--\s*officemd:\s*(.+?)\s*-->").expect("frontmatter regex"));

fn parse_frontmatter(input: &str) -> Option<FrontmatterOptions> {
    let first_line = input.lines().next()?;
    let caps = FRONTMATTER_RE.captures(first_line)?;
    let attrs = caps.get(1)?.as_str();

    let map: HashMap<&str, &str> = attrs
        .split_whitespace()
        .filter_map(|pair| pair.split_once('='))
        .collect();

    let kind = DocumentKind::from_str_opt(map.get("kind")?)?;

    let profile = match map.get("profile").copied().unwrap_or("compact") {
        "human" => MarkdownProfile::Human,
        _ => MarkdownProfile::LlmCompact,
    };

    let first_row_as_header = map.get("first_row_as_header").is_none_or(|v| *v == "true");
    let formulas = map.get("formulas").is_none_or(|v| *v == "true");
    let headers_footers = map.get("headers_footers").is_none_or(|v| *v == "true");
    let properties = map.get("properties").is_some_and(|v| *v == "true");

    Some(FrontmatterOptions {
        kind,
        render_options: RenderOptions {
            include_document_properties: properties,
            use_first_row_as_header: first_row_as_header,
            include_headers_footers: headers_footers,
            include_formulas: formulas,
            markdown_profile: profile,
        },
    })
}

fn strip_frontmatter(input: &str) -> &str {
    // Skip the frontmatter line and the blank line after it
    if let Some(rest) = input.strip_prefix("<!--")
        && let Some(after_comment) = rest.find("-->")
    {
        let mut s = &rest[after_comment + 3..];
        // Strip up to two newlines (the line ending + the blank line)
        if let Some(stripped) = s.strip_prefix('\n') {
            s = stripped;
        }
        if let Some(stripped) = s.strip_prefix('\n') {
            s = stripped;
        }
        return s;
    }
    input
}

// ---------------------------------------------------------------------------
// Kind detection (fallback when no frontmatter)
// ---------------------------------------------------------------------------

fn detect_kind(input: &str) -> Option<DocumentKind> {
    for line in input.lines() {
        if let Some(rest) = line.strip_prefix("## ") {
            if rest.starts_with("Sheet:") {
                return Some(DocumentKind::Xlsx);
            }
            if rest.starts_with("Section:") {
                return Some(DocumentKind::Docx);
            }
            if rest.starts_with("Slide ") {
                return Some(DocumentKind::Pptx);
            }
            if rest.starts_with("Page:") {
                return Some(DocumentKind::Pdf);
            }
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Properties parsing
// ---------------------------------------------------------------------------

fn parse_properties_from_body(body: &str, profile: MarkdownProfile) -> Option<DocumentProperties> {
    match profile {
        MarkdownProfile::Human => parse_properties_human(body),
        MarkdownProfile::LlmCompact => parse_properties_compact(body),
    }
}

fn parse_properties_human(body: &str) -> Option<DocumentProperties> {
    // Look for "### Document Properties" section
    let mut in_props = false;
    let mut core = HashMap::new();

    for line in body.lines() {
        if line == "### Document Properties" {
            in_props = true;
            continue;
        }
        if in_props {
            if line.starts_with("## ") || line == "---" {
                break;
            }
            if let Some(rest) = line.strip_prefix("- ")
                && let Some((k, v)) = rest.split_once(": ")
            {
                core.insert(k.to_string(), unescape_pipes(v));
            }
        }
    }

    if core.is_empty() {
        None
    } else {
        Some(DocumentProperties {
            core,
            app: HashMap::new(),
            custom: HashMap::new(),
        })
    }
}

fn parse_properties_compact(body: &str) -> Option<DocumentProperties> {
    // Look for "properties: key=val; key=val" line
    for line in body.lines() {
        if let Some(rest) = line.strip_prefix("properties: ") {
            let mut core = HashMap::new();
            for pair in rest.split("; ") {
                if let Some((k, v)) = pair.split_once('=') {
                    core.insert(k.to_string(), unescape_pipes(v));
                }
            }
            if !core.is_empty() {
                return Some(DocumentProperties {
                    core,
                    app: HashMap::new(),
                    custom: HashMap::new(),
                });
            }
        }
        // Stop looking after the first ## header
        if line.starts_with("## ") {
            break;
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Section splitting
// ---------------------------------------------------------------------------

/// Split markdown at `## ` boundaries, returning (header_line, body_text) pairs.
fn split_at_h2(body: &str) -> Vec<(&str, &str)> {
    split_at_line(body, |line| line.starts_with("## "))
}

// ---------------------------------------------------------------------------
// Inline parsing
// ---------------------------------------------------------------------------

static LINK_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\[([^\]]*)\]\(([^)]*)\)").expect("link regex"));

static AUTOLINK_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"<(https?://[^>]+)>").expect("autolink regex"));

fn parse_inlines(text: &str) -> Vec<Inline> {
    if text.is_empty() {
        return vec![Inline::Text(String::new())];
    }

    enum Match<'a> {
        Link(regex::Captures<'a>),
        Auto(regex::Captures<'a>),
    }

    let mut inlines = Vec::new();
    let mut remaining = text;

    while !remaining.is_empty() {
        let link_caps = LINK_RE.captures(remaining);
        let auto_caps = AUTOLINK_RE.captures(remaining);

        let link_start = link_caps.as_ref().map(|c| c.get(0).unwrap().start());
        let auto_start = auto_caps.as_ref().map(|c| c.get(0).unwrap().start());

        let earliest = match (link_caps, auto_caps) {
            (Some(l), Some(a)) => {
                if link_start.unwrap() <= auto_start.unwrap() {
                    Some(Match::Link(l))
                } else {
                    Some(Match::Auto(a))
                }
            }
            (Some(l), None) => Some(Match::Link(l)),
            (None, Some(a)) => Some(Match::Auto(a)),
            (None, None) => None,
        };

        match earliest {
            Some(m) => {
                let (full, inline) = match m {
                    Match::Link(caps) => {
                        let full = caps.get(0).unwrap();
                        let display = caps.get(1).unwrap().as_str().to_string();
                        let target = caps.get(2).unwrap().as_str().to_string();
                        (
                            full,
                            Inline::Link(Hyperlink {
                                display,
                                target,
                                rel_id: None,
                            }),
                        )
                    }
                    Match::Auto(caps) => {
                        let full = caps.get(0).unwrap();
                        let target = caps.get(1).unwrap().as_str().to_string();
                        (
                            full,
                            Inline::Link(Hyperlink {
                                display: String::new(),
                                target,
                                rel_id: None,
                            }),
                        )
                    }
                };
                if full.start() > 0 {
                    inlines.push(Inline::Text(remaining[..full.start()].to_string()));
                }
                inlines.push(inline);
                remaining = &remaining[full.end()..];
            }
            None => {
                inlines.push(Inline::Text(remaining.to_string()));
                break;
            }
        }
    }

    if inlines.is_empty() {
        inlines.push(Inline::Text(String::new()));
    }
    inlines
}

// ---------------------------------------------------------------------------
// Table parsing
// ---------------------------------------------------------------------------

fn parse_table_block(lines: &[&str], use_first_row_as_header: bool) -> (Table, usize) {
    // Find header row (first pipe-delimited row)
    let mut consumed = 0;

    // Optional ### caption above the table
    let mut caption: Option<String> = None;
    if let Some(line) = lines.first()
        && let Some(cap) = line.strip_prefix("### ")
    {
        caption = Some(cap.to_string());
        consumed += 1;
    }

    let table_lines = &lines[consumed..];
    let mut header_line: Option<&str> = None;
    let mut data_lines: Vec<&str> = Vec::new();
    let mut saw_separator = false;

    for line in table_lines {
        if line.starts_with('|') {
            if header_line.is_none() {
                header_line = Some(line);
                consumed += 1;
            } else if !saw_separator && is_table_separator(line) {
                saw_separator = true;
                consumed += 1;
            } else {
                data_lines.push(line);
                consumed += 1;
            }
        } else {
            break;
        }
    }

    let header_cells = header_line.map(split_table_row).unwrap_or_default();

    let (headers, rows, synthetic_headers) = if use_first_row_as_header {
        // The markdown header row was originally the first data row.
        // Reconstruct: put header cells back as rows[0], generate synthetic headers.
        let num_cols = header_cells.len().max(1);
        let headers = ir::synthetic_col_headers(num_cols);

        let mut all_rows = Vec::new();
        // header cells become the first data row
        let first_row: Vec<TableCell> = header_cells.iter().map(|s| parse_table_cell(s)).collect();
        all_rows.push(first_row);

        for line in &data_lines {
            let cells: Vec<TableCell> = split_table_row(line)
                .iter()
                .map(|s| parse_table_cell(s))
                .collect();
            all_rows.push(cells);
        }

        (headers, all_rows, true)
    } else {
        // The markdown header row maps directly to Table.headers.
        let headers: Vec<String> = header_cells.iter().map(|s| unescape_pipes(s)).collect();

        let rows: Vec<Vec<TableCell>> = data_lines
            .iter()
            .map(|line| {
                split_table_row(line)
                    .iter()
                    .map(|s| parse_table_cell(s))
                    .collect()
            })
            .collect();

        (headers, rows, false)
    };

    let table = Table {
        caption,
        headers,
        rows,
        synthetic_headers,
    };

    (table, consumed)
}

fn is_table_separator(line: &str) -> bool {
    let trimmed = line.trim();
    if !trimmed.starts_with('|') || !trimmed.ends_with('|') {
        return false;
    }
    trimmed.trim_matches('|').split('|').all(|cell| {
        let c = cell.trim();
        c.chars().all(|ch| ch == '-' || ch == ' ' || ch == ':') && c.contains('-')
    })
}

fn split_table_row(line: &str) -> Vec<String> {
    let trimmed = line.trim();
    let inner = trimmed.strip_prefix('|').unwrap_or(trimmed);
    let inner = inner.strip_suffix('|').unwrap_or(inner);

    // Fast path: no escaped pipes, use simple split
    if !inner.contains("\\|") {
        return inner.split('|').map(|s| s.trim().to_string()).collect();
    }

    // Slow path: handle escaped pipes character by character
    let mut cells = Vec::new();
    let mut current = String::new();
    let mut chars = inner.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch == '\\' {
            if let Some(&next) = chars.peek()
                && next == '|'
            {
                current.push('|');
                chars.next();
                continue;
            }
            current.push('\\');
        } else if ch == '|' {
            cells.push(current.trim().to_string());
            current = String::new();
        } else {
            current.push(ch);
        }
    }

    cells.push(current.trim().to_string());
    cells
}

fn parse_table_cell(s: &str) -> TableCell {
    // Multi-paragraph cells use <br> as separator
    let parts: Vec<&str> = s.split("<br>").collect();
    let content = parts
        .into_iter()
        .map(|part| Paragraph {
            inlines: parse_inlines(part),
        })
        .collect();
    TableCell { content }
}

fn unescape_pipes(s: &str) -> String {
    s.replace("\\|", "|")
}

// ---------------------------------------------------------------------------
// Comment parsing
// ---------------------------------------------------------------------------

static COMMENT_WITH_AUTHOR_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"^\[\^([^|\]]+)\|([^\]]*)\]:\s*(.*)$").expect("comment with author regex")
});

static COMMENT_NO_AUTHOR_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^\[\^([^\]]+)\]:\s*(.*)$").expect("comment no author regex"));

fn parse_comment_line(line: &str) -> Option<CommentNote> {
    // Try with-author first: [^id|author]: text
    if let Some(caps) = COMMENT_WITH_AUTHOR_RE.captures(line) {
        return Some(CommentNote {
            id: caps.get(1).unwrap().as_str().to_string(),
            author: unescape_pipes(caps.get(2).unwrap().as_str()),
            text: unescape_pipes(caps.get(3).unwrap().as_str()),
        });
    }
    // Fallback: [^id]: text (no author)
    if let Some(caps) = COMMENT_NO_AUTHOR_RE.captures(line) {
        return Some(CommentNote {
            id: caps.get(1).unwrap().as_str().to_string(),
            author: String::new(),
            text: unescape_pipes(caps.get(2).unwrap().as_str()),
        });
    }
    None
}

// ---------------------------------------------------------------------------
// Formula parsing
// ---------------------------------------------------------------------------

static FORMULA_HUMAN_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"^\[\^f\d+\]:\s*(\S+)\s*=\s*`(=[^`]+)`$").expect("formula human regex")
});

static FORMULA_COMPACT_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^(\S+)=`(=[^`]+)`$").expect("formula compact regex"));

fn parse_formula_line(line: &str, profile: MarkdownProfile) -> Option<FormulaNote> {
    let (re, cell_idx, formula_idx) = match profile {
        MarkdownProfile::Human => (&*FORMULA_HUMAN_RE, 1, 2),
        MarkdownProfile::LlmCompact => (&*FORMULA_COMPACT_RE, 1, 2),
    };
    let caps = re.captures(line)?;
    Some(FormulaNote {
        cell_ref: caps.get(cell_idx)?.as_str().to_string(),
        formula: caps.get(formula_idx)?.as_str().to_string(),
    })
}

// ---------------------------------------------------------------------------
// PDF parser
// ---------------------------------------------------------------------------

static PAGE_HEADER_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^## Page: \d+$").expect("page header regex"));

fn parse_pdf(body: &str) -> OoxmlDocument {
    // PDF pages can contain arbitrary markdown including `## ` headers,
    // so we split only at `## Page: N` boundaries.
    let sections = split_at_line(body, |line| PAGE_HEADER_RE.is_match(line));
    let mut pages = Vec::new();

    for (header, content) in sections {
        if let Some(rest) = header.strip_prefix("## Page: ")
            && let Ok(num) = rest.trim().parse::<usize>()
        {
            pages.push(PdfPage {
                number: num,
                markdown: content.to_string(),
            });
        }
    }

    OoxmlDocument {
        kind: DocumentKind::Pdf,
        pdf: Some(PdfDocument {
            pages,
            diagnostics: Default::default(),
        }),
        ..Default::default()
    }
}

/// Split markdown at lines matching a predicate.
/// Returns (matched_header_line, body_text_between_headers) pairs.
/// Body text has leading/trailing whitespace trimmed.
fn split_at_line(body: &str, is_boundary: impl Fn(&str) -> bool) -> Vec<(&str, &str)> {
    let mut sections = Vec::new();
    let mut current_header: Option<&str> = None;
    let mut current_start: usize = 0;
    let mut pos: usize = 0;

    for line in body.lines() {
        let line_end = pos + line.len() + 1;
        if is_boundary(line) {
            if let Some(header) = current_header {
                sections.push((header, body[current_start..pos].trim()));
            }
            current_header = Some(line);
            current_start = line_end.min(body.len());
        }
        pos = line_end.min(body.len());
    }
    if let Some(header) = current_header {
        sections.push((header, body[current_start..].trim()));
    }
    sections
}

// ---------------------------------------------------------------------------
// XLSX parser
// ---------------------------------------------------------------------------

fn parse_xlsx(
    body: &str,
    profile: MarkdownProfile,
    use_first_row_as_header: bool,
) -> OoxmlDocument {
    let sections = split_at_h2(body);
    let mut sheets = Vec::new();

    for (header, content) in sections {
        if let Some(name) = header.strip_prefix("## Sheet: ") {
            let name = name.trim().to_string();
            let lines: Vec<&str> = content.lines().collect();
            let mut tables = Vec::new();
            let mut formulas = Vec::new();
            let mut i = 0;

            while i < lines.len() {
                let line = lines[i];

                // Skip empty lines
                if line.is_empty() {
                    i += 1;
                    continue;
                }

                // Check for formula section
                if line == "### Formulas" {
                    i += 1;
                    while i < lines.len() && !lines[i].is_empty() {
                        if let Some(f) = parse_formula_line(lines[i], profile) {
                            formulas.push(f);
                        }
                        i += 1;
                    }
                    continue;
                }

                // Check for table (starts with ### caption or |)
                if line.starts_with("### ") || line.starts_with('|') {
                    let (table, consumed) = parse_table_block(&lines[i..], use_first_row_as_header);
                    tables.push(table);
                    i += consumed;
                    continue;
                }

                // Formula lines in compact mode (no ### Formulas header)
                if let Some(f) = parse_formula_line(line, profile) {
                    formulas.push(f);
                    i += 1;
                    continue;
                }

                i += 1;
            }

            sheets.push(Sheet {
                name,
                tables,
                formulas,
                hyperlinks: Vec::new(),
            });
        }
    }

    OoxmlDocument {
        kind: DocumentKind::Xlsx,
        sheets,
        ..Default::default()
    }
}

// ---------------------------------------------------------------------------
// DOCX parser
// ---------------------------------------------------------------------------

fn parse_docx(
    body: &str,
    _profile: MarkdownProfile,
    use_first_row_as_header: bool,
) -> OoxmlDocument {
    let sections = split_at_h2(body);
    let mut doc_sections = Vec::new();

    for (header, content) in sections {
        if let Some(name) = header.strip_prefix("## Section: ") {
            let name = name.trim().to_string();
            let lines: Vec<&str> = content.lines().collect();
            let mut blocks = Vec::new();
            let mut comments = Vec::new();
            let mut i = 0;
            let mut in_comments = false;

            while i < lines.len() {
                let line = lines[i];

                if line == "### Comments" {
                    in_comments = true;
                    i += 1;
                    continue;
                }

                if in_comments {
                    if line.is_empty() {
                        i += 1;
                        continue;
                    }
                    if let Some(c) = parse_comment_line(line) {
                        comments.push(c);
                    }
                    i += 1;
                    continue;
                }

                // Empty line - skip
                if line.is_empty() {
                    i += 1;
                    continue;
                }

                // Separator
                if line == "---" {
                    blocks.push(Block::Separator);
                    i += 1;
                    // Skip blank line after separator
                    if i < lines.len() && lines[i].is_empty() {
                        i += 1;
                    }
                    continue;
                }

                // Table (### caption or |)
                if line.starts_with("### ") || line.starts_with('|') {
                    let (table, consumed) = parse_table_block(&lines[i..], use_first_row_as_header);
                    blocks.push(Block::Table(table));
                    i += consumed;
                    continue;
                }

                // Paragraph
                blocks.push(Block::Paragraph(Paragraph {
                    inlines: parse_inlines(line),
                }));
                i += 1;
            }

            doc_sections.push(DocSection {
                name,
                blocks,
                comments,
            });
        }
    }

    OoxmlDocument {
        kind: DocumentKind::Docx,
        sections: doc_sections,
        ..Default::default()
    }
}

// ---------------------------------------------------------------------------
// PPTX parser
// ---------------------------------------------------------------------------

static SLIDE_HEADER_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^## Slide (\d+)(?:\s*-\s*(.+))?$").expect("slide header regex"));

fn parse_pptx(
    body: &str,
    _profile: MarkdownProfile,
    use_first_row_as_header: bool,
) -> OoxmlDocument {
    let sections = split_at_h2(body);
    let mut slides = Vec::new();

    for (header, content) in sections {
        if let Some(caps) = SLIDE_HEADER_RE.captures(header) {
            let number: usize = caps.get(1).unwrap().as_str().parse().unwrap_or(0);
            let title = caps.get(2).map(|m| unescape_pipes(m.as_str().trim()));

            let lines: Vec<&str> = content.lines().collect();
            let mut blocks: Vec<Block> = Vec::new();
            let mut notes: Option<Vec<Paragraph>> = None;
            let mut comments = Vec::new();
            let mut i = 0;

            #[derive(PartialEq)]
            enum Section {
                Body,
                Notes,
                Comments,
            }
            let mut current_section = Section::Body;

            while i < lines.len() {
                let line = lines[i];

                if line == "### Notes" {
                    current_section = Section::Notes;
                    notes = Some(Vec::new());
                    i += 1;
                    continue;
                }

                if line == "### Comments" {
                    current_section = Section::Comments;
                    i += 1;
                    continue;
                }

                // Don't process past a ### in the wrong section
                if line.starts_with("### ") && current_section != Section::Body {
                    // Unknown subsection, skip
                    i += 1;
                    continue;
                }

                match current_section {
                    Section::Body => {
                        if line.is_empty() {
                            i += 1;
                            continue;
                        }

                        if line == "---" {
                            blocks.push(Block::Separator);
                            i += 1;
                            if i < lines.len() && lines[i].is_empty() {
                                i += 1;
                            }
                            continue;
                        }

                        // Table
                        if line.starts_with("### ") || line.starts_with('|') {
                            let (table, consumed) =
                                parse_table_block(&lines[i..], use_first_row_as_header);
                            blocks.push(Block::Table(table));
                            i += consumed;
                            continue;
                        }

                        // Paragraph
                        blocks.push(Block::Paragraph(Paragraph {
                            inlines: parse_inlines(line),
                        }));
                        i += 1;
                    }
                    Section::Notes => {
                        if line.is_empty() {
                            i += 1;
                            continue;
                        }
                        if let Some(ref mut note_paras) = notes {
                            note_paras.push(Paragraph {
                                inlines: parse_inlines(line),
                            });
                        }
                        i += 1;
                    }
                    Section::Comments => {
                        if line.is_empty() {
                            i += 1;
                            continue;
                        }
                        if let Some(c) = parse_comment_line(line) {
                            comments.push(c);
                        }
                        i += 1;
                    }
                }
            }

            slides.push(Slide {
                number,
                title,
                blocks,
                notes,
                comments,
            });
        }
    }

    OoxmlDocument {
        kind: DocumentKind::Pptx,
        slides,
        ..Default::default()
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{RenderOptions, render_document_with_options};

    #[test]
    fn frontmatter_round_trip() {
        let input = "<!-- officemd: kind=xlsx profile=compact first_row_as_header=true formulas=true headers_footers=true properties=false -->\n\n## Sheet: Data\n\n";
        let fm = parse_frontmatter(input).unwrap();
        assert_eq!(fm.kind, DocumentKind::Xlsx);
        assert_eq!(
            fm.render_options.markdown_profile,
            MarkdownProfile::LlmCompact
        );
        assert!(fm.render_options.use_first_row_as_header);
        assert!(fm.render_options.include_formulas);
    }

    #[test]
    fn detect_kind_from_headers() {
        assert_eq!(detect_kind("## Sheet: Data\n"), Some(DocumentKind::Xlsx));
        assert_eq!(detect_kind("## Section: body\n"), Some(DocumentKind::Docx));
        assert_eq!(detect_kind("## Slide 1\n"), Some(DocumentKind::Pptx));
        assert_eq!(detect_kind("## Page: 1\n"), Some(DocumentKind::Pdf));
        assert_eq!(detect_kind("nothing here\n"), None);
    }

    #[test]
    fn split_table_row_basic() {
        let cells = split_table_row("| A | B | C |");
        assert_eq!(cells, vec!["A", "B", "C"]);
    }

    #[test]
    fn split_table_row_escaped_pipes() {
        let cells = split_table_row("| A\\|B | C |");
        assert_eq!(cells, vec!["A|B", "C"]);
    }

    #[test]
    fn parse_inlines_plain() {
        let inlines = parse_inlines("hello world");
        assert_eq!(inlines.len(), 1);
        assert!(matches!(&inlines[0], Inline::Text(t) if t == "hello world"));
    }

    #[test]
    fn parse_inlines_link() {
        let inlines = parse_inlines("Visit [Example](https://example.com) now");
        assert_eq!(inlines.len(), 3);
        assert!(matches!(&inlines[0], Inline::Text(t) if t == "Visit "));
        assert!(
            matches!(&inlines[1], Inline::Link(h) if h.display == "Example" && h.target == "https://example.com")
        );
        assert!(matches!(&inlines[2], Inline::Text(t) if t == " now"));
    }

    #[test]
    fn parse_inlines_autolink() {
        let inlines = parse_inlines("<https://example.com>");
        assert_eq!(inlines.len(), 1);
        assert!(
            matches!(&inlines[0], Inline::Link(h) if h.display.is_empty() && h.target == "https://example.com")
        );
    }

    #[test]
    fn parse_comment_with_author() {
        let c = parse_comment_line("[^c1|Alice]: Review").unwrap();
        assert_eq!(c.id, "c1");
        assert_eq!(c.author, "Alice");
        assert_eq!(c.text, "Review");
    }

    #[test]
    fn parse_comment_without_author() {
        let c = parse_comment_line("[^c1]: Anonymous note").unwrap();
        assert_eq!(c.id, "c1");
        assert!(c.author.is_empty());
        assert_eq!(c.text, "Anonymous note");
    }

    #[test]
    fn roundtrip_pdf() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pdf,
            pdf: Some(PdfDocument {
                pages: vec![
                    PdfPage {
                        number: 1,
                        markdown: "Page one body".into(),
                    },
                    PdfPage {
                        number: 2,
                        markdown: "Page two body".into(),
                    },
                ],
                diagnostics: Default::default(),
            }),
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_xlsx() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![Sheet {
                name: "Data".into(),
                tables: vec![Table {
                    caption: Some("Table 1 (rows 1-3, cols A-B)".into()),
                    headers: vec!["Col1".into(), "Col2".into()],
                    rows: vec![
                        vec![
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("Region".into())],
                                }],
                            },
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("Revenue".into())],
                                }],
                            },
                        ],
                        vec![
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("US".into())],
                                }],
                            },
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("$100".into())],
                                }],
                            },
                        ],
                        vec![
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("EU".into())],
                                }],
                            },
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("$200".into())],
                                }],
                            },
                        ],
                    ],
                    synthetic_headers: true,
                }],
                formulas: vec![FormulaNote {
                    cell_ref: "C1".into(),
                    formula: "=A1+B1".into(),
                }],
                hyperlinks: Vec::new(),
            }],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_xlsx_human() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![Sheet {
                name: "Data".into(),
                tables: vec![Table {
                    caption: Some("Table 1 (rows 1-2, cols A-B)".into()),
                    headers: vec!["Col1".into(), "Col2".into()],
                    rows: vec![
                        vec![
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("A".into())],
                                }],
                            },
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("B".into())],
                                }],
                            },
                        ],
                        vec![
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("1".into())],
                                }],
                            },
                            TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("2".into())],
                                }],
                            },
                        ],
                    ],
                    synthetic_headers: true,
                }],
                formulas: vec![FormulaNote {
                    cell_ref: "C1".into(),
                    formula: "=SUM(A1:B1)".into(),
                }],
                hyperlinks: Vec::new(),
            }],
            ..Default::default()
        };

        let opts = RenderOptions {
            markdown_profile: MarkdownProfile::Human,
            ..Default::default()
        };
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_docx() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![
                DocSection {
                    name: "body".into(),
                    blocks: vec![
                        Block::Paragraph(Paragraph {
                            inlines: vec![
                                Inline::Text("Hello ".into()),
                                Inline::Link(Hyperlink {
                                    display: "World".into(),
                                    target: "https://example.com".into(),
                                    rel_id: None,
                                }),
                            ],
                        }),
                        Block::Separator,
                        Block::Paragraph(Paragraph {
                            inlines: vec![Inline::Text("After separator".into())],
                        }),
                    ],
                    comments: vec![CommentNote {
                        id: "c0".into(),
                        author: "Bob".into(),
                        text: "Check this".into(),
                    }],
                },
                DocSection {
                    name: "footnotes".into(),
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Footnote text".into())],
                    })],
                    comments: vec![],
                },
            ],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_pptx() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![
                Slide {
                    number: 1,
                    title: Some("Intro".into()),
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Hello".into())],
                    })],
                    notes: Some(vec![Paragraph {
                        inlines: vec![Inline::Text("Speaker note".into())],
                    }]),
                    comments: vec![CommentNote {
                        id: "c1".into(),
                        author: "Alice".into(),
                        text: "Review".into(),
                    }],
                },
                Slide {
                    number: 2,
                    title: None,
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Content".into())],
                    })],
                    notes: None,
                    comments: vec![],
                },
            ],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_docx_with_table() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![Block::Table(Table {
                    caption: None,
                    headers: vec!["Col1".into()],
                    rows: vec![vec![TableCell {
                        content: vec![Paragraph {
                            inlines: vec![Inline::Text("Cell".into())],
                        }],
                    }]],
                    synthetic_headers: true,
                })],
                comments: vec![],
            }],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_comment_without_author() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![],
                comments: vec![CommentNote {
                    id: "c1".into(),
                    author: "".into(),
                    text: "Anonymous note".into(),
                }],
            }],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_multi_paragraph_cell() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![Sheet {
                name: "Data".into(),
                tables: vec![Table {
                    caption: None,
                    headers: vec!["Col1".into()],
                    rows: vec![vec![TableCell {
                        content: vec![
                            Paragraph {
                                inlines: vec![Inline::Text("Line 1".into())],
                            },
                            Paragraph {
                                inlines: vec![Inline::Text("Line 2".into())],
                            },
                        ],
                    }]],
                    synthetic_headers: true,
                }],
                formulas: vec![],
                hyperlinks: Vec::new(),
            }],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_xlsx_no_first_row_header() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![Sheet {
                name: "Data".into(),
                tables: vec![Table {
                    caption: None,
                    headers: vec!["Name".into(), "Value".into()],
                    rows: vec![vec![
                        TableCell {
                            content: vec![Paragraph {
                                inlines: vec![Inline::Text("foo".into())],
                            }],
                        },
                        TableCell {
                            content: vec![Paragraph {
                                inlines: vec![Inline::Text("42".into())],
                            }],
                        },
                    ]],
                    synthetic_headers: false,
                }],
                formulas: vec![],
                hyperlinks: Vec::new(),
            }],
            ..Default::default()
        };

        let opts = RenderOptions {
            use_first_row_as_header: false,
            ..Default::default()
        };
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_escaped_pipes_in_cells() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![Sheet {
                name: "Data".into(),
                tables: vec![Table {
                    caption: None,
                    headers: vec!["Col1".into()],
                    rows: vec![vec![TableCell {
                        content: vec![Paragraph {
                            inlines: vec![Inline::Text("A|B".into())],
                        }],
                    }]],
                    synthetic_headers: true,
                }],
                formulas: vec![],
                hyperlinks: Vec::new(),
            }],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }

    #[test]
    fn roundtrip_pptx_with_escaped_pipes() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![Slide {
                number: 1,
                title: Some("A | B".into()),
                blocks: vec![],
                notes: None,
                comments: vec![CommentNote {
                    id: "c1".into(),
                    author: "A|B".into(),
                    text: "C|D".into(),
                }],
            }],
            ..Default::default()
        };

        let opts = RenderOptions::default();
        let md1 = render_document_with_options(&doc, opts);
        let parsed = parse_document(&md1).unwrap();
        let md2 = render_document_with_options(&parsed, opts);
        assert_eq!(md1, md2);
    }
}
