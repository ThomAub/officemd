//! Markdown renderer and parser for OOXML IR.

pub mod parse;

pub use parse::{ParseError, ParseOptions, parse_document, parse_document_with_options};

use std::borrow::Cow;
use std::fmt::Write;

use officemd_core::ir::{
    Block, CommentNote, DocumentKind, Inline, OoxmlDocument, Paragraph, Table, TableCell,
};

/// Markdown output profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkdownProfile {
    /// Compact markdown tuned for LLM consumption.
    LlmCompact,
    /// Verbose markdown tuned for human readability.
    Human,
}

/// Options controlling markdown rendering behavior.
#[derive(Debug, Clone, Copy)]
pub struct RenderOptions {
    /// Include document-level properties as a header block.
    pub include_document_properties: bool,
    /// Use the first data row as the table header instead of synthetic Col1/Col2 names.
    pub use_first_row_as_header: bool,
    /// Include DOCX header/footer sections in the output.
    pub include_headers_footers: bool,
    /// Include XLSX formula footnotes in the output.
    pub include_formulas: bool,
    /// Markdown profile controlling compact vs human-oriented formatting.
    pub markdown_profile: MarkdownProfile,
}

impl Default for RenderOptions {
    fn default() -> Self {
        Self {
            include_document_properties: false,
            use_first_row_as_header: true,
            include_headers_footers: true,
            include_formulas: true,
            markdown_profile: MarkdownProfile::LlmCompact,
        }
    }
}

/// Render a full OOXML document to Markdown.
#[must_use]
pub fn render_document(doc: &OoxmlDocument) -> String {
    render_document_with_options(doc, RenderOptions::default())
}

/// Render a full OOXML document to Markdown with explicit options.
#[must_use]
pub fn render_document_with_options(doc: &OoxmlDocument, options: RenderOptions) -> String {
    let mut out = render_frontmatter(doc, options);
    match doc.kind {
        DocumentKind::Xlsx => out.push_str(&render_xlsx(doc, options)),
        DocumentKind::Docx => out.push_str(&render_docx(doc, options)),
        DocumentKind::Pptx => out.push_str(&render_pptx(doc, options)),
        DocumentKind::Pdf => out.push_str(&render_pdf(doc, options)),
    }
    out
}

fn render_frontmatter(doc: &OoxmlDocument, options: RenderOptions) -> String {
    let kind = doc.kind.as_str();
    let profile = match options.markdown_profile {
        MarkdownProfile::LlmCompact => "compact",
        MarkdownProfile::Human => "human",
    };
    format!(
        "<!-- officemd: kind={kind} profile={profile} first_row_as_header={} formulas={} headers_footers={} properties={} -->\n\n",
        options.use_first_row_as_header,
        options.include_formulas,
        options.include_headers_footers,
        options.include_document_properties,
    )
}

fn render_xlsx(doc: &OoxmlDocument, options: RenderOptions) -> String {
    let mut out = String::new();

    if options.include_document_properties {
        render_properties(doc, &mut out, options);
    }

    for sheet in &doc.sheets {
        let _ = write!(out, "## Sheet: {}\n\n", sheet.name);

        for table in &sheet.tables {
            out.push_str(&render_table(table, options));
            out.push('\n');
        }

        if options.include_formulas && !sheet.formulas.is_empty() {
            if matches!(options.markdown_profile, MarkdownProfile::Human) {
                out.push_str("### Formulas\n");
            }
            for (i, note) in sheet.formulas.iter().enumerate() {
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
    out
}

fn render_docx(doc: &OoxmlDocument, options: RenderOptions) -> String {
    let mut out = String::new();

    if options.include_document_properties {
        render_properties(doc, &mut out, options);
    }

    let mut wrote_section = false;

    for section in &doc.sections {
        if !options.include_headers_footers
            && (section.name.starts_with("header") || section.name.starts_with("footer"))
        {
            continue;
        }

        if wrote_section {
            out.push('\n');
        }
        wrote_section = true;

        let _ = write!(out, "## Section: {}\n\n", section.name);

        for block in &section.blocks {
            match block {
                Block::Paragraph(p) => {
                    out.push_str(&render_paragraph(p, false));
                    out.push('\n');
                }
                Block::Table(table) => {
                    out.push_str(&render_table(table, options));
                    out.push('\n');
                }
                Block::Separator => out.push_str("---\n\n"),
            }
        }

        if !section.comments.is_empty() {
            out.push_str("### Comments\n");
            for note in &section.comments {
                render_comment(note, &mut out, false);
            }
            out.push('\n');
        }
    }
    out
}

fn render_pptx(doc: &OoxmlDocument, options: RenderOptions) -> String {
    let mut out = String::new();

    if options.include_document_properties {
        render_properties(doc, &mut out, options);
    }

    for slide in &doc.slides {
        match &slide.title {
            Some(title) => {
                let _ = write!(
                    out,
                    "## Slide {} - {}\n\n",
                    slide.number,
                    escape_pipes(title)
                );
            }
            None => {
                let _ = write!(out, "## Slide {}\n\n", slide.number);
            }
        }

        let mut skipped_title_duplicate = false;
        for block in &slide.blocks {
            match block {
                Block::Paragraph(p) => {
                    if !skipped_title_duplicate
                        && slide
                            .title
                            .as_ref()
                            .is_some_and(|title| is_duplicate_slide_title_paragraph(p, title))
                    {
                        skipped_title_duplicate = true;
                        continue;
                    }
                    out.push_str(&render_paragraph(p, false));
                    out.push('\n');
                }
                Block::Table(table) => {
                    out.push_str(&render_table(table, options));
                    out.push('\n');
                }
                Block::Separator => out.push_str("---\n\n"),
            }
        }

        if let Some(notes) = &slide.notes
            && !notes.is_empty()
        {
            out.push_str("### Notes\n");
            for paragraph in notes {
                out.push_str(&render_paragraph(paragraph, false));
                out.push('\n');
            }
            out.push('\n');
        }

        if !slide.comments.is_empty() {
            out.push_str("### Comments\n");
            for note in &slide.comments {
                render_comment(note, &mut out, true);
            }
            out.push('\n');
        }
    }

    out
}

fn is_duplicate_slide_title_paragraph(paragraph: &Paragraph, title: &str) -> bool {
    if paragraph
        .inlines
        .iter()
        .any(|inline| matches!(inline, Inline::Link(_)))
    {
        return false;
    }

    let text = render_paragraph(paragraph, false);
    text.trim() == title.trim()
}

fn render_pdf(doc: &OoxmlDocument, options: RenderOptions) -> String {
    let mut out = String::new();

    if options.include_document_properties {
        render_properties(doc, &mut out, options);
    }

    let Some(pdf) = &doc.pdf else {
        return out;
    };

    for page in &pdf.pages {
        let _ = write!(out, "## Page: {}\n\n", page.number);
        if !page.markdown.is_empty() {
            out.push_str(&page.markdown);
            if !page.markdown.ends_with('\n') {
                out.push('\n');
            }
            out.push('\n');
        }
    }

    out
}

fn render_properties(doc: &OoxmlDocument, out: &mut String, options: RenderOptions) {
    if let Some(props) = &doc.properties
        && (!props.core.is_empty() || !props.app.is_empty() || !props.custom.is_empty())
    {
        let mut entries = props
            .core
            .iter()
            .chain(props.app.iter())
            .chain(props.custom.iter())
            .map(|(k, v)| (k.as_str(), v.as_str()))
            .collect::<Vec<_>>();
        entries.sort_unstable_by(|(ka, va), (kb, vb)| ka.cmp(kb).then_with(|| va.cmp(vb)));

        if matches!(options.markdown_profile, MarkdownProfile::Human) {
            out.push_str("### Document Properties\n");
            for (k, v) in entries {
                let _ = writeln!(out, "- {}: {}", k, escape_pipes(v));
            }
            out.push_str("\n---\n\n");
        } else {
            out.push_str("properties: ");
            for (idx, (k, v)) in entries.iter().enumerate() {
                if idx > 0 {
                    out.push_str("; ");
                }
                let _ = write!(out, "{k}={}", escape_pipes(v));
            }
            out.push_str("\n\n");
        }
    }
}

fn render_table(table: &Table, options: RenderOptions) -> String {
    let mut out = String::new();
    if let Some(caption) = &table.caption {
        let _ = writeln!(out, "### {caption}");
    }

    let (header_labels, data_rows): (Vec<String>, &[Vec<TableCell>]) =
        if options.use_first_row_as_header && !table.rows.is_empty() {
            let labels = table.rows[0].iter().map(render_cell).collect();
            (labels, &table.rows[1..])
        } else {
            (
                table.headers.iter().map(|h| escape_pipes(h)).collect(),
                &table.rows,
            )
        };

    // headers
    out.push('|');
    for h in &header_labels {
        out.push(' ');
        out.push_str(h);
        out.push(' ');
        out.push('|');
    }
    out.push('\n');

    // separator
    out.push('|');
    for _ in &header_labels {
        out.push_str(" --- |");
    }
    out.push('\n');

    for row in data_rows {
        out.push('|');
        for cell in row {
            out.push(' ');
            out.push_str(&render_cell(cell));
            out.push(' ');
            out.push('|');
        }
        out.push('\n');
    }

    out
}

fn render_cell(cell: &TableCell) -> String {
    let mut out = String::new();
    for (i, para) in cell.content.iter().enumerate() {
        if i > 0 {
            out.push_str("<br>");
        }
        out.push_str(&render_paragraph(para, true));
    }
    out
}

fn render_paragraph(paragraph: &Paragraph, escape_pipes_in_text: bool) -> String {
    render_inlines(&paragraph.inlines, escape_pipes_in_text)
}

fn render_inlines(inlines: &[Inline], escape_pipes_in_text: bool) -> String {
    let mut buf = String::new();
    for inline in inlines {
        match inline {
            Inline::Text(text) => buf.push_str(&escape_text(text, escape_pipes_in_text)),
            Inline::Link(link) => {
                let display = escape_text(&link.display, escape_pipes_in_text);
                let target = escape_text(&link.target, escape_pipes_in_text);
                if target.is_empty() {
                    buf.push_str(&display);
                } else if link.display.is_empty() {
                    let _ = write!(buf, "<{target}>");
                } else {
                    let _ = write!(buf, "[{display}]({target})");
                }
            }
        }
    }
    buf
}

fn escape_text(s: &str, escape_pipes_in_text: bool) -> Cow<'_, str> {
    if escape_pipes_in_text && s.contains('|') {
        Cow::Owned(escape_pipes(s))
    } else {
        Cow::Borrowed(s)
    }
}

fn escape_pipes(s: &str) -> String {
    s.replace('|', "\\|")
}

fn render_comment(note: &CommentNote, out: &mut String, do_escape_pipes: bool) {
    if note.author.is_empty() {
        let _ = writeln!(
            out,
            "[^{}]: {}",
            note.id,
            escape_text(&note.text, do_escape_pipes)
        );
    } else {
        let _ = writeln!(
            out,
            "[^{}|{}]: {}",
            note.id,
            escape_text(&note.author, do_escape_pipes),
            escape_text(&note.text, do_escape_pipes),
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use officemd_core::ir::{
        Block, CommentNote, DocSection, DocumentProperties, FormulaNote, Hyperlink, Inline,
        Paragraph, Sheet, Slide, Table, TableCell,
    };
    use std::collections::HashMap;

    #[test]
    fn renders_basic_table_with_synthetic_headers() {
        let table = Table {
            caption: None,
            headers: vec!["Col1".into()],
            rows: vec![vec![TableCell {
                content: vec![Paragraph {
                    inlines: vec![Inline::Text("Hello".into())],
                }],
            }]],
            synthetic_headers: true,
        };
        let opts = RenderOptions {
            use_first_row_as_header: false,
            ..Default::default()
        };
        let md = render_table(&table, opts);
        assert!(md.contains("Col1"));
        assert!(md.contains("Hello"));
    }

    #[test]
    fn first_row_promoted_to_header() {
        let table = Table {
            caption: None,
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
            ],
            synthetic_headers: true,
        };
        let md = render_table(&table, RenderOptions::default());
        assert!(md.contains("| Region | Revenue |"));
        assert!(md.contains("| US | $100 |"));
        assert!(!md.contains("Col1"));
    }

    #[test]
    fn synthetic_headers_when_disabled() {
        let table = Table {
            caption: None,
            headers: vec!["Col1".into(), "Col2".into()],
            rows: vec![vec![
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
            ]],
            synthetic_headers: true,
        };
        let opts = RenderOptions {
            use_first_row_as_header: false,
            ..Default::default()
        };
        let md = render_table(&table, opts);
        assert!(md.contains("| Col1 | Col2 |"));
        assert!(md.contains("| A | B |"));
    }

    #[test]
    fn empty_table_uses_synthetic_headers() {
        let table = Table {
            caption: None,
            headers: vec!["Col1".into()],
            rows: vec![],
            synthetic_headers: true,
        };
        let md = render_table(&table, RenderOptions::default());
        assert!(md.contains("| Col1 |"));
    }

    #[test]
    fn renders_table_with_caption() {
        let table = Table {
            caption: Some("Sales Data".into()),
            headers: vec!["Product".into(), "Revenue".into()],
            rows: vec![
                vec![
                    TableCell {
                        content: vec![Paragraph {
                            inlines: vec![Inline::Text("Product".into())],
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
                            inlines: vec![Inline::Text("Widget".into())],
                        }],
                    },
                    TableCell {
                        content: vec![Paragraph {
                            inlines: vec![Inline::Text("$100".into())],
                        }],
                    },
                ],
            ],
            synthetic_headers: false,
        };
        let md = render_table(
            &table,
            RenderOptions {
                markdown_profile: MarkdownProfile::Human,
                ..Default::default()
            },
        );
        assert!(md.contains("### Sales Data"));
        assert!(md.contains("| Product | Revenue |"));
    }

    #[test]
    fn renders_pptx_slide() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![Slide {
                number: 1,
                title: Some("Intro".into()),
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Hello".into())],
                })],
                notes: Some(vec![Paragraph {
                    inlines: vec![Inline::Text("Note".into())],
                }]),
                comments: vec![CommentNote {
                    id: "c1".into(),
                    author: "Alice".into(),
                    text: "Review".into(),
                }],
            }],
            ..Default::default()
        };

        let md = render_document(&doc);
        assert!(md.contains("## Slide 1 - Intro"));
        assert!(md.contains("Hello"));
        assert!(md.contains("### Notes"));
        assert!(md.contains("### Comments"));
        assert!(md.contains("[^c1|Alice]: Review"));
    }

    #[test]
    fn renders_xlsx_sheets() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![
                Sheet {
                    name: "Data".into(),
                    tables: vec![Table {
                        caption: None,
                        headers: vec!["A".into(), "B".into()],
                        rows: vec![vec![
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
                        ]],
                        synthetic_headers: true,
                    }],
                    formulas: vec![FormulaNote {
                        cell_ref: "C1".into(),
                        formula: "=A1+B1".into(),
                    }],
                    hyperlinks: vec![],
                },
                Sheet {
                    name: "Summary".into(),
                    tables: vec![],
                    formulas: vec![],
                    hyperlinks: vec![],
                },
            ],
            ..Default::default()
        };

        let md = render_document_with_options(
            &doc,
            RenderOptions {
                markdown_profile: MarkdownProfile::Human,
                ..Default::default()
            },
        );
        assert!(md.contains("## Sheet: Data"));
        assert!(md.contains("## Sheet: Summary"));
        // With default use_first_row_as_header, first data row ("1", "2") becomes header
        assert!(md.contains("| 1 | 2 |"));
        assert!(md.contains("### Formulas"));
        assert!(md.contains("[^f1]: C1 = `=A1+B1`"));
    }

    #[test]
    fn renders_docx_sections() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![
                DocSection {
                    name: "body".into(),
                    blocks: vec![
                        Block::Paragraph(Paragraph {
                            inlines: vec![Inline::Text("Hello World".into())],
                        }),
                        Block::Separator,
                        Block::Table(Table {
                            caption: None,
                            headers: vec!["Col".into()],
                            rows: vec![vec![TableCell {
                                content: vec![Paragraph {
                                    inlines: vec![Inline::Text("Cell".into())],
                                }],
                            }]],
                            synthetic_headers: true,
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

        let md = render_document(&doc);
        assert!(md.contains("## Section: body"));
        assert!(md.contains("Hello World"));
        assert!(md.contains("---")); // separator
        // With default use_first_row_as_header, first row ("Cell") becomes header
        assert!(md.contains("| Cell |"));
        assert!(md.contains("### Comments"));
        assert!(md.contains("[^c0|Bob]: Check this"));
        assert!(md.contains("## Section: footnotes"));
        assert!(md.contains("Footnote text"));
    }

    #[test]
    fn renders_hyperlinks() {
        let paragraph = Paragraph {
            inlines: vec![
                Inline::Text("Visit ".into()),
                Inline::Link(Hyperlink {
                    display: "Example".into(),
                    target: "https://example.com".into(),
                    rel_id: None,
                }),
                Inline::Text(" for more.".into()),
            ],
        };

        let md = render_paragraph(&paragraph, false);
        assert!(md.contains("Visit "));
        assert!(md.contains("[Example](https://example.com)"));
        assert!(md.contains(" for more."));
    }

    #[test]
    fn renders_link_without_display() {
        let paragraph = Paragraph {
            inlines: vec![Inline::Link(Hyperlink {
                display: "".into(),
                target: "https://example.com".into(),
                rel_id: None,
            })],
        };

        let md = render_paragraph(&paragraph, false);
        assert!(md.contains("<https://example.com>"));
    }

    #[test]
    fn renders_document_properties() {
        let mut core = HashMap::new();
        core.insert("title".into(), "My Document".into());
        core.insert("creator".into(), "John Doe".into());

        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            properties: Some(DocumentProperties {
                core,
                app: HashMap::new(),
                custom: HashMap::new(),
            }),
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![],
                comments: vec![],
            }],
            ..Default::default()
        };

        let md = render_document_with_options(
            &doc,
            RenderOptions {
                include_document_properties: true,
                markdown_profile: MarkdownProfile::Human,
                ..Default::default()
            },
        );
        assert!(md.contains("### Document Properties"));
        assert!(md.contains("title: My Document") || md.contains("creator: John Doe"));
    }

    #[test]
    fn renders_document_properties_sorted_by_key_then_value() {
        let mut core = HashMap::new();
        core.insert("z_key".into(), "from_core".into());
        core.insert("dup".into(), "2".into());

        let mut app = HashMap::new();
        app.insert("a_key".into(), "from_app".into());
        app.insert("dup".into(), "1".into());

        let mut custom = HashMap::new();
        custom.insert("m_key".into(), "from_custom".into());

        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            properties: Some(DocumentProperties { core, app, custom }),
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![],
                comments: vec![],
            }],
            ..Default::default()
        };

        let md = render_document_with_options(
            &doc,
            RenderOptions {
                include_document_properties: true,
                markdown_profile: MarkdownProfile::Human,
                ..Default::default()
            },
        );

        let a_idx = md.find("- a_key: from_app").expect("a_key line");
        let dup_one_idx = md.find("- dup: 1").expect("dup:1 line");
        let dup_two_idx = md.find("- dup: 2").expect("dup:2 line");
        let m_idx = md.find("- m_key: from_custom").expect("m_key line");
        let z_idx = md.find("- z_key: from_core").expect("z_key line");

        assert!(a_idx < dup_one_idx);
        assert!(dup_one_idx < dup_two_idx);
        assert!(dup_two_idx < m_idx);
        assert!(m_idx < z_idx);
    }

    #[test]
    fn omits_document_properties_by_default() {
        let mut core = HashMap::new();
        core.insert("title".into(), "Hidden By Default".into());

        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            properties: Some(DocumentProperties {
                core,
                app: HashMap::new(),
                custom: HashMap::new(),
            }),
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![],
                comments: vec![],
            }],
            ..Default::default()
        };

        let md = render_document(&doc);
        assert!(!md.contains("### Document Properties"));
        assert!(!md.contains("Hidden By Default"));
    }

    #[test]
    fn escapes_pipes_in_tables() {
        let table = Table {
            caption: None,
            headers: vec!["Data".into()],
            rows: vec![vec![TableCell {
                content: vec![Paragraph {
                    inlines: vec![Inline::Text("A|B".into())],
                }],
            }]],
            synthetic_headers: false,
        };
        let opts = RenderOptions {
            use_first_row_as_header: false,
            ..Default::default()
        };
        let md = render_table(&table, opts);
        assert!(md.contains("A\\|B"));
    }

    #[test]
    fn renders_multiple_slides() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![
                Slide {
                    number: 1,
                    title: Some("First".into()),
                    blocks: vec![],
                    notes: None,
                    comments: vec![],
                },
                Slide {
                    number: 2,
                    title: Some("Second".into()),
                    blocks: vec![],
                    notes: None,
                    comments: vec![],
                },
            ],
            ..Default::default()
        };

        let md = render_document_with_options(
            &doc,
            RenderOptions {
                markdown_profile: MarkdownProfile::Human,
                ..Default::default()
            },
        );
        assert!(md.contains("## Slide 1 - First"));
        assert!(md.contains("## Slide 2 - Second"));
    }

    #[test]
    fn renders_slide_without_title() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![Slide {
                number: 3,
                title: None,
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Content".into())],
                })],
                notes: None,
                comments: vec![],
            }],
            ..Default::default()
        };

        let md = render_document(&doc);
        assert!(md.contains("## Slide 3\n"));
        assert!(!md.contains("## Slide 3 -"));
    }

    #[test]
    fn skips_duplicate_slide_title_paragraph_once() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![Slide {
                number: 1,
                title: Some("Quarterly Review".into()),
                blocks: vec![
                    Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Quarterly Review".into())],
                    }),
                    Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Body text".into())],
                    }),
                ],
                notes: None,
                comments: vec![],
            }],
            ..Default::default()
        };

        let md = render_document(&doc);
        assert_eq!(md.matches("Quarterly Review").count(), 1);
        assert!(md.contains("Body text"));
    }

    #[test]
    fn renders_comment_without_author() {
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

        let md = render_document(&doc);
        assert!(md.contains("[^c1]: Anonymous note"));
        assert!(!md.contains("[^c1]: : Anonymous note"));
    }

    #[test]
    fn headers_footers_excluded() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![
                DocSection {
                    name: "header1".into(),
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Header text".into())],
                    })],
                    comments: vec![],
                },
                DocSection {
                    name: "body".into(),
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Body text".into())],
                    })],
                    comments: vec![],
                },
                DocSection {
                    name: "footer1".into(),
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Footer text".into())],
                    })],
                    comments: vec![],
                },
            ],
            ..Default::default()
        };

        let md = render_document_with_options(
            &doc,
            RenderOptions {
                include_headers_footers: false,
                ..Default::default()
            },
        );
        assert!(!md.contains("header1"));
        assert!(!md.contains("Header text"));
        assert!(!md.contains("footer1"));
        assert!(!md.contains("Footer text"));
        assert!(md.contains("Body text"));
    }

    #[test]
    fn headers_footers_included_by_default() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![
                DocSection {
                    name: "header1".into(),
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Header text".into())],
                    })],
                    comments: vec![],
                },
                DocSection {
                    name: "body".into(),
                    blocks: vec![Block::Paragraph(Paragraph {
                        inlines: vec![Inline::Text("Body text".into())],
                    })],
                    comments: vec![],
                },
            ],
            ..Default::default()
        };

        let md = render_document(&doc);
        assert!(md.contains("## Section: header1"));
        assert!(md.contains("Header text"));
        assert!(md.contains("Body text"));
    }

    #[test]
    fn renders_pdf_pages_with_sections() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pdf,
            pdf: Some(officemd_core::ir::PdfDocument {
                pages: vec![
                    officemd_core::ir::PdfPage {
                        number: 1,
                        markdown: "Page one body".into(),
                    },
                    officemd_core::ir::PdfPage {
                        number: 2,
                        markdown: "Page two body".into(),
                    },
                ],
                diagnostics: officemd_core::ir::PdfDiagnostics::default(),
            }),
            ..Default::default()
        };

        let md = render_document_with_options(
            &doc,
            RenderOptions {
                markdown_profile: MarkdownProfile::Human,
                ..Default::default()
            },
        );
        assert!(md.contains("## Page: 1"));
        assert!(md.contains("Page one body"));
        assert!(md.contains("## Page: 2"));
    }

    #[test]
    fn renders_pdf_frontmatter_only_when_no_pages_and_no_props() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pdf,
            pdf: Some(officemd_core::ir::PdfDocument {
                pages: vec![],
                diagnostics: officemd_core::ir::PdfDiagnostics::default(),
            }),
            ..Default::default()
        };

        let md = render_document(&doc);
        assert!(md.starts_with("<!-- officemd:"));
        assert!(!md.contains("## Page:"));
    }
}
