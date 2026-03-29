//! Generate DOCX files from the officemd IR.
//!
//! Converts an [`OoxmlDocument`] with `kind: Docx` into a valid `.docx` ZIP
//! archive that opens in Microsoft Word and LibreOffice.

use std::fmt::Write as _;

use officemd_core::ir::{Block, CommentNote, Inline, OoxmlDocument, Paragraph, Table};
use officemd_core::opc::writer::{OpcWriter, RelEntry, xml_escape_attr, xml_escape_text};

use crate::error::DocxError;

// --- OOXML constants ---

const CT_DOCUMENT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const REL_TYPE_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_TYPE_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_TYPE_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const NS_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Generate a `.docx` file from an officemd IR document.
///
/// Only the "body" section is written. Header/footer/footnote sections are
/// skipped. Comments attached to the body section are written to
/// `word/comments.xml`.
///
/// # Errors
///
/// Returns an error if ZIP assembly fails.
pub fn generate_docx(doc: &OoxmlDocument) -> Result<Vec<u8>, DocxError> {
    let body_section = doc
        .sections
        .iter()
        .find(|s| s.name == "body")
        .or_else(|| doc.sections.first());

    let (blocks, comments) = match body_section {
        Some(section) => (&section.blocks[..], &section.comments[..]),
        None => (&[][..], &[][..]),
    };

    let mut doc_rels: Vec<RelEntry> = Vec::new();
    let mut rel_counter: usize = 1;

    // Build word/document.xml
    let document_xml = build_document_xml(blocks, comments, &mut doc_rels, &mut rel_counter);

    // Build word/comments.xml if comments exist
    let comments_xml = if comments.is_empty() {
        None
    } else {
        // Add relationship for comments
        doc_rels.push(RelEntry {
            id: format!("rId{rel_counter}"),
            rel_type: REL_TYPE_COMMENTS.to_string(),
            target: "comments.xml".to_string(),
            target_mode: None,
        });
        rel_counter += 1;
        let _ = rel_counter; // suppress unused warning
        Some(build_comments_xml(comments))
    };

    // Assemble ZIP
    let mut w = OpcWriter::new();
    w.register_content_type_default(
        "rels",
        "application/vnd.openxmlformats-package.relationships+xml",
    );
    w.register_content_type_default("xml", "application/xml");
    w.register_content_type_override("/word/document.xml", CT_DOCUMENT);

    w.add_xml_part("word/document.xml", &document_xml)?;

    if let Some(ref xml) = comments_xml {
        w.register_content_type_override("/word/comments.xml", CT_COMMENTS);
        w.add_xml_part("word/comments.xml", xml)?;
    }

    w.add_part_rels("word/document.xml", &doc_rels)?;

    w.add_root_relationship(RelEntry {
        id: "rId1".to_string(),
        rel_type: REL_TYPE_OFFICE_DOC.to_string(),
        target: "word/document.xml".to_string(),
        target_mode: None,
    });

    Ok(w.finish()?)
}

// --- XML builders ---

fn build_document_xml(
    blocks: &[Block],
    comments: &[CommentNote],
    rels: &mut Vec<RelEntry>,
    rel_counter: &mut usize,
) -> String {
    let mut body = String::new();

    // Open comment ranges before body content
    for comment in comments {
        let id = xml_escape_attr(&comment.id);
        let _ = write!(body, "<w:commentRangeStart w:id=\"{id}\"/>");
    }

    for block in blocks {
        write_block(&mut body, block, rels, rel_counter);
    }

    // Close comment ranges after body content
    for comment in comments {
        let id = xml_escape_attr(&comment.id);
        let _ = write!(body, "<w:commentRangeEnd w:id=\"{id}\"/>");
    }

    // Emit comment references in a paragraph (required for Word to associate comments)
    if !comments.is_empty() {
        body.push_str("<w:p>");
        for comment in comments {
            let id = xml_escape_attr(&comment.id);
            let _ = write!(body, "<w:r><w:commentReference w:id=\"{id}\"/></w:r>");
        }
        body.push_str("</w:p>");
    }

    // Ensure at least one paragraph (required by Word)
    if blocks.is_empty() && comments.is_empty() {
        body.push_str("<w:p/>");
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <w:document xmlns:w=\"{NS_W}\" xmlns:r=\"{NS_R}\">\
         <w:body>{body}<w:sectPr/></w:body>\
         </w:document>"
    )
}

fn write_block(out: &mut String, block: &Block, rels: &mut Vec<RelEntry>, rel_counter: &mut usize) {
    match block {
        Block::Paragraph(para) => write_paragraph(out, para, rels, rel_counter),
        Block::Table(table) => write_table(out, table, rels, rel_counter),
        Block::Separator => write_separator(out),
    }
}

fn write_paragraph(
    out: &mut String,
    para: &Paragraph,
    rels: &mut Vec<RelEntry>,
    rel_counter: &mut usize,
) {
    out.push_str("<w:p>");
    for inline in &para.inlines {
        write_inline(out, inline, rels, rel_counter);
    }
    out.push_str("</w:p>");
}

fn write_inline(
    out: &mut String,
    inline: &Inline,
    rels: &mut Vec<RelEntry>,
    rel_counter: &mut usize,
) {
    match inline {
        Inline::Text(text) => {
            let escaped = xml_escape_text(text);
            // xml:space="preserve" keeps leading/trailing whitespace
            let _ = write!(
                out,
                "<w:r><w:t xml:space=\"preserve\">{escaped}</w:t></w:r>"
            );
        }
        Inline::Link(link) => {
            let rid = format!("rId{rel_counter}");
            *rel_counter += 1;
            rels.push(RelEntry {
                id: rid.clone(),
                rel_type: REL_TYPE_HYPERLINK.to_string(),
                target: link.target.clone(),
                target_mode: Some("External".to_string()),
            });
            let display = xml_escape_text(&link.display);
            let _ = write!(
                out,
                "<w:hyperlink r:id=\"{rid}\">\
                 <w:r>\
                 <w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr>\
                 <w:t xml:space=\"preserve\">{display}</w:t>\
                 </w:r>\
                 </w:hyperlink>"
            );
        }
    }
}

fn write_table(out: &mut String, table: &Table, rels: &mut Vec<RelEntry>, rel_counter: &mut usize) {
    // Optional caption as a paragraph before the table
    if let Some(caption) = &table.caption {
        let escaped = xml_escape_text(caption);
        let _ = write!(
            out,
            "<w:p><w:pPr><w:pStyle w:val=\"Caption\"/></w:pPr>\
             <w:r><w:t xml:space=\"preserve\">{escaped}</w:t></w:r></w:p>"
        );
    }

    out.push_str(
        "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/>\
                  <w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>",
    );

    // Grid columns — fall back to max row width if headers are empty
    let col_count = if table.headers.is_empty() {
        table.rows.iter().map(Vec::len).max().unwrap_or(0)
    } else {
        table.headers.len()
    };
    if col_count > 0 {
        out.push_str("<w:tblGrid>");
        for _ in 0..col_count {
            out.push_str("<w:gridCol/>");
        }
        out.push_str("</w:tblGrid>");
    }

    // Header row (unless synthetic headers that shouldn't be written)
    if !table.synthetic_headers {
        out.push_str("<w:tr>");
        for header in &table.headers {
            let escaped = xml_escape_text(header);
            let _ = write!(
                out,
                "<w:tc><w:p>\
                 <w:r><w:rPr><w:b/></w:rPr>\
                 <w:t xml:space=\"preserve\">{escaped}</w:t></w:r></w:p></w:tc>"
            );
        }
        out.push_str("</w:tr>");
    }

    // Data rows
    for row in &table.rows {
        out.push_str("<w:tr>");
        for cell in row {
            out.push_str("<w:tc>");
            if cell.content.is_empty() {
                out.push_str("<w:p/>");
            } else {
                for para in &cell.content {
                    write_paragraph(out, para, rels, rel_counter);
                }
            }
            out.push_str("</w:tc>");
        }
        // Pad short rows
        for _ in row.len()..col_count {
            out.push_str("<w:tc><w:p/></w:tc>");
        }
        out.push_str("</w:tr>");
    }

    out.push_str("</w:tbl>");
}

fn write_separator(out: &mut String) {
    out.push_str(
        "<w:p><w:pPr><w:pBdr>\
         <w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/>\
         </w:pBdr></w:pPr></w:p>",
    );
}

fn build_comments_xml(comments: &[CommentNote]) -> String {
    let mut body = String::new();
    for comment in comments {
        let id = xml_escape_attr(&comment.id);
        let author = xml_escape_attr(&comment.author);
        let text = xml_escape_text(&comment.text);
        let _ = write!(
            body,
            "<w:comment w:id=\"{id}\" w:author=\"{author}\">\
             <w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>\
             </w:comment>"
        );
    }
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <w:comments xmlns:w=\"{NS_W}\">{body}</w:comments>"
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use officemd_core::ir::{
        Block, DocSection, DocumentKind, Hyperlink, Inline, OoxmlDocument, Paragraph, Table,
        TableCell,
    };

    fn simple_para(text: &str) -> Paragraph {
        Paragraph {
            inlines: vec![Inline::Text(text.to_string())],
        }
    }

    fn simple_doc(blocks: Vec<Block>) -> OoxmlDocument {
        OoxmlDocument {
            kind: DocumentKind::Docx,
            properties: None,
            sheets: vec![],
            slides: vec![],
            sections: vec![DocSection {
                name: "body".to_string(),
                blocks,
                comments: vec![],
            }],
            pdf: None,
        }
    }

    #[test]
    fn generates_valid_docx_with_paragraph() {
        let doc = simple_doc(vec![Block::Paragraph(simple_para("Hello World"))]);
        let bytes = generate_docx(&doc).expect("generate");
        assert!(!bytes.is_empty());

        // Verify it's a valid ZIP that extracts back
        let ir = crate::extract_ir(&bytes).expect("extract");
        let body = ir.sections.iter().find(|s| s.name == "body").unwrap();
        assert!(!body.blocks.is_empty());
        // The extracted text should contain "Hello World"
        if let Block::Paragraph(p) = &body.blocks[0] {
            let text: String = p
                .inlines
                .iter()
                .map(|i| match i {
                    Inline::Text(t) => t.as_str(),
                    Inline::Link(l) => l.display.as_str(),
                })
                .collect();
            assert!(text.contains("Hello World"));
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn generates_docx_with_hyperlink() {
        let doc = simple_doc(vec![Block::Paragraph(Paragraph {
            inlines: vec![
                Inline::Text("Visit ".to_string()),
                Inline::Link(Hyperlink {
                    display: "Example".to_string(),
                    target: "https://example.com".to_string(),
                    rel_id: None,
                }),
            ],
        })]);
        let bytes = generate_docx(&doc).expect("generate");
        let ir = crate::extract_ir(&bytes).expect("extract");
        let body = ir.sections.iter().find(|s| s.name == "body").unwrap();
        // Should have the hyperlink in some form
        let has_link = body.blocks.iter().any(|b| {
            if let Block::Paragraph(p) = b {
                p.inlines.iter().any(|i| matches!(i, Inline::Link(_)))
            } else {
                false
            }
        });
        assert!(has_link, "hyperlink should survive round-trip");
    }

    #[test]
    fn generates_docx_with_table() {
        let doc = simple_doc(vec![Block::Table(Table {
            caption: Some("Test Table".to_string()),
            headers: vec!["Name".to_string(), "Value".to_string()],
            rows: vec![vec![
                TableCell {
                    content: vec![simple_para("Alice")],
                },
                TableCell {
                    content: vec![simple_para("100")],
                },
            ]],
            synthetic_headers: false,
        })]);
        let bytes = generate_docx(&doc).expect("generate");
        let ir = crate::extract_ir(&bytes).expect("extract");
        let body = ir.sections.iter().find(|s| s.name == "body").unwrap();
        // Should have a table block
        let has_table = body.blocks.iter().any(|b| matches!(b, Block::Table(_)));
        assert!(has_table, "table should survive round-trip");
    }

    #[test]
    fn generates_docx_with_separator() {
        let doc = simple_doc(vec![
            Block::Paragraph(simple_para("Before")),
            Block::Separator,
            Block::Paragraph(simple_para("After")),
        ]);
        let bytes = generate_docx(&doc).expect("generate");
        // Just verify it's a valid DOCX
        let ir = crate::extract_ir(&bytes).expect("extract");
        assert!(ir.sections.iter().any(|s| s.name == "body"));
    }

    #[test]
    fn generates_docx_with_comments() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            properties: None,
            sheets: vec![],
            slides: vec![],
            sections: vec![DocSection {
                name: "body".to_string(),
                blocks: vec![Block::Paragraph(simple_para("Text with comment"))],
                comments: vec![CommentNote {
                    id: "1".to_string(),
                    author: "Reviewer".to_string(),
                    text: "Looks good".to_string(),
                }],
            }],
            pdf: None,
        };
        let bytes = generate_docx(&doc).expect("generate");
        // Verify the ZIP contains comments.xml
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid package");
        assert!(pkg.has_part("word/comments.xml"));

        // Verify the document body has comment anchors
        let doc_xml = pkg.read_part_string("word/document.xml").unwrap().unwrap();
        assert!(
            doc_xml.contains("w:commentRangeStart"),
            "should have commentRangeStart"
        );
        assert!(
            doc_xml.contains("w:commentRangeEnd"),
            "should have commentRangeEnd"
        );
        assert!(
            doc_xml.contains("w:commentReference"),
            "should have commentReference"
        );

        // Verify round-trip: comments should survive extraction
        let ir = crate::extract_ir(&bytes).expect("extract");
        let body = ir.sections.iter().find(|s| s.name == "body").unwrap();
        assert_eq!(body.comments.len(), 1, "comment should survive round-trip");
        assert_eq!(body.comments[0].author, "Reviewer");
        assert_eq!(body.comments[0].text, "Looks good");
    }

    #[test]
    fn empty_document_produces_valid_docx() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            properties: None,
            sheets: vec![],
            slides: vec![],
            sections: vec![],
            pdf: None,
        };
        let bytes = generate_docx(&doc).expect("generate");
        // Should still be a valid DOCX (with empty body)
        let ir = crate::extract_ir(&bytes).expect("extract");
        assert_eq!(ir.kind, DocumentKind::Docx);
    }

    #[test]
    fn xml_special_chars_escaped_in_text() {
        let doc = simple_doc(vec![Block::Paragraph(simple_para("A & B < C > D"))]);
        let bytes = generate_docx(&doc).expect("generate");
        let ir = crate::extract_ir(&bytes).expect("extract");
        let body = ir.sections.iter().find(|s| s.name == "body").unwrap();
        if let Block::Paragraph(p) = &body.blocks[0]
            && let Inline::Text(t) = &p.inlines[0]
        {
            assert!(t.contains("A & B < C > D") || t.contains("A &amp;"));
        }
    }

    #[test]
    fn multi_paragraph_table_cell() {
        let doc = simple_doc(vec![Block::Table(Table {
            caption: None,
            headers: vec!["Col".to_string()],
            rows: vec![vec![TableCell {
                content: vec![simple_para("Line 1"), simple_para("Line 2")],
            }]],
            synthetic_headers: false,
        })]);
        let bytes = generate_docx(&doc).expect("generate");
        let ir = crate::extract_ir(&bytes).expect("extract");
        let body = ir.sections.iter().find(|s| s.name == "body").unwrap();
        let has_table = body.blocks.iter().any(|b| matches!(b, Block::Table(_)));
        assert!(has_table);
    }

    #[test]
    fn hyperlink_url_with_ampersand() {
        let doc = simple_doc(vec![Block::Paragraph(Paragraph {
            inlines: vec![Inline::Link(Hyperlink {
                display: "Search".to_string(),
                target: "https://example.com?a=1&b=2".to_string(),
                rel_id: None,
            })],
        })]);
        let bytes = generate_docx(&doc).expect("generate");
        // Should produce valid XML despite & in URL
        let ir = crate::extract_ir(&bytes).expect("extract");
        let body = ir.sections.iter().find(|s| s.name == "body").unwrap();
        assert!(!body.blocks.is_empty());
    }

    #[test]
    fn synthetic_headers_not_written_as_row() {
        let doc = simple_doc(vec![Block::Table(Table {
            caption: None,
            headers: vec!["Col1".to_string(), "Col2".to_string()],
            rows: vec![vec![
                TableCell {
                    content: vec![simple_para("A")],
                },
                TableCell {
                    content: vec![simple_para("B")],
                },
            ]],
            synthetic_headers: true,
        })]);
        let bytes = generate_docx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let xml = pkg.read_part_string("word/document.xml").unwrap().unwrap();
        // Synthetic headers should NOT appear as text in the document
        assert!(!xml.contains("Col1"), "synthetic headers should be skipped");
    }

    #[test]
    fn short_rows_padded_with_empty_cells() {
        let doc = simple_doc(vec![Block::Table(Table {
            caption: None,
            headers: vec!["A".to_string(), "B".to_string(), "C".to_string()],
            rows: vec![vec![TableCell {
                content: vec![simple_para("only one")],
            }]],
            synthetic_headers: false,
        })]);
        let bytes = generate_docx(&doc).expect("generate");
        // Should not crash; verify round-trip
        let ir = crate::extract_ir(&bytes).expect("extract");
        assert!(ir.sections.iter().any(|s| s.name == "body"));
    }

    #[test]
    fn table_with_no_headers_uses_row_width() {
        let doc = simple_doc(vec![Block::Table(Table {
            caption: None,
            headers: vec![],
            rows: vec![vec![
                TableCell {
                    content: vec![simple_para("X")],
                },
                TableCell {
                    content: vec![simple_para("Y")],
                },
            ]],
            synthetic_headers: true,
        })]);
        let bytes = generate_docx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let xml = pkg.read_part_string("word/document.xml").unwrap().unwrap();
        // Should have tblGrid with 2 gridCol elements
        assert!(xml.contains("<w:tblGrid>"));
        assert!(xml.contains("<w:gridCol/>"));
    }
}
