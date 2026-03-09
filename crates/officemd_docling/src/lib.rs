//! Docling Document JSON renderer for OOXML IR.
//!
//! Translates the shared `OoxmlDocument` intermediate representation into
//! Docling Document JSON (v1.9.0).

pub mod builder;
pub mod convert;
pub mod convert_docx;
pub mod convert_pdf;
pub mod convert_pptx;
pub mod convert_xlsx;
pub mod model;

use officemd_core::ir::{DocumentKind, OoxmlDocument};

use builder::DoclingDocumentBuilder;
use model::DoclingDocument;

/// Convert an OOXML IR document to a Docling Document.
#[must_use]
pub fn convert_document(doc: &OoxmlDocument) -> DoclingDocument {
    let name = match doc.kind {
        DocumentKind::Xlsx => "spreadsheet",
        DocumentKind::Pptx => "presentation",
        DocumentKind::Docx | DocumentKind::Pdf => "document",
    };

    let mut builder = DoclingDocumentBuilder::new(name);

    match doc.kind {
        DocumentKind::Docx => convert_docx::convert_docx(doc, &mut builder),
        DocumentKind::Xlsx => convert_xlsx::convert_xlsx(doc, &mut builder),
        DocumentKind::Pptx => convert_pptx::convert_pptx(doc, &mut builder),
        DocumentKind::Pdf => convert_pdf::convert_pdf(doc, &mut builder),
    }

    builder.build()
}

/// Convert an OOXML IR document to Docling Document JSON string.
///
/// # Errors
///
/// Returns `serde_json::Error` if the document cannot be serialized to JSON.
pub fn convert_document_json(doc: &OoxmlDocument) -> Result<String, serde_json::Error> {
    let docling = convert_document(doc);
    serde_json::to_string(&docling)
}

#[cfg(test)]
mod tests {
    use super::*;
    use officemd_core::ir::{
        Block, CommentNote, DocSection, FormulaNote, Hyperlink, Inline, Paragraph, PdfDiagnostics,
        PdfDocument, PdfPage, Sheet, Slide, Table, TableCell,
    };

    #[test]
    fn converts_empty_docx() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![],
                comments: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.schema_name, "DoclingDocument");
        assert_eq!(result.version, "1.9.0");
        assert_eq!(result.name, "document");
        assert!(result.body.children.is_empty());
        assert!(result.texts.is_empty());
        assert!(result.tables.is_empty());
    }

    #[test]
    fn converts_docx_body_paragraph() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Hello World".into())],
                })],
                comments: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.texts.len(), 1);
        assert_eq!(result.texts[0].text, "Hello World");
        assert_eq!(result.texts[0].self_ref, "#/texts/0");
        assert_eq!(result.body.children.len(), 1);
        assert_eq!(result.body.children[0].cref, "#/texts/0");
    }

    #[test]
    fn converts_docx_header_to_furniture() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "header1".into(),
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Page Header".into())],
                })],
                comments: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.furniture.children.len(), 1);
        assert_eq!(result.texts[0].text, "Page Header");
    }

    #[test]
    fn converts_docx_footnotes_section_to_group() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "footnotes".into(),
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Note text".into())],
                })],
                comments: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.groups.len(), 1);
        assert_eq!(result.groups[0].name, "footnotes");
        assert_eq!(result.body.children.len(), 1);
    }

    #[test]
    fn converts_docx_comments() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![],
                comments: vec![CommentNote {
                    id: "c1".into(),
                    author: "Alice".into(),
                    text: "Review this".into(),
                }],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.groups.len(), 1);
        assert_eq!(result.groups[0].name, "comments");
        assert_eq!(result.texts.len(), 1);
        assert_eq!(result.texts[0].text, "Alice: Review this");
    }

    #[test]
    fn converts_docx_table_with_caption() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![Block::Table(Table {
                    caption: Some("Sales Data".into()),
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
                })],
                comments: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        // Caption + nothing else in texts
        assert_eq!(result.texts.len(), 1);
        assert_eq!(result.texts[0].text, "Sales Data");
        // One table
        assert_eq!(result.tables.len(), 1);
        assert_eq!(result.tables[0].data.num_rows, 2);
        assert_eq!(result.tables[0].data.num_cols, 2);
        assert_eq!(result.tables[0].data.table_cells.len(), 4);
        // First row is column header
        assert!(result.tables[0].data.table_cells[0].column_header);
        assert!(!result.tables[0].data.table_cells[2].column_header);
        // Caption reference
        assert_eq!(result.tables[0].captions.len(), 1);
    }

    #[test]
    fn converts_xlsx_sheets() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![Sheet {
                name: "Data".into(),
                tables: vec![Table {
                    caption: None,
                    headers: vec!["A".into()],
                    rows: vec![vec![TableCell {
                        content: vec![Paragraph {
                            inlines: vec![Inline::Text("1".into())],
                        }],
                    }]],
                }],
                formulas: vec![FormulaNote {
                    cell_ref: "B1".into(),
                    formula: "=A1+1".into(),
                }],
                hyperlinks: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        // One group for the sheet
        assert_eq!(result.groups.len(), 1);
        assert_eq!(result.groups[0].name, "Data");
        // One table + one formula text
        assert_eq!(result.tables.len(), 1);
        assert_eq!(result.texts.len(), 1);
        assert_eq!(result.texts[0].text, "B1 = =A1+1");
    }

    #[test]
    fn converts_pptx_slides() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![Slide {
                number: 1,
                title: Some("Intro".into()),
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Welcome".into())],
                })],
                notes: Some(vec![Paragraph {
                    inlines: vec![Inline::Text("Speaker note".into())],
                }]),
                comments: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.groups.len(), 1);
        assert_eq!(result.groups[0].name, "Slide 1 - Intro");
        assert_eq!(result.pages.len(), 1);
        assert!(result.pages.contains_key("1"));
        // Title + paragraph + note = 3 text items
        assert_eq!(result.texts.len(), 3);
        assert_eq!(result.texts[0].text, "Intro");
        assert_eq!(result.texts[1].text, "Welcome");
        assert_eq!(result.texts[2].text, "Speaker note");
    }

    #[test]
    fn converts_pdf_pages() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pdf,
            pdf: Some(PdfDocument {
                pages: vec![
                    PdfPage {
                        number: 1,
                        markdown: "Page one content".into(),
                    },
                    PdfPage {
                        number: 2,
                        markdown: "Page two content".into(),
                    },
                ],
                diagnostics: PdfDiagnostics::default(),
            }),
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.pages.len(), 2);
        assert_eq!(result.texts.len(), 2);
        assert_eq!(result.texts[0].text, "Page one content");
        assert_eq!(result.texts[1].text, "Page two content");
    }

    #[test]
    fn converts_hyperlinks_to_text() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![
                        Inline::Text("Visit ".into()),
                        Inline::Link(Hyperlink {
                            display: "Example".into(),
                            target: "https://example.com".into(),
                            rel_id: None,
                        }),
                    ],
                })],
                comments: vec![],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.texts[0].text, "Visit Example");
    }

    #[test]
    fn json_output_is_valid() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Docx,
            sections: vec![DocSection {
                name: "body".into(),
                blocks: vec![Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Test".into())],
                })],
                comments: vec![],
            }],
            ..Default::default()
        };
        let json = convert_document_json(&doc).expect("valid json");
        let parsed: serde_json::Value = serde_json::from_str(&json).expect("parseable json");
        assert_eq!(parsed["schema_name"], "DoclingDocument");
        assert_eq!(parsed["version"], "1.9.0");
        assert_eq!(parsed["texts"][0]["text"], "Test");
        assert_eq!(parsed["texts"][0]["self_ref"], "#/texts/0");
        assert_eq!(parsed["texts"][0]["parent"]["$ref"], "#/body");
    }

    #[test]
    fn xlsx_hyperlinks_become_text_items() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            sheets: vec![Sheet {
                name: "Sheet1".into(),
                tables: vec![],
                formulas: vec![],
                hyperlinks: vec![Hyperlink {
                    display: "Click here".into(),
                    target: "https://example.com".into(),
                    rel_id: None,
                }],
            }],
            ..Default::default()
        };
        let result = convert_document(&doc);
        assert_eq!(result.texts.len(), 1);
        assert_eq!(result.texts[0].text, "Click here (https://example.com)");
    }
}
