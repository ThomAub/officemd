//! Core types and helpers for OOXML processing.
//!
//! Provides shared IR and relationship parsing used by format-specific crates.

pub mod format;
pub mod ir;
pub mod opc;
pub mod patch;
pub mod rels;

#[cfg(any(test, feature = "test-helpers"))]
#[doc(hidden)]
pub mod test_helpers;

pub use ir::*;
pub use opc::*;
pub use patch::*;
pub use rels::*;

#[cfg(test)]
mod tests {
    use super::ir::*;
    use super::rels::*;

    #[test]
    fn ir_document_defaults() {
        let doc = OoxmlDocument::default();
        assert_eq!(doc.kind, DocumentKind::Xlsx);
        assert!(doc.sheets.is_empty());
        assert!(doc.slides.is_empty());
        assert!(doc.sections.is_empty());
        assert!(doc.pdf.is_none());
        assert!(doc.properties.is_none());
    }

    #[test]
    fn ir_document_kinds() {
        assert_eq!(DocumentKind::default(), DocumentKind::Xlsx);

        let xlsx = OoxmlDocument {
            kind: DocumentKind::Xlsx,
            ..Default::default()
        };
        let docx = OoxmlDocument {
            kind: DocumentKind::Docx,
            ..Default::default()
        };
        let pptx = OoxmlDocument {
            kind: DocumentKind::Pptx,
            ..Default::default()
        };
        let pdf = OoxmlDocument {
            kind: DocumentKind::Pdf,
            ..Default::default()
        };

        assert_eq!(xlsx.kind, DocumentKind::Xlsx);
        assert_eq!(docx.kind, DocumentKind::Docx);
        assert_eq!(pptx.kind, DocumentKind::Pptx);
        assert_eq!(pdf.kind, DocumentKind::Pdf);
    }

    #[test]
    fn ir_sheet_construction() {
        let sheet = Sheet {
            name: "Data".into(),
            tables: vec![Table {
                caption: Some("Sales".into()),
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
                cell_ref: "A1".into(),
                formula: "=SUM(B1:B10)".into(),
            }],
            hyperlinks: vec![Hyperlink {
                display: "Link".into(),
                target: "https://example.com".into(),
                rel_id: None,
            }],
        };

        assert_eq!(sheet.name, "Data");
        assert_eq!(sheet.tables.len(), 1);
        assert_eq!(sheet.tables[0].headers.len(), 2);
        assert_eq!(sheet.formulas.len(), 1);
        assert_eq!(sheet.hyperlinks.len(), 1);
    }

    #[test]
    fn ir_slide_construction() {
        let slide = Slide {
            number: 1,
            title: Some("Introduction".into()),
            blocks: vec![
                Block::Paragraph(Paragraph {
                    inlines: vec![Inline::Text("Hello".into())],
                }),
                Block::Separator,
            ],
            notes: Some(vec![Paragraph {
                inlines: vec![Inline::Text("Speaker notes".into())],
            }]),
            comments: vec![CommentNote {
                id: "c1".into(),
                author: "Alice".into(),
                text: "Review this".into(),
            }],
        };

        assert_eq!(slide.number, 1);
        assert_eq!(slide.title, Some("Introduction".into()));
        assert_eq!(slide.blocks.len(), 2);
        assert!(slide.notes.is_some());
        assert_eq!(slide.comments.len(), 1);
    }

    #[test]
    fn ir_doc_section_construction() {
        let section = DocSection {
            name: "body".into(),
            blocks: vec![Block::Paragraph(Paragraph {
                inlines: vec![Inline::Text("Content".into())],
            })],
            comments: vec![],
        };

        assert_eq!(section.name, "body");
        assert_eq!(section.blocks.len(), 1);
    }

    #[test]
    fn ir_inline_variants() {
        let text = Inline::Text("Hello".into());
        let link = Inline::Link(Hyperlink {
            display: "Example".into(),
            target: "https://example.com".into(),
            rel_id: Some("rId1".into()),
        });

        match text {
            Inline::Text(t) => assert_eq!(t, "Hello"),
            Inline::Link(_) => panic!("Expected Text"),
        }

        match link {
            Inline::Link(h) => {
                assert_eq!(h.display, "Example");
                assert_eq!(h.target, "https://example.com");
                assert_eq!(h.rel_id, Some("rId1".into()));
            }
            Inline::Text(_) => panic!("Expected Link"),
        }
    }

    #[test]
    fn ir_serialization_roundtrip() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pptx,
            slides: vec![Slide {
                number: 1,
                title: Some("Test".into()),
                blocks: vec![],
                notes: None,
                comments: vec![],
            }],
            ..Default::default()
        };

        let json = serde_json::to_string(&doc).unwrap();
        let parsed: OoxmlDocument = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.kind, DocumentKind::Pptx);
        assert_eq!(parsed.slides.len(), 1);
        assert_eq!(parsed.slides[0].title, Some("Test".into()));
    }

    #[test]
    fn ir_pdf_serialization_roundtrip() {
        let doc = OoxmlDocument {
            kind: DocumentKind::Pdf,
            pdf: Some(PdfDocument {
                pages: vec![PdfPage {
                    number: 1,
                    markdown: "Hello".into(),
                }],
                diagnostics: PdfDiagnostics {
                    classification: PdfClassification::TextBased,
                    confidence: 0.95,
                    page_count: 1,
                    pages_needing_ocr: vec![],
                    has_encoding_issues: false,
                },
            }),
            ..Default::default()
        };

        let json = serde_json::to_string(&doc).unwrap();
        let parsed: OoxmlDocument = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.kind, DocumentKind::Pdf);
        let pdf = parsed.pdf.expect("pdf payload");
        assert_eq!(pdf.pages.len(), 1);
        assert_eq!(pdf.pages[0].number, 1);
        assert_eq!(pdf.pages[0].markdown, "Hello");
        assert_eq!(pdf.diagnostics.classification, PdfClassification::TextBased);
    }

    #[test]
    fn rels_parse_empty() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;

        let rels = parse_relationships(xml).unwrap();
        assert!(rels.is_empty());
    }

    #[test]
    fn rels_parse_single() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;

        let rels = parse_relationships(xml).unwrap();
        assert_eq!(rels.len(), 1);
        assert_eq!(rels[0].id, "rId1");
        assert_eq!(rels[0].target, "slides/slide1.xml");
        assert!(rels[0].rel_type.contains("slide"));
        assert!(rels[0].target_mode.is_none());
    }

    #[test]
    fn rels_parse_unescapes_attribute_values() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com?a=1&amp;b=2" TargetMode="External"/>
</Relationships>"#;

        let rels = parse_relationships(xml).unwrap();
        assert_eq!(rels.len(), 1);
        assert_eq!(rels[0].target, "https://example.com?a=1&b=2");
        assert_eq!(rels[0].target_mode.as_deref(), Some("External"));
    }

    #[test]
    fn rels_parse_multiple_with_target_mode() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

        let rels = parse_relationships(xml).unwrap();
        assert_eq!(rels.len(), 2);

        assert_eq!(rels[0].id, "rId1");
        assert!(rels[0].target_mode.is_none());

        assert_eq!(rels[1].id, "rId2");
        assert_eq!(rels[1].target, "https://example.com");
        assert_eq!(rels[1].target_mode, Some("External".into()));
    }

    #[test]
    fn rels_target_map_filtering() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#;

        let rels = parse_relationships(xml).unwrap();

        // No filter - all relationships
        let all_map = rel_target_map(&rels, None);
        assert_eq!(all_map.len(), 3);

        // Filter by worksheet type
        let worksheet_type =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        let worksheet_map = rel_target_map(&rels, Some(worksheet_type));
        assert_eq!(worksheet_map.len(), 2);
        assert!(worksheet_map.contains_key("rId1"));
        assert!(worksheet_map.contains_key("rId3"));
        assert!(!worksheet_map.contains_key("rId2"));
    }
}
