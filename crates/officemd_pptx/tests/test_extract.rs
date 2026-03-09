use std::io::Write;

use officemd_core::ir::{Block, DocumentKind, Inline};
use officemd_pptx::{extract_ir, extract_ir_json, markdown_from_bytes};
use zip::ZipWriter;
use zip::write::FileOptions;

fn build_pptx(parts: Vec<(&str, &str)>) -> Vec<u8> {
    let mut buffer = Vec::new();
    let mut writer = ZipWriter::new(std::io::Cursor::new(&mut buffer));
    let options: FileOptions<'_, ()> = FileOptions::default();

    for (path, contents) in parts {
        writer.start_file(path, options).expect("start file");
        writer
            .write_all(contents.as_bytes())
            .expect("write contents");
    }

    writer.finish().expect("finish zip");
    buffer
}

fn minimal_pptx() -> Vec<u8> {
    let presentation = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>
"#;

    let presentation_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    Target="slides/slide1.xml"/>
</Relationships>
"#;

    let slide = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:txBody>
          <a:p><a:r><a:t>Title One</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Body line</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
"#;

    build_pptx(vec![
        ("ppt/presentation.xml", presentation),
        ("ppt/_rels/presentation.xml.rels", presentation_rels),
        ("ppt/slides/slide1.xml", slide),
    ])
}

fn collect_slide_text(blocks: &[Block]) -> String {
    let mut text = String::new();
    for block in blocks {
        if let Block::Paragraph(paragraph) = block {
            for inline in &paragraph.inlines {
                match inline {
                    Inline::Text(value) => text.push_str(value),
                    Inline::Link(link) => text.push_str(&link.display),
                }
            }
        }
    }
    text
}

#[test]
fn extracts_slides_from_synthetic_package() {
    let content = minimal_pptx();
    let doc = extract_ir(&content).expect("extract IR");

    assert_eq!(doc.kind, DocumentKind::Pptx);
    assert_eq!(doc.slides.len(), 1);
    assert_eq!(doc.slides[0].number, 1);
    assert_eq!(doc.slides[0].title.as_deref(), Some("Title One"));
    assert!(collect_slide_text(&doc.slides[0].blocks).contains("Body line"));
}

#[test]
fn extract_ir_json_returns_valid_json() {
    let content = minimal_pptx();
    let json = extract_ir_json(&content).expect("extract IR JSON");
    let parsed: serde_json::Value = serde_json::from_str(&json).expect("parse JSON");

    assert_eq!(parsed["kind"], "Pptx");
    assert!(parsed["slides"].is_array());
}

#[test]
fn renders_markdown_without_duplicate_title_line() {
    let content = minimal_pptx();
    let markdown = markdown_from_bytes(&content).expect("render markdown");

    assert!(markdown.contains("## Slide 1 - Title One"));
    assert_eq!(markdown.matches("Title One").count(), 1);
    assert!(markdown.contains("Body line"));
}
