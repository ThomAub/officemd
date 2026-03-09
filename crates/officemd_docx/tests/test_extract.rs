use std::io::Write;

use officemd_core::ir::{Block, DocumentKind, Inline, Paragraph};
use officemd_docx::{extract_ir, markdown_from_bytes};
use zip::ZipWriter;
use zip::write::FileOptions;

fn build_docx(parts: Vec<(&str, &str)>) -> Vec<u8> {
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

fn minimal_docx_with_hyperlink(target_mode: &str) -> Vec<u8> {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t>Visit </w:t></w:r>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>Example</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"#;

    let document_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com"
    TargetMode="{target_mode}"/>
</Relationships>"#
    );

    build_docx(vec![
        ("word/document.xml", document_xml),
        ("word/_rels/document.xml.rels", &document_rels),
    ])
}

fn minimal_docx_with_comment_markers() -> Vec<u8> {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>Hello</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
    </w:p>
    <w:p>
      <w:r><w:t>After</w:t></w:r>
      <w:commentReference w:id="0"/>
    </w:p>
  </w:body>
</w:document>"#;

    let comments_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Alice">
    <w:p><w:r><w:t>Check this</w:t></w:r></w:p>
  </w:comment>
</w:comments>"#;

    build_docx(vec![
        ("word/document.xml", document_xml),
        ("word/comments.xml", comments_xml),
    ])
}

fn collect_links(blocks: &[Block]) -> Vec<(String, String)> {
    let mut links = Vec::new();
    for block in blocks {
        if let Block::Paragraph(paragraph) = block {
            for inline in &paragraph.inlines {
                if let Inline::Link(link) = inline {
                    links.push((link.display.clone(), link.target.clone()));
                }
            }
        }
    }
    links
}

fn collect_paragraph_texts(blocks: &[Block]) -> Vec<String> {
    blocks
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(render_paragraph_text(paragraph)),
            _ => None,
        })
        .collect()
}

fn render_paragraph_text(paragraph: &Paragraph) -> String {
    let mut text = String::new();
    for inline in &paragraph.inlines {
        match inline {
            Inline::Text(value) => text.push_str(value),
            Inline::Link(link) => text.push_str(&link.display),
        }
    }
    text
}

#[test]
fn extracts_hyperlink_with_case_insensitive_target_mode() {
    let bytes = minimal_docx_with_hyperlink("external");
    let doc = extract_ir(&bytes).expect("extract docx");

    assert_eq!(doc.kind, DocumentKind::Docx);
    let body = doc
        .sections
        .iter()
        .find(|section| section.name == "body")
        .expect("body section");

    let links = collect_links(&body.blocks);
    assert!(
        links
            .iter()
            .any(|(display, target)| display == "Example" && target == "https://example.com"),
        "links: {links:?}"
    );
}

#[test]
fn renders_markdown_with_docx_hyperlink() {
    let bytes = minimal_docx_with_hyperlink("External");
    let markdown = markdown_from_bytes(&bytes).expect("render markdown");

    assert!(markdown.contains("## Section: body"));
    assert!(markdown.contains("[Example](https://example.com)"));
}

#[test]
fn comment_range_start_adds_section_comment_without_inline_anchor() {
    let bytes = minimal_docx_with_comment_markers();
    let doc = extract_ir(&bytes).expect("extract docx");

    let body = doc
        .sections
        .iter()
        .find(|section| section.name == "body")
        .expect("body section");

    assert_eq!(body.comments.len(), 1);
    assert_eq!(body.comments[0].id, "c0");
    assert_eq!(body.comments[0].author, "Alice");
    assert_eq!(body.comments[0].text, "Check this");

    let paragraphs = collect_paragraph_texts(&body.blocks);
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[0], "Hello");
    assert_eq!(paragraphs[1], "After[^c0]");
    assert_eq!(
        paragraphs
            .iter()
            .map(|p| p.matches("[^c0]").count())
            .sum::<usize>(),
        1
    );
}
