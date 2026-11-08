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

const W_NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

fn cell(text: &str) -> String {
    format!("<w:tc><w:tcPr/><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>")
}

fn row(cells: &[&str]) -> String {
    let cells: String = cells.iter().map(|text| cell(text)).collect();
    format!("<w:tr>{cells}</w:tr>")
}

fn table(cols: usize, rows: &str) -> String {
    let grid = r#"<w:gridCol w:w="900"/>"#.repeat(cols);
    format!("<w:tbl><w:tblPr/><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>")
}

/// One row whose only cell spans `span` columns and holds `inner` before an
/// empty paragraph, as Word writes a cell that ends with a table.
fn spanning_row(span: usize, inner: &str) -> String {
    format!(
        r#"<w:tr><w:tc><w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr>{inner}<w:p/></w:tc></w:tr>"#
    )
}

fn docx_with_body(body: &str) -> Vec<u8> {
    let document_xml =
        format!(r#"<w:document {W_NS}><w:body>{body}<w:sectPr/></w:body></w:document>"#);
    build_docx(vec![("word/document.xml", &document_xml)])
}

fn table_count(blocks: &[Block]) -> usize {
    blocks
        .iter()
        .filter(|block| matches!(block, Block::Table(_)))
        .count()
}

fn nested_itinerary() -> String {
    table(
        3,
        &(row(&["Day", "Flight", "Hotel"]) + &row(&["Mon", "Gate 4", "Harbor Inn"])),
    )
}

fn assert_once_in_order(markdown: &str, texts: &[&str]) {
    let mut last = 0;
    for text in texts {
        assert_eq!(markdown.matches(text).count(), 1, "{text} in {markdown}");
        let at = markdown.find(text).expect("present");
        assert!(at >= last, "{text} out of reading order in {markdown}");
        last = at;
    }
}

#[test]
fn nested_table_keeps_every_outer_row() {
    let body = table(
        4,
        &(row(&["Owner", "Alex Fernsby", "Desk", "Room 4"])
            + &row(&["Backup", "Jordan Pell", "Desk", "Room 5"])
            + &spanning_row(4, &nested_itinerary())
            + &row(&["Signed", "Quinn Vale", "Date", "09/12"])),
    );
    let content = docx_with_body(&body);

    let doc = extract_ir(&content).expect("extract");
    let blocks = &doc.sections[0].blocks;
    assert_eq!(table_count(blocks), 1, "nested text stays inside the outer table");
    let Block::Table(outer) = &blocks[0] else {
        panic!("expected the outer table");
    };
    assert_eq!(outer.rows.len(), 4);

    let markdown = markdown_from_bytes(&content).expect("markdown");
    assert_once_in_order(
        &markdown,
        &["Alex Fernsby", "Jordan Pell", "Gate 4", "Harbor Inn", "Quinn Vale"],
    );
    assert!(
        markdown.contains("| Day<br>Flight<br>Hotel<br>Mon<br>Gate 4<br>Harbor Inn<br> |"),
        "nested cells fold into the spanning cell: {markdown}"
    );
}

#[test]
fn nested_table_in_the_first_row_keeps_later_rows() {
    let first = format!(
        "<w:tr><w:tc>{}<w:p/></w:tc>{}</w:tr>",
        nested_itinerary(),
        cell("Casey Morrow")
    );
    let body = table(2, &(first + &row(&["Signed", "Quinn Vale"])));
    let content = docx_with_body(&body);

    let doc = extract_ir(&content).expect("extract");
    assert_eq!(table_count(&doc.sections[0].blocks), 1);
    let markdown = markdown_from_bytes(&content).expect("markdown");

    assert_once_in_order(&markdown, &["Harbor Inn", "Casey Morrow", "Quinn Vale"]);
}

#[test]
fn tables_nested_three_deep_keep_every_level() {
    let deepest = table(1, &row(&["Riley Ash"]));
    let middle = table(
        2,
        &(row(&["Morgan Hale", "Desk"])
            + &spanning_row(2, &deepest)
            + &row(&["Sam Drew", "Desk"])),
    );
    let body = table(
        2,
        &(row(&["Owner", "Alex Fernsby"])
            + &spanning_row(2, &middle)
            + &row(&["Signed", "Quinn Vale"])),
    );
    let content = docx_with_body(&body);

    let doc = extract_ir(&content).expect("extract");
    assert_eq!(table_count(&doc.sections[0].blocks), 1);
    let markdown = markdown_from_bytes(&content).expect("markdown");
    assert_once_in_order(
        &markdown,
        &["Alex Fernsby", "Morgan Hale", "Riley Ash", "Sam Drew", "Quinn Vale"],
    );
}

#[test]
fn nested_table_beside_a_vertical_merge_keeps_every_row() {
    let merged_start = r#"<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Team</w:t></w:r></w:p></w:tc>"#;
    let merged_continue = r#"<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"#;
    let body = table(
        2,
        &(format!("<w:tr>{merged_start}{}</w:tr>", cell("Alex Fernsby"))
            + &format!(
                "<w:tr>{merged_continue}<w:tc>{}<w:p/></w:tc></w:tr>",
                nested_itinerary()
            )
            + &format!("<w:tr>{merged_continue}{}</w:tr>", cell("Quinn Vale"))),
    );
    let markdown = markdown_from_bytes(&docx_with_body(&body)).expect("markdown");

    assert_once_in_order(&markdown, &["Team", "Alex Fernsby", "Harbor Inn", "Quinn Vale"]);
}

#[test]
fn table_after_a_nested_table_stays_a_separate_block() {
    let outer = table(
        2,
        &(row(&["Owner", "Alex Fernsby"]) + &spanning_row(2, &nested_itinerary())),
    );
    let after = table(1, &row(&["Quinn Vale"]));
    let doc = extract_ir(&docx_with_body(&(outer + &after))).expect("extract");

    assert_eq!(table_count(&doc.sections[0].blocks), 2);
}

#[test]
fn text_after_a_text_box_stays_in_its_paragraph() {
    let body = r#"<w:p><w:r><w:t>Before Alex Fernsby</w:t></w:r><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed Jordan Pell</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:t>After Quinn Vale</w:t></w:r></w:p><w:p><w:r><w:t>Next Casey Morrow</w:t></w:r></w:p>"#;
    let content = docx_with_body(body);

    let doc = extract_ir(&content).expect("extract");
    assert_eq!(
        collect_paragraph_texts(&doc.sections[0].blocks),
        [
            "Before Alex Fernsby",
            "Boxed Jordan Pell",
            "After Quinn Vale",
            "Next Casey Morrow"
        ]
    );
}

#[test]
fn a_text_box_in_a_table_cell_keeps_the_cell_text_after_it() {
    let boxed = r#"<w:tc><w:p><w:r><w:t>Owner</w:t></w:r><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed Jordan Pell</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:t>Alex Fernsby</w:t></w:r></w:p></w:tc>"#;
    let body = table(2, &(format!("<w:tr>{boxed}{}</w:tr>", cell("Desk")) + &row(&["Signed", "Quinn Vale"])));
    let markdown = markdown_from_bytes(&docx_with_body(&body)).expect("markdown");

    assert_once_in_order(&markdown, &["Owner", "Boxed Jordan Pell", "Alex Fernsby", "Quinn Vale"]);
}
