use officemd_markdown::RenderOptions;
use officemd_snapshot_tests::{canonical_json, fixtures, normalize_markdown};

#[test]
fn showcase_docx_ir() {
    let json = officemd_docx::extract_ir_json(fixtures::SHOWCASE_DOCX).expect("extract DOCX IR");
    insta::assert_snapshot!("showcase_docx_ir", canonical_json(&json));
}

#[test]
fn showcase_docx_markdown() {
    let md = officemd_docx::markdown_from_bytes_with_options(
        fixtures::SHOWCASE_DOCX,
        RenderOptions::default(),
    )
    .expect("render DOCX markdown");
    insta::assert_snapshot!("showcase_docx_markdown", normalize_markdown(&md));
}

#[test]
fn showcase_02_docx_ir() {
    let json = officemd_docx::extract_ir_json(fixtures::SHOWCASE_02_DOCX).expect("extract DOCX IR");
    insta::assert_snapshot!("showcase_02_docx_ir", canonical_json(&json));
}

#[test]
fn showcase_02_docx_markdown() {
    let md = officemd_docx::markdown_from_bytes_with_options(
        fixtures::SHOWCASE_02_DOCX,
        RenderOptions::default(),
    )
    .expect("render DOCX markdown");
    insta::assert_snapshot!("showcase_02_docx_markdown", normalize_markdown(&md));
}
