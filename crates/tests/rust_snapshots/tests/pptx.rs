use officemd_markdown::RenderOptions;
use officemd_snapshot_tests::{canonical_json, fixtures, normalize_markdown};

#[test]
fn showcase_pptx_ir() {
    let json = officemd_pptx::extract_ir_json(fixtures::SHOWCASE_PPTX).expect("extract PPTX IR");
    insta::assert_snapshot!("showcase_pptx_ir", canonical_json(&json));
}

#[test]
fn showcase_pptx_markdown() {
    let md = officemd_pptx::markdown_from_bytes_with_options(
        fixtures::SHOWCASE_PPTX,
        RenderOptions::default(),
    )
    .expect("render PPTX markdown");
    insta::assert_snapshot!("showcase_pptx_markdown", normalize_markdown(&md));
}
