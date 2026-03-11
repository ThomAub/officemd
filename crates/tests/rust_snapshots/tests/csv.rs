use officemd_csv::markdown_from_bytes_with_options;
use officemd_markdown::RenderOptions;
use officemd_snapshot_tests::{canonical_json, fixtures, normalize_markdown};

#[test]
fn showcase_csv_ir() {
    let json =
        officemd_csv::extract_ir::extract_ir_json(fixtures::SHOWCASE_CSV).expect("extract CSV IR");
    insta::assert_snapshot!("showcase_csv_ir", canonical_json(&json));
}

#[test]
fn showcase_csv_markdown() {
    let md = markdown_from_bytes_with_options(fixtures::SHOWCASE_CSV, RenderOptions::default())
        .expect("render CSV markdown");
    insta::assert_snapshot!("showcase_csv_markdown", normalize_markdown(&md));
}
