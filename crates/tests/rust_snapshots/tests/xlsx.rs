use officemd_markdown::RenderOptions;
use officemd_snapshot_tests::{canonical_json, fixtures, normalize_markdown};

#[test]
fn showcase_xlsx_ir() {
    let json = officemd_xlsx::extract_ir::extract_ir_json(fixtures::SHOWCASE_XLSX)
        .expect("extract XLSX IR");
    insta::assert_snapshot!("showcase_xlsx_ir", canonical_json(&json));
}

#[test]
fn showcase_xlsx_markdown() {
    let md = officemd_xlsx::markdown_from_bytes_with_options(
        fixtures::SHOWCASE_XLSX,
        RenderOptions::default(),
    )
    .expect("render XLSX markdown");
    insta::assert_snapshot!("showcase_xlsx_markdown", normalize_markdown(&md));
}

#[test]
fn trim_sparse_trailing_xlsx_ir() {
    let json = officemd_xlsx::extract_ir::extract_ir_json(fixtures::TRIM_SPARSE_TRAILING_XLSX)
        .expect("extract XLSX IR");
    insta::assert_snapshot!("trim_sparse_trailing_xlsx_ir", canonical_json(&json));
}

#[test]
fn trim_sparse_trailing_xlsx_markdown() {
    let md = officemd_xlsx::markdown_from_bytes_with_options(
        fixtures::TRIM_SPARSE_TRAILING_XLSX,
        RenderOptions::default(),
    )
    .expect("render XLSX markdown");
    insta::assert_snapshot!(
        "trim_sparse_trailing_xlsx_markdown",
        normalize_markdown(&md)
    );
}

#[test]
fn trim_wide_sparse_xlsx_ir() {
    let json = officemd_xlsx::extract_ir::extract_ir_json(fixtures::TRIM_WIDE_SPARSE_XLSX)
        .expect("extract XLSX IR");
    insta::assert_snapshot!("trim_wide_sparse_xlsx_ir", canonical_json(&json));
}

#[test]
fn trim_wide_sparse_xlsx_markdown() {
    let md = officemd_xlsx::markdown_from_bytes_with_options(
        fixtures::TRIM_WIDE_SPARSE_XLSX,
        RenderOptions::default(),
    )
    .expect("render XLSX markdown");
    insta::assert_snapshot!("trim_wide_sparse_xlsx_markdown", normalize_markdown(&md));
}

#[test]
fn trim_all_empty_xlsx_ir() {
    let json = officemd_xlsx::extract_ir::extract_ir_json(fixtures::TRIM_ALL_EMPTY_XLSX)
        .expect("extract XLSX IR");
    insta::assert_snapshot!("trim_all_empty_xlsx_ir", canonical_json(&json));
}

#[test]
fn trim_all_empty_xlsx_markdown() {
    let md = officemd_xlsx::markdown_from_bytes_with_options(
        fixtures::TRIM_ALL_EMPTY_XLSX,
        RenderOptions::default(),
    )
    .expect("render XLSX markdown");
    insta::assert_snapshot!("trim_all_empty_xlsx_markdown", normalize_markdown(&md));
}
