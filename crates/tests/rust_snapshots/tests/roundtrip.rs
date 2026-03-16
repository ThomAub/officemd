//! Integration round-trip tests: render(parse(render(ir))) == render(ir)
//!
//! These verify that real extracted documents survive a markdown round-trip.

use officemd_markdown::{MarkdownProfile, RenderOptions, parse_document};

/// Helper: render an IR with given options, parse it back, re-render, and assert equality.
fn assert_roundtrip(doc: &officemd_core::ir::OoxmlDocument, opts: RenderOptions) {
    let md1 = officemd_markdown::render_document_with_options(doc, opts);
    let parsed = parse_document(&md1).expect("parse_document should succeed");
    let md2 = officemd_markdown::render_document_with_options(&parsed, opts);
    assert_eq!(md1, md2, "render-parse-render stability failed");
}

// ---------------------------------------------------------------------------
// XLSX
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_showcase_xlsx_compact() {
    let ir =
        officemd_xlsx::extract_ir::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_XLSX)
            .expect("extract XLSX IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

#[test]
fn roundtrip_showcase_xlsx_human() {
    let ir =
        officemd_xlsx::extract_ir::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_XLSX)
            .expect("extract XLSX IR");
    assert_roundtrip(
        &ir,
        RenderOptions {
            markdown_profile: MarkdownProfile::Human,
            ..Default::default()
        },
    );
}

#[test]
fn roundtrip_trim_sparse_trailing_xlsx() {
    let ir = officemd_xlsx::extract_ir::extract_ir(
        officemd_snapshot_tests::fixtures::TRIM_SPARSE_TRAILING_XLSX,
    )
    .expect("extract XLSX IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

#[test]
fn roundtrip_trim_wide_sparse_xlsx() {
    let ir = officemd_xlsx::extract_ir::extract_ir(
        officemd_snapshot_tests::fixtures::TRIM_WIDE_SPARSE_XLSX,
    )
    .expect("extract XLSX IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

#[test]
fn roundtrip_trim_all_empty_xlsx() {
    let ir = officemd_xlsx::extract_ir::extract_ir(
        officemd_snapshot_tests::fixtures::TRIM_ALL_EMPTY_XLSX,
    )
    .expect("extract XLSX IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

// ---------------------------------------------------------------------------
// CSV
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_showcase_csv() {
    let ir = officemd_csv::extract_ir::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_CSV)
        .expect("extract CSV IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_showcase_docx_compact() {
    let ir = officemd_docx::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_DOCX)
        .expect("extract DOCX IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

#[test]
fn roundtrip_showcase_docx_human() {
    let ir = officemd_docx::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_DOCX)
        .expect("extract DOCX IR");
    assert_roundtrip(
        &ir,
        RenderOptions {
            markdown_profile: MarkdownProfile::Human,
            ..Default::default()
        },
    );
}

#[test]
fn roundtrip_showcase_02_docx() {
    let ir = officemd_docx::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_02_DOCX)
        .expect("extract DOCX IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

// ---------------------------------------------------------------------------
// PPTX
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_showcase_pptx_compact() {
    let ir = officemd_pptx::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_PPTX)
        .expect("extract PPTX IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

#[test]
fn roundtrip_showcase_pptx_human() {
    let ir = officemd_pptx::extract_ir(officemd_snapshot_tests::fixtures::SHOWCASE_PPTX)
        .expect("extract PPTX IR");
    assert_roundtrip(
        &ir,
        RenderOptions {
            markdown_profile: MarkdownProfile::Human,
            ..Default::default()
        },
    );
}

// ---------------------------------------------------------------------------
// PDF
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_openxml_whitepaper_pdf() {
    let ir = officemd_pdf::extract_ir(officemd_snapshot_tests::fixtures::OPENXML_WHITEPAPER_PDF)
        .expect("extract PDF IR");
    assert_roundtrip(&ir, RenderOptions::default());
}

#[test]
fn roundtrip_encoding_heuristic_pdf() {
    let ir = officemd_pdf::extract_ir(officemd_snapshot_tests::fixtures::ENCODING_HEURISTIC_PDF)
        .expect("extract PDF IR");
    assert_roundtrip(&ir, RenderOptions::default());
}
