use officemd_markdown::RenderOptions;
use officemd_snapshot_tests::{canonical_json, fixtures, normalize_markdown};

#[test]
fn openxml_whitepaper_ir() {
    let json =
        officemd_pdf::extract_ir_json(fixtures::OPENXML_WHITEPAPER_PDF).expect("extract PDF IR");
    insta::assert_snapshot!("openxml_whitepaper_ir", canonical_json(&json));
}

#[test]
fn openxml_whitepaper_markdown() {
    let md = officemd_pdf::markdown_from_bytes_with_options(
        fixtures::OPENXML_WHITEPAPER_PDF,
        RenderOptions::default(),
    )
    .expect("render PDF markdown");
    insta::assert_snapshot!("openxml_whitepaper_markdown", normalize_markdown(&md));
}

#[test]
fn ocr_graph_ocred_ir() {
    let json =
        officemd_pdf::extract_ir_json(fixtures::OCR_GRAPH_OCRED_PDF).expect("extract PDF IR");
    insta::assert_snapshot!("ocr_graph_ocred_ir", canonical_json(&json));
}

#[test]
fn ocr_graph_ocred_markdown() {
    let md = officemd_pdf::markdown_from_bytes_with_options(
        fixtures::OCR_GRAPH_OCRED_PDF,
        RenderOptions::default(),
    )
    .expect("render PDF markdown");
    insta::assert_snapshot!("ocr_graph_ocred_markdown", normalize_markdown(&md));
}

#[test]
fn ocr_graph_scanned_ir() {
    let json =
        officemd_pdf::extract_ir_json(fixtures::OCR_GRAPH_SCANNED_PDF).expect("extract PDF IR");
    insta::assert_snapshot!("ocr_graph_scanned_ir", canonical_json(&json));
}

#[test]
fn ocr_graph_scanned_markdown() {
    let md = officemd_pdf::markdown_from_bytes_with_options(
        fixtures::OCR_GRAPH_SCANNED_PDF,
        RenderOptions::default(),
    )
    .expect("render PDF markdown");
    insta::assert_snapshot!("ocr_graph_scanned_markdown", normalize_markdown(&md));
}

#[test]
fn ocr_tagged_textbased_ir() {
    let json =
        officemd_pdf::extract_ir_json(fixtures::OCR_TAGGED_TEXTBASED_PDF).expect("extract PDF IR");
    insta::assert_snapshot!("ocr_tagged_textbased_ir", canonical_json(&json));
}

#[test]
fn ocr_tagged_textbased_markdown() {
    let md = officemd_pdf::markdown_from_bytes_with_options(
        fixtures::OCR_TAGGED_TEXTBASED_PDF,
        RenderOptions::default(),
    )
    .expect("render PDF markdown");
    insta::assert_snapshot!("ocr_tagged_textbased_markdown", normalize_markdown(&md));
}

#[test]
fn encoding_heuristic_ir() {
    let json =
        officemd_pdf::extract_ir_json(fixtures::ENCODING_HEURISTIC_PDF).expect("extract PDF IR");
    insta::assert_snapshot!("encoding_heuristic_ir", canonical_json(&json));
}

#[test]
fn encoding_heuristic_markdown() {
    let md = officemd_pdf::markdown_from_bytes_with_options(
        fixtures::ENCODING_HEURISTIC_PDF,
        RenderOptions::default(),
    )
    .expect("render PDF markdown");
    insta::assert_snapshot!("encoding_heuristic_markdown", normalize_markdown(&md));
}
