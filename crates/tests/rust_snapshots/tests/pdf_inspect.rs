use officemd_snapshot_tests::fixtures;

fn inspect_json(content: &[u8]) -> String {
    let diagnostics = officemd_pdf::inspect_pdf(content).expect("inspect PDF");
    let value = serde_json::to_value(&diagnostics).expect("serialize");
    serde_json::to_string_pretty(&value).expect("pretty print")
}

#[test]
fn inspect_openxml_whitepaper() {
    insta::assert_snapshot!(
        "inspect_openxml_whitepaper",
        inspect_json(fixtures::OPENXML_WHITEPAPER_PDF)
    );
}

#[test]
fn inspect_ocr_graph_ocred() {
    insta::assert_snapshot!(
        "inspect_ocr_graph_ocred",
        inspect_json(fixtures::OCR_GRAPH_OCRED_PDF)
    );
}

#[test]
fn inspect_ocr_graph_scanned() {
    insta::assert_snapshot!(
        "inspect_ocr_graph_scanned",
        inspect_json(fixtures::OCR_GRAPH_SCANNED_PDF)
    );
}

#[test]
fn inspect_ocr_tagged_textbased() {
    insta::assert_snapshot!(
        "inspect_ocr_tagged_textbased",
        inspect_json(fixtures::OCR_TAGGED_TEXTBASED_PDF)
    );
}

#[test]
fn inspect_encoding_heuristic() {
    insta::assert_snapshot!(
        "inspect_encoding_heuristic",
        inspect_json(fixtures::ENCODING_HEURISTIC_PDF)
    );
}
