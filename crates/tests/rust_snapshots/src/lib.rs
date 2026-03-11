//! Shared fixtures and helpers for snapshot tests.

/// All test fixtures loaded from `examples/data/`.
pub mod fixtures {
    pub static SHOWCASE_DOCX: &[u8] = include_bytes!("../../../../examples/data/showcase.docx");
    pub static SHOWCASE_02_DOCX: &[u8] =
        include_bytes!("../../../../examples/data/showcase_02.docx");
    pub static SHOWCASE_XLSX: &[u8] = include_bytes!("../../../../examples/data/showcase.xlsx");
    pub static TRIM_SPARSE_TRAILING_XLSX: &[u8] =
        include_bytes!("../../../../examples/data/trim_sparse_trailing.xlsx");
    pub static TRIM_WIDE_SPARSE_XLSX: &[u8] =
        include_bytes!("../../../../examples/data/trim_wide_sparse.xlsx");
    pub static TRIM_ALL_EMPTY_XLSX: &[u8] =
        include_bytes!("../../../../examples/data/trim_all_empty.xlsx");
    pub static SHOWCASE_CSV: &[u8] = include_bytes!("../../../../examples/data/showcase.csv");
    pub static SHOWCASE_PPTX: &[u8] = include_bytes!("../../../../examples/data/showcase.pptx");
    pub static OPENXML_WHITEPAPER_PDF: &[u8] =
        include_bytes!("../../../../examples/data/OpenXML_WhitePaper.pdf");
    pub static OCR_GRAPH_OCRED_PDF: &[u8] =
        include_bytes!("../../../../examples/data/ocr_graph_ocred.pdf");
    pub static OCR_GRAPH_SCANNED_PDF: &[u8] =
        include_bytes!("../../../../examples/data/ocr_graph_scanned.pdf");
    pub static OCR_TAGGED_TEXTBASED_PDF: &[u8] =
        include_bytes!("../../../../examples/data/ocr_tagged_textbased.pdf");
    pub static ENCODING_HEURISTIC_PDF: &[u8] =
        include_bytes!("../../../../examples/data/encoding_heuristic_fixture.pdf");
}

/// Canonical JSON formatting for stable snapshot comparisons.
pub fn canonical_json(payload: &str) -> String {
    let value: serde_json::Value = serde_json::from_str(payload).expect("valid JSON");
    serde_json::to_string_pretty(&value).expect("JSON serialization")
}

/// Normalize line endings for stable markdown snapshot comparisons.
pub fn normalize_markdown(payload: &str) -> String {
    payload.replace("\r\n", "\n")
}
