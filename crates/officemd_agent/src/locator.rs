use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum ArtifactLocator {
    DocxParagraph {
        part: DocxPartLocator,
        paragraph_index: u32,
    },
    DocxTableCell {
        part: DocxPartLocator,
        table_index: u32,
        row_index: u32,
        column_index: u32,
    },
    XlsxCell {
        sheet: String,
        address: String,
    },
    XlsxSheet {
        name: String,
    },
    PptxShape {
        slide_number: u32,
        shape_id: u32,
    },
    PptxSlide {
        slide_number: u32,
    },
    PdfPage {
        page_number: u32,
    },
    PdfRegion {
        page_number: u32,
        bounds: PdfBounds,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct DocxPartLocator {
    pub part: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct XlsxCellLocator {
    pub sheet: String,
    pub address: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct XlsxSheetLocator {
    pub name: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct PdfBounds {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}
