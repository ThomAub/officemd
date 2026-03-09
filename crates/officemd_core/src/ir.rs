//! Shared document intermediate representation (IR) across XLSX/DOCX/PPTX/PDF.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// Document modality.
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq, Eq)]
pub enum DocumentKind {
    #[default]
    Xlsx,
    Docx,
    Pptx,
    Pdf,
}

/// Document properties collected from docProps/*.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentProperties {
    pub core: HashMap<String, String>,
    pub app: HashMap<String, String>,
    pub custom: HashMap<String, String>,
}

/// Hyperlink with display text and target.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Hyperlink {
    pub display: String,
    pub target: String,
    pub rel_id: Option<String>,
}

/// Inline content within a paragraph.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Inline {
    Text(String),
    Link(Hyperlink),
}

impl Default for Inline {
    fn default() -> Self {
        Inline::Text(String::new())
    }
}

/// Logical paragraph.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Paragraph {
    pub inlines: Vec<Inline>,
}

/// Table cell content (multi-paragraph).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCell {
    pub content: Vec<Paragraph>,
}

/// Table with synthetic headers.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Table {
    pub caption: Option<String>,
    pub headers: Vec<String>,
    pub rows: Vec<Vec<TableCell>>,
}

/// Formula footnote entry.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct FormulaNote {
    pub cell_ref: String,
    pub formula: String,
}

/// Comment or note footnote entry.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CommentNote {
    pub id: String,
    pub author: String,
    pub text: String,
}

/// Block element for rendering.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Separator,
}

/// XLSX sheet.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    pub tables: Vec<Table>,
    pub formulas: Vec<FormulaNote>,
    pub hyperlinks: Vec<Hyperlink>,
}

/// PPTX slide.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Slide {
    pub number: usize,
    pub title: Option<String>,
    pub blocks: Vec<Block>,
    pub notes: Option<Vec<Paragraph>>,
    pub comments: Vec<CommentNote>,
}

/// DOCX section (body, header/footer, notes).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocSection {
    pub name: String,
    pub blocks: Vec<Block>,
    pub comments: Vec<CommentNote>,
}

/// Top-level document container.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct OoxmlDocument {
    pub kind: DocumentKind,
    pub properties: Option<DocumentProperties>,
    pub sheets: Vec<Sheet>,
    pub slides: Vec<Slide>,
    pub sections: Vec<DocSection>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pdf: Option<PdfDocument>,
}

/// PDF classification from pdf-inspector.
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq, Eq)]
pub enum PdfClassification {
    #[default]
    TextBased,
    Scanned,
    ImageBased,
    Mixed,
}

/// PDF parseability and OCR-routing diagnostics.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PdfDiagnostics {
    pub classification: PdfClassification,
    pub confidence: f32,
    pub page_count: usize,
    pub pages_needing_ocr: Vec<usize>,
    pub has_encoding_issues: bool,
}

/// Page-local markdown payload for a PDF page.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PdfPage {
    pub number: usize,
    pub markdown: String,
}

/// PDF payload for the shared document container.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PdfDocument {
    pub pages: Vec<PdfPage>,
    pub diagnostics: PdfDiagnostics,
}
