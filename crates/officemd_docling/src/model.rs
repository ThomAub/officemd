//! Docling Document v1.9.0 model types.

use serde::Serialize;
use std::collections::BTreeMap;

/// Content layer discriminator.
#[derive(Debug, Clone, Copy, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum ContentLayer {
    Body,
    Furniture,
}

/// Group label discriminator.
#[derive(Debug, Clone, Copy, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum GroupLabel {
    List,
    OrderedList,
    Section,
    Chapter,
    Unspecified,
}

/// Document item label discriminator.
#[derive(Debug, Clone, Copy, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum DocItemLabel {
    Title,
    SectionHeader,
    Paragraph,
    ListItem,
    Text,
    Caption,
    Table,
    Footnote,
    PageHeader,
    PageFooter,
    Formula,
    Code,
    Reference,
}

/// JSON pointer reference: `{"$ref": "#/texts/0"}`.
#[derive(Debug, Clone, Serialize)]
pub struct RefItem {
    #[serde(rename = "$ref")]
    pub cref: String,
}

impl RefItem {
    pub fn new(path: impl Into<String>) -> Self {
        Self { cref: path.into() }
    }
}

/// Group container (used for body, furniture, and groups list entries).
#[derive(Debug, Clone, Serialize)]
pub struct GroupItem {
    pub self_ref: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent: Option<RefItem>,
    pub children: Vec<RefItem>,
    pub content_layer: ContentLayer,
    pub name: String,
    pub label: GroupLabel,
}

/// Text document item (covers Title, `SectionHeader`, `ListItem`, Paragraph, etc.).
#[derive(Debug, Clone, Serialize)]
pub struct TextItem {
    pub self_ref: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent: Option<RefItem>,
    pub children: Vec<RefItem>,
    pub content_layer: ContentLayer,
    pub label: DocItemLabel,
    pub prov: Vec<serde_json::Value>,
    pub orig: String,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub level: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub enumerated: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker: Option<String>,
}

/// Single cell in a Docling table.
#[derive(Debug, Clone, Serialize)]
pub struct DoclingTableCell {
    pub row_span: usize,
    pub col_span: usize,
    pub start_row_offset_idx: usize,
    pub end_row_offset_idx: usize,
    pub start_col_offset_idx: usize,
    pub end_col_offset_idx: usize,
    pub text: String,
    pub column_header: bool,
    pub row_header: bool,
    pub row_section: bool,
}

/// Flat table data with cell list and dimensions.
#[derive(Debug, Clone, Serialize)]
pub struct TableData {
    pub table_cells: Vec<DoclingTableCell>,
    pub num_rows: usize,
    pub num_cols: usize,
}

/// Table document item.
#[derive(Debug, Clone, Serialize)]
pub struct TableItem {
    pub self_ref: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent: Option<RefItem>,
    pub children: Vec<RefItem>,
    pub content_layer: ContentLayer,
    pub label: DocItemLabel,
    pub prov: Vec<serde_json::Value>,
    pub data: TableData,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub captions: Vec<RefItem>,
}

/// Page dimensions.
#[derive(Debug, Clone, Serialize)]
pub struct Size {
    pub width: f64,
    pub height: f64,
}

/// Single page entry.
#[derive(Debug, Clone, Serialize)]
pub struct PageItem {
    pub size: Size,
    pub page_no: usize,
}

/// Top-level Docling Document (v1.9.0).
#[derive(Debug, Clone, Serialize)]
pub struct DoclingDocument {
    pub schema_name: String,
    pub version: String,
    pub name: String,
    pub furniture: GroupItem,
    pub body: GroupItem,
    pub groups: Vec<GroupItem>,
    pub texts: Vec<TextItem>,
    pub pictures: Vec<serde_json::Value>,
    pub tables: Vec<TableItem>,
    pub key_value_items: Vec<serde_json::Value>,
    pub pages: BTreeMap<String, PageItem>,
}
