//! XLSX-specific helpers using `officemd_core`.
//!
//! This crate provides XLSX extraction components built on the
//! shared OOXML core + markdown crates.

pub mod error;
pub mod extract_ir;
pub mod inspect;
pub mod render;
mod sheet_reader;
mod style_format;
pub mod table_ir;

pub use error::XlsxError;
pub use extract_ir::extract_sheet_names;
pub use inspect::{XlsxSheetSummary, inspect_sheet_summaries};
pub use render::{markdown_from_bytes, markdown_from_bytes_with_options};
pub use table_ir::{
    SheetFilter, XlsxExtractOptions, extract_tables_ir, extract_tables_ir_with_options,
};
