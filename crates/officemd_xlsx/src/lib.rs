//! XLSX extraction and generation using `officemd_core`.
//!
//! This crate provides XLSX extraction and generation components built on the
//! shared OOXML core + markdown crates.

pub mod error;
pub mod extract_ir;
pub mod inspect;
pub mod mutate;
pub mod render;
mod sheet_reader;
mod style_format;
pub mod table_ir;
pub mod write;

pub use error::XlsxError;
pub use extract_ir::extract_sheet_names;
pub use inspect::{
    XlsxCellValue, XlsxFormulaError, XlsxSheetSummary, inspect_cells, inspect_formula_errors,
    inspect_sheet_summaries,
};
pub use mutate::{XlsxCellUpdate, set_cell};
pub use render::{
    markdown_from_bytes, markdown_from_bytes_with_extract_options, markdown_from_bytes_with_options,
};
pub use table_ir::{
    SheetFilter, XlsxExtractOptions, extract_tables_ir, extract_tables_ir_with_options,
};
pub use write::generate_xlsx;
