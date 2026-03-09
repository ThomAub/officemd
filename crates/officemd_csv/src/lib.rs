//! CSV-specific helpers to build table IR and render Markdown.

pub mod error;
pub mod extract_ir;
pub mod render;
pub mod table_ir;

pub use error::CsvError;
pub use extract_ir::extract_sheet_names;
pub use render::{markdown_from_bytes, markdown_from_bytes_with_options};
pub use table_ir::{
    CsvExtractOptions, extract_tables_ir, extract_tables_ir_json,
    extract_tables_ir_json_with_options, extract_tables_ir_with_options,
};
