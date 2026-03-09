//! DOCX extraction helpers using `officemd_core`.

pub mod error;
mod extract;
mod render;

pub use error::DocxError;
pub use extract::{extract_ir, extract_ir_json};
pub use render::{markdown_from_bytes, markdown_from_bytes_with_options};
