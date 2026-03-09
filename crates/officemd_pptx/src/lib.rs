//! PPTX extraction helpers using `officemd_core`.

pub mod error;
pub mod extract;
mod render;

pub use error::PptxError;
pub use extract::{PptxExtractOptions, extract_ir, extract_ir_json, extract_ir_with_options};
pub use render::{markdown_from_bytes, markdown_from_bytes_with_options};
