//! DOCX extraction and generation using `officemd_core`.

pub mod error;
mod extract;
pub mod locator;
mod render;
pub mod write;

pub use error::DocxError;
pub use extract::{extract_ir, extract_ir_json};
pub use locator::{DocxLocator, replace_locator_text};
pub use render::{markdown_from_bytes, markdown_from_bytes_with_options};
pub use write::generate_docx;
