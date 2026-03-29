use officemd_core::opc::OpcError;
use thiserror::Error;
use zip::result::ZipError;

/// Errors that can occur during DOCX extraction.
#[derive(Debug, Error)]
pub enum DocxError {
    #[error("ZIP error: {0}")]
    Zip(#[from] ZipError),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("XML parse error: {0}")]
    Xml(String),
    #[error("Missing required part: {0}")]
    MissingPart(String),
    #[error("JSON error: {0}")]
    Json(String),
}

impl From<OpcError> for DocxError {
    fn from(err: OpcError) -> Self {
        match err {
            OpcError::Zip(e) => DocxError::Zip(e),
            OpcError::Io(e) => DocxError::Io(e),
            OpcError::Xml(e) => DocxError::Xml(e),
            OpcError::MissingPart(path) => DocxError::MissingPart(path),
            OpcError::Write(msg) => DocxError::Xml(msg),
        }
    }
}
