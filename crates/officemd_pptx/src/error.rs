use officemd_core::opc::OpcError;
use thiserror::Error;
use zip::result::ZipError;

/// Errors that can occur during PPTX extraction.
#[derive(Debug, Error)]
pub enum PptxError {
    #[error("ZIP error: {0}")]
    Zip(#[from] ZipError),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("XML parse error: {0}")]
    Xml(String),
    #[error("Missing required part: {0}")]
    MissingPart(String),
}

impl From<OpcError> for PptxError {
    fn from(err: OpcError) -> Self {
        match err {
            OpcError::Zip(e) => PptxError::Zip(e),
            OpcError::Io(e) => PptxError::Io(e),
            OpcError::Xml(e) => PptxError::Xml(e),
            OpcError::MissingPart(path) => PptxError::MissingPart(path),
            OpcError::Write(msg) => PptxError::Xml(msg),
        }
    }
}
