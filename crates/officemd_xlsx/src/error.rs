use officemd_core::opc::OpcError;
use thiserror::Error;
use zip::result::ZipError;

/// Errors that can occur during XLSX extraction.
#[derive(Debug, Error)]
pub enum XlsxError {
    #[error("ZIP error: {0}")]
    Zip(#[from] ZipError),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("XML parse error: {0}")]
    Xml(String),
    #[error("Missing required part: {0}")]
    MissingPart(String),
}

impl From<OpcError> for XlsxError {
    fn from(err: OpcError) -> Self {
        match err {
            OpcError::Zip(e) => XlsxError::Zip(e),
            OpcError::Io(e) => XlsxError::Io(e),
            OpcError::Xml(e) => XlsxError::Xml(e),
            OpcError::MissingPart(path) => XlsxError::MissingPart(path),
            OpcError::Write(msg) => XlsxError::Xml(msg),
        }
    }
}
