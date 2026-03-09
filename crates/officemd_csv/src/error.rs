use thiserror::Error;

/// Errors that can occur during CSV extraction.
#[derive(Debug, Error)]
pub enum CsvError {
    #[error("CSV parse error: {0}")]
    Csv(#[from] csv::Error),
    #[error("JSON serialization error: {0}")]
    Json(String),
}
