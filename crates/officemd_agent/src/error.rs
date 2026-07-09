use thiserror::Error;

pub type AgentResult<T> = Result<T, AgentError>;

#[derive(Debug, Error)]
pub enum AgentError {
    #[error("failed to read artifact '{path}': {source}")]
    Read {
        path: String,
        source: std::io::Error,
    },
    #[error("unsupported or unrecognized artifact format: {0}")]
    Format(String),
    #[error("artifact extraction failed: {0}")]
    Extraction(String),
    #[error("invalid request: {0}")]
    InvalidRequest(String),
    #[error("unsupported capability: {0}")]
    UnsupportedCapability(String),
    #[error("rendering is unavailable: {0}")]
    RenderUnavailable(String),
    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),
}
