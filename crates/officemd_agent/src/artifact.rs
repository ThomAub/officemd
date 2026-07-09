use std::fmt::{Display, Formatter};
use std::path::{Path, PathBuf};

use officemd_core::format as core_format;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::error::{AgentError, AgentResult};

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum AgentDocumentFormat {
    Docx,
    Xlsx,
    Csv,
    Pptx,
    Pdf,
}

impl Display for AgentDocumentFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Docx => write!(f, "docx"),
            Self::Xlsx => write!(f, "xlsx"),
            Self::Csv => write!(f, "csv"),
            Self::Pptx => write!(f, "pptx"),
            Self::Pdf => write!(f, "pdf"),
        }
    }
}

impl From<core_format::DocumentFormat> for AgentDocumentFormat {
    fn from(value: core_format::DocumentFormat) -> Self {
        match value {
            core_format::DocumentFormat::Docx => Self::Docx,
            core_format::DocumentFormat::Xlsx => Self::Xlsx,
            core_format::DocumentFormat::Csv => Self::Csv,
            core_format::DocumentFormat::Pptx => Self::Pptx,
            core_format::DocumentFormat::Pdf => Self::Pdf,
        }
    }
}

impl From<AgentDocumentFormat> for core_format::DocumentFormat {
    fn from(value: AgentDocumentFormat) -> Self {
        match value {
            AgentDocumentFormat::Docx => Self::Docx,
            AgentDocumentFormat::Xlsx => Self::Xlsx,
            AgentDocumentFormat::Csv => Self::Csv,
            AgentDocumentFormat::Pptx => Self::Pptx,
            AgentDocumentFormat::Pdf => Self::Pdf,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ArtifactFingerprint {
    pub sha256: String,
    pub byte_length: u64,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ArtifactRef {
    pub path: PathBuf,
    pub format: AgentDocumentFormat,
    pub fingerprint: ArtifactFingerprint,
}

pub(crate) fn read_artifact(path: &Path) -> AgentResult<Vec<u8>> {
    std::fs::read(path).map_err(|source| AgentError::Read {
        path: path.display().to_string(),
        source,
    })
}

#[must_use]
pub fn fingerprint_bytes(bytes: &[u8]) -> ArtifactFingerprint {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    ArtifactFingerprint {
        sha256: format!("{:x}", hasher.finalize()),
        byte_length: u64::try_from(bytes.len()).unwrap_or(u64::MAX),
    }
}

pub(crate) fn resolve_artifact(
    path: &Path,
    bytes: &[u8],
    explicit: Option<AgentDocumentFormat>,
) -> AgentResult<ArtifactRef> {
    let format = if let Some(format) = explicit {
        format
    } else {
        detect_format(bytes, path)?.into()
    };
    Ok(ArtifactRef {
        path: path.to_path_buf(),
        format,
        fingerprint: fingerprint_bytes(bytes),
    })
}

fn detect_format(bytes: &[u8], path: &Path) -> AgentResult<core_format::DocumentFormat> {
    if let Some(extension) = path.extension().and_then(|ext| ext.to_str())
        && let Some(format) = core_format::parse_format(extension)
    {
        return Ok(format);
    }
    core_format::detect_format_from_bytes(bytes).map_err(AgentError::Format)
}
