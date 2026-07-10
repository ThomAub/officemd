//! Optional renderer adapters for OfficeMD.

pub mod discovery;
pub mod libreoffice;
pub mod poppler;
pub mod unavailable;

use std::path::Path;

use officemd_agent::{
    AgentDocumentFormat, ArtifactFingerprint, ArtifactRef,
    artifact::fingerprint_bytes,
    error::{AgentError, AgentResult},
};

pub use discovery::{SystemRenderer, discover};
pub use libreoffice::LibreOfficeRenderer;
pub use poppler::PopplerRenderer;
pub use unavailable::UnavailableRenderer;

fn artifact_ref_for_path(path: &Path) -> AgentResult<ArtifactRef> {
    let bytes = std::fs::read(path).map_err(|source| AgentError::Read {
        path: path.display().to_string(),
        source,
    })?;
    let format = path
        .extension()
        .and_then(|ext| ext.to_str())
        .and_then(format_from_extension)
        .ok_or_else(|| {
            AgentError::Format(format!(
                "unsupported render input extension for {}",
                path.display()
            ))
        })?;
    Ok(ArtifactRef {
        path: path.to_path_buf(),
        format,
        fingerprint: fingerprint_bytes(&bytes),
    })
}

fn format_from_extension(value: &str) -> Option<AgentDocumentFormat> {
    match value.to_ascii_lowercase().as_str() {
        "docx" => Some(AgentDocumentFormat::Docx),
        "xlsx" => Some(AgentDocumentFormat::Xlsx),
        "pptx" => Some(AgentDocumentFormat::Pptx),
        "pdf" => Some(AgentDocumentFormat::Pdf),
        "csv" => Some(AgentDocumentFormat::Csv),
        _ => None,
    }
}

fn artifact_ref_for_pdf_path(path: &Path) -> AgentResult<ArtifactRef> {
    let bytes = std::fs::read(path).map_err(|source| AgentError::Read {
        path: path.display().to_string(),
        source,
    })?;
    Ok(ArtifactRef {
        path: path.to_path_buf(),
        format: AgentDocumentFormat::Pdf,
        fingerprint: ArtifactFingerprint {
            sha256: fingerprint_bytes(&bytes).sha256,
            byte_length: u64::try_from(bytes.len()).unwrap_or(u64::MAX),
        },
    })
}
