use serde::{Deserialize, Serialize};

use crate::artifact::AgentDocumentFormat;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ArtifactCapabilityReport {
    pub readable: bool,
    pub mutable_operations: Vec<MutationKind>,
    pub render_capability: RenderCapability,
    pub verification_checks: Vec<VerificationCheckKind>,
    pub risks: Vec<ArtifactRisk>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MutationKind {
    ReplaceText,
    SetXlsxCellValue,
    SetXlsxCellFormula,
    RenameXlsxSheet,
    ReplacePptxShapeText,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(tag = "status", rename_all = "snake_case")]
pub enum RenderCapability {
    Available {
        backend: RenderBackendKind,
        formats: Vec<AgentDocumentFormat>,
    },
    Unavailable {
        reason: RenderUnavailableReason,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum RenderBackendKind {
    Poppler,
    LibreOffice,
    Unavailable,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum RenderUnavailableReason {
    BackendNotConfigured,
    MissingExecutable,
    UnsupportedPlatform,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum VerificationCheckKind {
    Structure,
    FormulaReferences,
    VisualRender,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum ArtifactRisk {
    MissingRenderingBackend,
    FormulaEvaluationUnavailable,
    PdfMayRequireOcr,
}

#[must_use]
pub fn capability_for(format: AgentDocumentFormat) -> ArtifactCapabilityReport {
    let mutable_operations = match format {
        AgentDocumentFormat::Docx => vec![MutationKind::ReplaceText],
        AgentDocumentFormat::Xlsx => vec![
            MutationKind::SetXlsxCellValue,
            MutationKind::SetXlsxCellFormula,
            MutationKind::RenameXlsxSheet,
        ],
        AgentDocumentFormat::Pptx => vec![MutationKind::ReplacePptxShapeText],
        AgentDocumentFormat::Csv | AgentDocumentFormat::Pdf => Vec::new(),
    };

    let mut risks = vec![ArtifactRisk::MissingRenderingBackend];
    if format == AgentDocumentFormat::Xlsx {
        risks.push(ArtifactRisk::FormulaEvaluationUnavailable);
    }

    ArtifactCapabilityReport {
        readable: true,
        mutable_operations,
        render_capability: RenderCapability::Unavailable {
            reason: RenderUnavailableReason::BackendNotConfigured,
        },
        verification_checks: vec![VerificationCheckKind::Structure],
        risks,
    }
}
