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
    PptxCanvasOverflow,
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
pub fn capability_for(
    format: AgentDocumentFormat,
    discovered_rendering: RenderCapability,
) -> ArtifactCapabilityReport {
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

    let render_capability = capability_for_format(format, discovered_rendering);
    let mut risks = Vec::new();
    if matches!(render_capability, RenderCapability::Unavailable { .. }) {
        risks.push(ArtifactRisk::MissingRenderingBackend);
    }
    if format == AgentDocumentFormat::Xlsx {
        risks.push(ArtifactRisk::FormulaEvaluationUnavailable);
    }
    if format == AgentDocumentFormat::Pdf {
        risks.push(ArtifactRisk::PdfMayRequireOcr);
    }

    let mut verification_checks = vec![VerificationCheckKind::Structure];
    if format == AgentDocumentFormat::Xlsx {
        verification_checks.push(VerificationCheckKind::FormulaReferences);
    }
    if format == AgentDocumentFormat::Pptx {
        verification_checks.push(VerificationCheckKind::PptxCanvasOverflow);
    }
    if matches!(render_capability, RenderCapability::Available { .. }) {
        verification_checks.push(VerificationCheckKind::VisualRender);
    }

    ArtifactCapabilityReport {
        readable: true,
        mutable_operations,
        render_capability,
        verification_checks,
        risks,
    }
}

fn capability_for_format(
    format: AgentDocumentFormat,
    capability: RenderCapability,
) -> RenderCapability {
    match capability {
        RenderCapability::Available { backend, formats } if formats.contains(&format) => {
            let backend = if format == AgentDocumentFormat::Pdf
                && backend == RenderBackendKind::LibreOffice
            {
                RenderBackendKind::Poppler
            } else {
                backend
            };
            RenderCapability::Available {
                backend,
                formats: vec![format],
            }
        }
        RenderCapability::Available { .. } => RenderCapability::Unavailable {
            reason: RenderUnavailableReason::BackendNotConfigured,
        },
        unavailable => unavailable,
    }
}
