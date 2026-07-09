use std::path::PathBuf;

use serde::{Deserialize, Serialize};

use crate::{
    ArtifactRef, Diagnostic,
    artifact::{AgentDocumentFormat, read_artifact, resolve_artifact},
    capability::VerificationCheckKind,
    error::{AgentError, AgentResult},
    locator::ArtifactLocator,
    render::RenderReport,
};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct VerifyRequest {
    pub input: PathBuf,
    pub checks: Vec<VerificationCheckKind>,
    pub render_output_dir: Option<PathBuf>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VerificationReport {
    pub artifact: ArtifactRef,
    pub status: VerificationStatus,
    pub checks: Vec<CheckReport>,
    pub diagnostics: Vec<Diagnostic>,
    pub rendered_evidence: Option<RenderReport>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum VerificationStatus {
    Passed,
    PassedWithWarnings,
    Failed,
    NotRunnable,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct CheckReport {
    pub kind: VerificationCheckKind,
    pub status: VerificationStatus,
    pub locator: Option<ArtifactLocator>,
    pub message: String,
}

pub fn verify(request: &VerifyRequest) -> AgentResult<VerificationReport> {
    let bytes = read_artifact(&request.input)?;
    let artifact = resolve_artifact(&request.input, &bytes, None)?;
    let requested = if request.checks.is_empty() {
        vec![VerificationCheckKind::Structure]
    } else {
        request.checks.clone()
    };

    let mut checks = Vec::with_capacity(requested.len());
    for check in requested {
        match check {
            VerificationCheckKind::Structure => checks.push(structure_check(&artifact, &bytes)),
            VerificationCheckKind::FormulaReferences => {
                checks.push(formula_reference_check(&artifact, &bytes)?);
            }
            VerificationCheckKind::VisualRender => checks.push(CheckReport {
                kind: VerificationCheckKind::VisualRender,
                status: VerificationStatus::NotRunnable,
                locator: None,
                message: "visual verification needs a configured renderer backend".to_string(),
            }),
        }
    }

    Ok(VerificationReport {
        artifact,
        status: aggregate_status(&checks),
        checks,
        diagnostics: Vec::new(),
        rendered_evidence: None,
    })
}

fn structure_check(artifact: &ArtifactRef, bytes: &[u8]) -> CheckReport {
    let result = match artifact.format {
        AgentDocumentFormat::Docx => officemd_docx::extract_ir(bytes)
            .map(|_| ())
            .map_err(|e| e.to_string()),
        AgentDocumentFormat::Xlsx => officemd_xlsx::extract_tables_ir(bytes)
            .map(|_| ())
            .map_err(|e| e.to_string()),
        AgentDocumentFormat::Csv => officemd_csv::extract_tables_ir(bytes)
            .map(|_| ())
            .map_err(|e| e.to_string()),
        AgentDocumentFormat::Pptx => officemd_pptx::extract_ir(bytes)
            .map(|_| ())
            .map_err(|e| e.to_string()),
        AgentDocumentFormat::Pdf => officemd_pdf::extract_ir(bytes)
            .map(|_| ())
            .map_err(|e| e.to_string()),
    };
    match result {
        Ok(()) => CheckReport {
            kind: VerificationCheckKind::Structure,
            status: VerificationStatus::Passed,
            locator: None,
            message: "artifact parsed successfully".to_string(),
        },
        Err(message) => CheckReport {
            kind: VerificationCheckKind::Structure,
            status: VerificationStatus::Failed,
            locator: None,
            message,
        },
    }
}

fn formula_reference_check(artifact: &ArtifactRef, bytes: &[u8]) -> AgentResult<CheckReport> {
    if artifact.format != AgentDocumentFormat::Xlsx {
        return Ok(CheckReport {
            kind: VerificationCheckKind::FormulaReferences,
            status: VerificationStatus::NotRunnable,
            locator: None,
            message: "formula reference checks only apply to XLSX artifacts".to_string(),
        });
    }
    let doc = officemd_xlsx::extract_tables_ir(bytes)
        .map_err(|e| AgentError::Extraction(e.to_string()))?;
    for sheet in doc.sheets {
        for formula in sheet.formulas {
            if formula.formula.contains("#REF!") {
                return Ok(CheckReport {
                    kind: VerificationCheckKind::FormulaReferences,
                    status: VerificationStatus::Failed,
                    locator: Some(ArtifactLocator::XlsxCell {
                        sheet: sheet.name,
                        address: formula.cell_ref,
                    }),
                    message: "formula contains #REF!".to_string(),
                });
            }
        }
    }
    Ok(CheckReport {
        kind: VerificationCheckKind::FormulaReferences,
        status: VerificationStatus::Passed,
        locator: None,
        message: "no stored formula reference errors detected".to_string(),
    })
}

fn aggregate_status(checks: &[CheckReport]) -> VerificationStatus {
    if checks
        .iter()
        .any(|check| check.status == VerificationStatus::Failed)
    {
        VerificationStatus::Failed
    } else if checks
        .iter()
        .any(|check| check.status == VerificationStatus::NotRunnable)
    {
        VerificationStatus::PassedWithWarnings
    } else {
        VerificationStatus::Passed
    }
}
