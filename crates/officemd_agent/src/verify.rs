use std::path::PathBuf;

use serde::{Deserialize, Serialize};

use crate::{
    ArtifactRef, Diagnostic,
    artifact::{AgentDocumentFormat, read_artifact, resolve_artifact},
    capability::VerificationCheckKind,
    error::{AgentError, AgentResult},
    locator::ArtifactLocator,
    render::{ArtifactRenderer, RenderReport, RenderRequest, RenderScale},
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
    pub schema_version: u32,
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

pub fn verify(
    request: &VerifyRequest,
    renderer: &(dyn ArtifactRenderer + Send + Sync),
) -> AgentResult<VerificationReport> {
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
            VerificationCheckKind::PptxCanvasOverflow => {
                checks.push(pptx_canvas_overflow_check(&artifact, &bytes)?);
            }
            VerificationCheckKind::VisualRender => {}
        }
    }

    let rendered_evidence = if request
        .checks
        .contains(&VerificationCheckKind::VisualRender)
    {
        let (check, evidence) = visual_render_check(request, renderer);
        checks.push(check);
        evidence
    } else {
        None
    };

    Ok(VerificationReport {
        schema_version: crate::AGENT_SCHEMA_VERSION,
        artifact,
        status: aggregate_status(&checks),
        checks,
        diagnostics: Vec::new(),
        rendered_evidence,
    })
}

fn visual_render_check(
    request: &VerifyRequest,
    renderer: &(dyn ArtifactRenderer + Send + Sync),
) -> (CheckReport, Option<RenderReport>) {
    let Some(output_dir) = &request.render_output_dir else {
        return (
            CheckReport {
                kind: VerificationCheckKind::VisualRender,
                status: VerificationStatus::NotRunnable,
                locator: None,
                message: "visual verification requires render_output_dir".to_string(),
            },
            None,
        );
    };
    let render_request = RenderRequest {
        input: request.input.clone(),
        output_dir: output_dir.clone(),
        pages_or_slides: None,
        scale: RenderScale::Screen,
    };
    match renderer.render(&render_request) {
        Ok(report) => (
            CheckReport {
                kind: VerificationCheckKind::VisualRender,
                status: VerificationStatus::Passed,
                locator: None,
                message: format!("rendered {} visual evidence image(s)", report.images.len()),
            },
            Some(report),
        ),
        Err(error) => (
            CheckReport {
                kind: VerificationCheckKind::VisualRender,
                status: if matches!(
                    error,
                    AgentError::RenderUnavailable(_) | AgentError::UnsupportedCapability(_)
                ) {
                    VerificationStatus::NotRunnable
                } else {
                    VerificationStatus::Failed
                },
                locator: None,
                message: error.to_string(),
            },
            None,
        ),
    }
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
    let errors = officemd_xlsx::inspect_formula_errors(bytes)
        .map_err(|error| AgentError::Extraction(error.to_string()))?;
    if let Some(error) = errors.first() {
        return Ok(CheckReport {
            kind: VerificationCheckKind::FormulaReferences,
            status: VerificationStatus::Failed,
            locator: Some(ArtifactLocator::XlsxCell {
                sheet: error.sheet.clone(),
                address: error.address.clone(),
            }),
            message: format!("stored formula error detected: {}", error.error),
        });
    }
    Ok(CheckReport {
        kind: VerificationCheckKind::FormulaReferences,
        status: VerificationStatus::Passed,
        locator: None,
        message: "no stored formula reference errors detected".to_string(),
    })
}

fn pptx_canvas_overflow_check(artifact: &ArtifactRef, bytes: &[u8]) -> AgentResult<CheckReport> {
    if artifact.format != AgentDocumentFormat::Pptx {
        return Ok(CheckReport {
            kind: VerificationCheckKind::PptxCanvasOverflow,
            status: VerificationStatus::NotRunnable,
            locator: None,
            message: "canvas overflow checks only apply to PPTX artifacts".to_string(),
        });
    }
    let overflows = officemd_pptx::inspect_canvas_overflows(bytes)
        .map_err(|error| AgentError::Extraction(error.to_string()))?;
    if let Some(first) = overflows.first() {
        return Ok(CheckReport {
            kind: VerificationCheckKind::PptxCanvasOverflow,
            status: VerificationStatus::Failed,
            locator: Some(ArtifactLocator::PptxShape {
                slide_number: first.slide_number,
                shape_id: first.shape_id,
            }),
            message: format!(
                "{} shape(s) extend beyond the slide canvas",
                overflows.len()
            ),
        });
    }
    Ok(CheckReport {
        kind: VerificationCheckKind::PptxCanvasOverflow,
        status: VerificationStatus::Passed,
        locator: None,
        message: "no shape canvas overflow detected".to_string(),
    })
}

fn aggregate_status(checks: &[CheckReport]) -> VerificationStatus {
    if checks
        .iter()
        .any(|check| check.status == VerificationStatus::Failed)
    {
        VerificationStatus::Failed
    } else if !checks.is_empty()
        && checks
            .iter()
            .all(|check| check.status == VerificationStatus::NotRunnable)
    {
        VerificationStatus::NotRunnable
    } else if checks
        .iter()
        .any(|check| check.status == VerificationStatus::NotRunnable)
    {
        VerificationStatus::PassedWithWarnings
    } else {
        VerificationStatus::Passed
    }
}
