use std::path::Path;

use officemd_core::{
    DocxPatch, DocxTextScope, PptxPatch, PptxTextScope, ScopedDocxReplace, ScopedPptxReplace,
    ScopedXlsxReplace, TextReplace, XlsxPatch, XlsxSheetRename, XlsxTextScope,
};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::{
    ArtifactRef,
    artifact::{AgentDocumentFormat, fingerprint_bytes, read_artifact, resolve_artifact},
    capability::MutationKind,
    diagnostic::Diagnostic,
    error::{AgentError, AgentResult},
    locator::ArtifactLocator,
    patch_plan::{ApplyPatchRequest, PatchOperation},
};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ApplyPatchReport {
    pub source: ArtifactRef,
    pub output: Option<ArtifactRef>,
    pub request_id: String,
    pub patch_sha256: String,
    pub status: ApplyPatchStatus,
    pub operations: Vec<OperationReport>,
    pub diagnostics: Vec<Diagnostic>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum ApplyPatchStatus {
    Applied,
    AlreadyApplied,
    Rejected,
    Failed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OperationReport {
    pub operation_index: u32,
    pub operation: MutationKind,
    pub target: ArtifactLocator,
    pub status: OperationStatus,
    pub message: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum OperationStatus {
    Applied,
    Skipped,
    Failed,
    Unsupported,
}

pub fn apply_patch(request: &ApplyPatchRequest) -> AgentResult<ApplyPatchReport> {
    let source_bytes = read_artifact(&request.input)?;
    let source = resolve_artifact(&request.input, &source_bytes, None)?;
    if source.fingerprint.sha256 != request.expected_source_sha256 {
        return Err(AgentError::PatchPreconditionFailed(format!(
            "expected source sha256 {}, got {}",
            request.expected_source_sha256, source.fingerprint.sha256
        )));
    }
    if request.input == request.output {
        return Err(AgentError::PatchPreconditionFailed(
            "output path must differ from input path".to_string(),
        ));
    }
    if request.output.exists() {
        return Err(AgentError::PatchPreconditionFailed(format!(
            "output already exists: {}",
            request.output.display()
        )));
    }

    let patch_sha256 = patch_sha256(&request.patch)?;
    let (patched_bytes, operations) = match source.format {
        AgentDocumentFormat::Docx => apply_docx(&source_bytes, &request.patch.operations)?,
        AgentDocumentFormat::Xlsx => apply_xlsx(&source_bytes, &request.patch.operations)?,
        AgentDocumentFormat::Pptx => apply_pptx(&source_bytes, &request.patch.operations)?,
        AgentDocumentFormat::Csv | AgentDocumentFormat::Pdf => {
            return Err(AgentError::UnsupportedCapability(format!(
                "{} mutation is not supported",
                source.format
            )));
        }
    };

    if operations.iter().any(|op| {
        matches!(
            op.status,
            OperationStatus::Failed | OperationStatus::Unsupported
        )
    }) {
        return Ok(ApplyPatchReport {
            source,
            output: None,
            request_id: request.patch.request_id.clone(),
            patch_sha256,
            status: ApplyPatchStatus::Rejected,
            operations,
            diagnostics: Vec::new(),
        });
    }

    write_atomic(&request.output, &patched_bytes)?;
    let output = ArtifactRef {
        path: request.output.clone(),
        format: source.format,
        fingerprint: fingerprint_bytes(&patched_bytes),
    };

    Ok(ApplyPatchReport {
        source,
        output: Some(output),
        request_id: request.patch.request_id.clone(),
        patch_sha256,
        status: ApplyPatchStatus::Applied,
        operations,
        diagnostics: Vec::new(),
    })
}

fn apply_docx(
    source_bytes: &[u8],
    operations: &[PatchOperation],
) -> AgentResult<(Vec<u8>, Vec<OperationReport>)> {
    let mut patch = DocxPatch::default();
    let mut reports = Vec::new();
    for (index, operation) in operations.iter().enumerate() {
        match operation {
            PatchOperation::ReplaceText {
                target,
                expected_text,
                replacement,
                preserve_formatting,
            } if matches!(
                target,
                ArtifactLocator::DocxParagraph { .. } | ArtifactLocator::DocxTableCell { .. }
            ) =>
            {
                patch.scoped_replacements.push(ScopedDocxReplace {
                    scope: DocxTextScope::AllText,
                    replace: TextReplace::all(expected_text, replacement)
                        .with_preserve_formatting(*preserve_formatting),
                });
                reports.push(operation_report(
                    index,
                    MutationKind::ReplaceText,
                    target.clone(),
                    OperationStatus::Applied,
                    "queued DOCX text replacement",
                ));
            }
            _ => reports.push(unsupported_operation(index, operation)),
        }
    }
    if reports
        .iter()
        .any(|report| report.status == OperationStatus::Unsupported)
    {
        return Ok((source_bytes.to_vec(), reports));
    }
    let patched = officemd_core::patch_docx_with_report(source_bytes, &patch)
        .map_err(|e| AgentError::Extraction(e.to_string()))?;
    if patched.report.replacements_applied == 0 && !operations.is_empty() {
        return Err(AgentError::PatchPreconditionFailed(
            "no DOCX replacement matched expected text".to_string(),
        ));
    }
    Ok((patched.content, reports))
}

fn apply_xlsx(
    source_bytes: &[u8],
    operations: &[PatchOperation],
) -> AgentResult<(Vec<u8>, Vec<OperationReport>)> {
    let mut patch = XlsxPatch::default();
    let mut reports = Vec::new();
    for (index, operation) in operations.iter().enumerate() {
        match operation {
            PatchOperation::RenameXlsxSheet { target, new_name } => {
                patch.rename_sheets.push(XlsxSheetRename {
                    from: target.name.clone(),
                    to: new_name.clone(),
                    update_references: true,
                });
                reports.push(operation_report(
                    index,
                    MutationKind::RenameXlsxSheet,
                    ArtifactLocator::XlsxSheet {
                        name: target.name.clone(),
                    },
                    OperationStatus::Applied,
                    "queued XLSX sheet rename",
                ));
            }
            PatchOperation::ReplaceText {
                target,
                expected_text,
                replacement,
                preserve_formatting: _,
            } => {
                patch.scoped_replacements.push(ScopedXlsxReplace {
                    scope: XlsxTextScope::AllText,
                    replace: TextReplace::all(expected_text, replacement),
                });
                reports.push(operation_report(
                    index,
                    MutationKind::ReplaceText,
                    target.clone(),
                    OperationStatus::Applied,
                    "queued XLSX text replacement",
                ));
            }
            _ => reports.push(unsupported_operation(index, operation)),
        }
    }
    if reports
        .iter()
        .any(|report| report.status == OperationStatus::Unsupported)
    {
        return Ok((source_bytes.to_vec(), reports));
    }
    let patched = officemd_core::patch_xlsx_with_report(source_bytes, &patch)
        .map_err(|e| AgentError::Extraction(e.to_string()))?;
    Ok((patched.content, reports))
}

fn apply_pptx(
    source_bytes: &[u8],
    operations: &[PatchOperation],
) -> AgentResult<(Vec<u8>, Vec<OperationReport>)> {
    let mut patch = PptxPatch::default();
    let mut reports = Vec::new();
    for (index, operation) in operations.iter().enumerate() {
        match operation {
            PatchOperation::ReplacePptxShapeText {
                target,
                expected_text,
                replacement,
            } => {
                patch.scoped_replacements.push(ScopedPptxReplace {
                    scope: PptxTextScope::AllText,
                    replace: TextReplace::all(expected_text, replacement),
                });
                reports.push(operation_report(
                    index,
                    MutationKind::ReplacePptxShapeText,
                    target.clone(),
                    OperationStatus::Applied,
                    "queued PPTX text replacement",
                ));
            }
            _ => reports.push(unsupported_operation(index, operation)),
        }
    }
    if reports
        .iter()
        .any(|report| report.status == OperationStatus::Unsupported)
    {
        return Ok((source_bytes.to_vec(), reports));
    }
    let patched = officemd_core::patch_pptx_with_report(source_bytes, &patch)
        .map_err(|e| AgentError::Extraction(e.to_string()))?;
    if patched.report.replacements_applied == 0 && !operations.is_empty() {
        return Err(AgentError::PatchPreconditionFailed(
            "no PPTX replacement matched expected text".to_string(),
        ));
    }
    Ok((patched.content, reports))
}

fn unsupported_operation(index: usize, operation: &PatchOperation) -> OperationReport {
    let (kind, target) = match operation {
        PatchOperation::ReplaceText { target, .. } => (MutationKind::ReplaceText, target.clone()),
        PatchOperation::SetXlsxCellValue { target, .. } => (
            MutationKind::SetXlsxCellValue,
            ArtifactLocator::XlsxCell {
                sheet: target.sheet.clone(),
                address: target.address.clone(),
            },
        ),
        PatchOperation::SetXlsxCellFormula { target, .. } => (
            MutationKind::SetXlsxCellFormula,
            ArtifactLocator::XlsxCell {
                sheet: target.sheet.clone(),
                address: target.address.clone(),
            },
        ),
        PatchOperation::RenameXlsxSheet { target, .. } => (
            MutationKind::RenameXlsxSheet,
            ArtifactLocator::XlsxSheet {
                name: target.name.clone(),
            },
        ),
        PatchOperation::ReplacePptxShapeText { target, .. } => {
            (MutationKind::ReplacePptxShapeText, target.clone())
        }
    };
    operation_report(
        index,
        kind,
        target,
        OperationStatus::Unsupported,
        "operation is not supported by this implementation slice",
    )
}

fn operation_report(
    index: usize,
    operation: MutationKind,
    target: ArtifactLocator,
    status: OperationStatus,
    message: &str,
) -> OperationReport {
    OperationReport {
        operation_index: u32::try_from(index).unwrap_or(u32::MAX),
        operation,
        target,
        status,
        message: message.to_string(),
    }
}

fn patch_sha256(patch: &crate::ArtifactPatchPlan) -> AgentResult<String> {
    let bytes = serde_json::to_vec(patch)?;
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    Ok(format!("{:x}", hasher.finalize()))
}

fn write_atomic(output: &Path, bytes: &[u8]) -> AgentResult<()> {
    let parent = output.parent().unwrap_or_else(|| Path::new("."));
    let file_name = output
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("artifact");
    let tmp = parent.join(format!(".{file_name}.officemd-tmp-{}", std::process::id()));
    std::fs::write(&tmp, bytes).map_err(|source| AgentError::Read {
        path: tmp.display().to_string(),
        source,
    })?;
    std::fs::rename(&tmp, output).map_err(|source| AgentError::Read {
        path: output.display().to_string(),
        source,
    })
}
