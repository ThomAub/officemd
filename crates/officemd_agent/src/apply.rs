use std::{
    ffi::OsString,
    io::Write,
    path::{Path, PathBuf},
};

use officemd_core::{XlsxPatch, XlsxSheetRename};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::{
    ArtifactRef,
    artifact::{AgentDocumentFormat, read_artifact, resolve_artifact},
    capability::MutationKind,
    diagnostic::Diagnostic,
    error::{AgentError, AgentResult},
    locator::ArtifactLocator,
    patch_plan::{ApplyPatchRequest, CellValue, PatchOperation},
};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ApplyPatchReport {
    pub schema_version: u32,
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
    validate_patch_request(request)?;
    let source_bytes = read_artifact(&request.input)?;
    let source = resolve_artifact(&request.input, &source_bytes, None)?;
    if source.fingerprint.sha256 != request.expected_source_sha256 {
        return Err(AgentError::PatchPreconditionFailed(format!(
            "expected source sha256 {}, got {}",
            request.expected_source_sha256, source.fingerprint.sha256
        )));
    }
    if same_path(&request.input, &request.output) {
        return Err(AgentError::PatchPreconditionFailed(
            "output path must differ from input path".to_string(),
        ));
    }
    let patch_sha256 = patch_sha256(&request.patch)?;
    let sidecar = sidecar_path(&request.output);
    if request.output.exists() || sidecar.exists() {
        if let Some(report) = matching_existing_report(request, &source, &patch_sha256, &sidecar)? {
            return Ok(report);
        }
        return Err(AgentError::PatchPreconditionFailed(format!(
            "output or patch sidecar already exists with conflicting provenance: {}",
            request.output.display()
        )));
    }

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
        let operations = operations
            .into_iter()
            .map(|mut operation| {
                if operation.status == OperationStatus::Applied {
                    operation.status = OperationStatus::Skipped;
                    operation.message =
                        "operation was not committed because the patch plan was rejected"
                            .to_string();
                }
                operation
            })
            .collect();
        return Ok(ApplyPatchReport {
            schema_version: crate::AGENT_SCHEMA_VERSION,
            source,
            output: None,
            request_id: request.patch.request_id.clone(),
            patch_sha256,
            status: ApplyPatchStatus::Rejected,
            operations,
            diagnostics: Vec::new(),
        });
    }

    validate_output(source.format, &patched_bytes)?;
    write_atomic(&request.output, &patched_bytes, &request.patch.request_id)?;
    let persisted_bytes = read_artifact(&request.output)?;
    let output = resolve_artifact(&request.output, &persisted_bytes, Some(source.format))?;

    let report = ApplyPatchReport {
        schema_version: crate::AGENT_SCHEMA_VERSION,
        source,
        output: Some(output),
        request_id: request.patch.request_id.clone(),
        patch_sha256,
        status: ApplyPatchStatus::Applied,
        operations,
        diagnostics: Vec::new(),
    };
    if let Err(error) = write_sidecar(&sidecar, &report, &request.patch.request_id) {
        let _ = std::fs::remove_file(&request.output);
        return Err(error);
    }
    Ok(report)
}

fn validate_patch_request(request: &ApplyPatchRequest) -> AgentResult<()> {
    if request.patch.request_id.trim().is_empty() {
        return Err(AgentError::InvalidRequest(
            "patch request_id must not be empty".to_string(),
        ));
    }
    if request.patch.operations.is_empty() {
        return Err(AgentError::InvalidRequest(
            "patch plan must contain at least one operation".to_string(),
        ));
    }
    for operation in &request.patch.operations {
        match operation {
            PatchOperation::ReplaceText { expected_text, .. }
            | PatchOperation::ReplacePptxShapeText { expected_text, .. }
                if expected_text.is_empty() =>
            {
                return Err(AgentError::InvalidRequest(
                    "replacement expected_text must not be empty".to_string(),
                ));
            }
            _ => {}
        }
    }
    Ok(())
}

fn same_path(left: &Path, right: &Path) -> bool {
    if left == right {
        return true;
    }
    match (std::fs::canonicalize(left), std::fs::canonicalize(right)) {
        (Ok(left), Ok(right)) => left == right,
        _ => false,
    }
}

fn sidecar_path(output: &Path) -> PathBuf {
    let mut value = OsString::from(output.as_os_str());
    value.push(".officemd.json");
    PathBuf::from(value)
}

fn matching_existing_report(
    request: &ApplyPatchRequest,
    source: &ArtifactRef,
    patch_sha256: &str,
    sidecar: &Path,
) -> AgentResult<Option<ApplyPatchReport>> {
    if !request.output.is_file() || !sidecar.is_file() {
        return Ok(None);
    }
    let sidecar_bytes = std::fs::read(sidecar).map_err(|source| AgentError::Read {
        path: sidecar.display().to_string(),
        source,
    })?;
    let mut report: ApplyPatchReport = serde_json::from_slice(&sidecar_bytes)?;
    let output_bytes = read_artifact(&request.output)?;
    let actual_output = resolve_artifact(&request.output, &output_bytes, Some(source.format))?;
    let matches = report.schema_version == crate::AGENT_SCHEMA_VERSION
        && report.request_id == request.patch.request_id
        && report.patch_sha256 == patch_sha256
        && report.source.fingerprint.sha256 == source.fingerprint.sha256
        && report.output.as_ref().is_some_and(|recorded| {
            recorded.path == request.output && recorded.fingerprint == actual_output.fingerprint
        });
    if !matches {
        return Ok(None);
    }
    report.status = ApplyPatchStatus::AlreadyApplied;
    for operation in &mut report.operations {
        operation.status = OperationStatus::Skipped;
        operation.message = "identical patch was already applied".to_string();
    }
    Ok(Some(report))
}

fn apply_docx(
    source_bytes: &[u8],
    operations: &[PatchOperation],
) -> AgentResult<(Vec<u8>, Vec<OperationReport>)> {
    let mut reports = Vec::new();
    let mut patched_bytes = source_bytes.to_vec();
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
                let replacements = apply_docx_locator_replace(
                    &mut patched_bytes,
                    target,
                    expected_text,
                    replacement,
                    *preserve_formatting,
                )?;
                if replacements == 0 {
                    return Err(AgentError::PatchPreconditionFailed(
                        "no DOCX replacement matched expected text at locator".to_string(),
                    ));
                }
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
    Ok((patched_bytes, reports))
}

fn apply_docx_locator_replace(
    content: &mut Vec<u8>,
    target: &ArtifactLocator,
    expected_text: &str,
    replacement: &str,
    preserve_formatting: bool,
) -> AgentResult<usize> {
    let locator = match target {
        ArtifactLocator::DocxParagraph {
            part,
            paragraph_index,
        } => officemd_docx::DocxLocator::Paragraph {
            part: part.part.clone(),
            paragraph_index: *paragraph_index,
        },
        ArtifactLocator::DocxTableCell {
            part,
            table_index,
            row_index,
            column_index,
        } => officemd_docx::DocxLocator::TableCell {
            part: part.part.clone(),
            table_index: *table_index,
            row_index: *row_index,
            column_index: *column_index,
        },
        _ => {
            return Err(AgentError::InvalidRequest(
                "target must be a DOCX locator".to_string(),
            ));
        }
    };
    let (rewritten, replacements) = officemd_docx::replace_locator_text(
        content,
        &locator,
        expected_text,
        replacement,
        preserve_formatting,
    )
    .map_err(|error| AgentError::PatchPreconditionFailed(error.to_string()))?;
    if replacements > 0 {
        *content = rewritten;
    }
    Ok(replacements)
}

fn apply_xlsx(
    source_bytes: &[u8],
    operations: &[PatchOperation],
) -> AgentResult<(Vec<u8>, Vec<OperationReport>)> {
    let mut patched_bytes = source_bytes.to_vec();
    let mut reports = Vec::new();
    for (index, operation) in operations.iter().enumerate() {
        match operation {
            PatchOperation::SetXlsxCellValue { target, value } => {
                let update = match value {
                    CellValue::Text(value) => officemd_xlsx::XlsxCellUpdate::Text(value.clone()),
                    CellValue::Number(value) => officemd_xlsx::XlsxCellUpdate::Number(*value),
                    CellValue::Bool(value) => officemd_xlsx::XlsxCellUpdate::Bool(*value),
                    CellValue::Blank => officemd_xlsx::XlsxCellUpdate::Blank,
                };
                patched_bytes = officemd_xlsx::set_cell(
                    &patched_bytes,
                    &target.sheet,
                    &target.address,
                    &update,
                )
                .map_err(|error| AgentError::PatchPreconditionFailed(error.to_string()))?;
                reports.push(operation_report(
                    index,
                    MutationKind::SetXlsxCellValue,
                    ArtifactLocator::XlsxCell {
                        sheet: target.sheet.clone(),
                        address: target.address.clone(),
                    },
                    OperationStatus::Applied,
                    "set XLSX cell value",
                ));
            }
            PatchOperation::SetXlsxCellFormula { target, formula } => {
                patched_bytes = officemd_xlsx::set_cell(
                    &patched_bytes,
                    &target.sheet,
                    &target.address,
                    &officemd_xlsx::XlsxCellUpdate::Formula(formula.clone()),
                )
                .map_err(|error| AgentError::PatchPreconditionFailed(error.to_string()))?;
                reports.push(operation_report(
                    index,
                    MutationKind::SetXlsxCellFormula,
                    ArtifactLocator::XlsxCell {
                        sheet: target.sheet.clone(),
                        address: target.address.clone(),
                    },
                    OperationStatus::Applied,
                    "set XLSX cell formula",
                ));
            }
            PatchOperation::RenameXlsxSheet { target, new_name } => {
                validate_xlsx_sheet_rename(&patched_bytes, &target.name, new_name)?;
                let patch = XlsxPatch {
                    rename_sheets: vec![XlsxSheetRename {
                        from: target.name.clone(),
                        to: new_name.clone(),
                        update_references: true,
                    }],
                    ..Default::default()
                };
                patched_bytes = officemd_core::patch_xlsx_with_report(&patched_bytes, &patch)
                    .map_err(|error| AgentError::PatchPreconditionFailed(error.to_string()))?
                    .content;
                reports.push(operation_report(
                    index,
                    MutationKind::RenameXlsxSheet,
                    ArtifactLocator::XlsxSheet {
                        name: target.name.clone(),
                    },
                    OperationStatus::Applied,
                    "renamed XLSX sheet",
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
    Ok((patched_bytes, reports))
}

fn validate_xlsx_sheet_rename(content: &[u8], from: &str, to: &str) -> AgentResult<()> {
    if to.trim().is_empty() || to.chars().count() > 31 {
        return Err(AgentError::InvalidRequest(
            "XLSX sheet name must contain between 1 and 31 characters".to_string(),
        ));
    }
    if to
        .chars()
        .any(|character| matches!(character, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(AgentError::InvalidRequest(
            "XLSX sheet name contains a forbidden character".to_string(),
        ));
    }
    let sheets = officemd_xlsx::inspect_sheet_summaries(content, None)
        .map_err(|error| AgentError::Extraction(error.to_string()))?;
    if !sheets.iter().any(|sheet| sheet.name == from) {
        return Err(AgentError::PatchPreconditionFailed(format!(
            "XLSX sheet not found: {from}"
        )));
    }
    if sheets
        .iter()
        .any(|sheet| sheet.name.eq_ignore_ascii_case(to) && !sheet.name.eq_ignore_ascii_case(from))
    {
        return Err(AgentError::PatchPreconditionFailed(format!(
            "XLSX sheet already exists: {to}"
        )));
    }
    Ok(())
}

fn apply_pptx(
    source_bytes: &[u8],
    operations: &[PatchOperation],
) -> AgentResult<(Vec<u8>, Vec<OperationReport>)> {
    let mut patched_bytes = source_bytes.to_vec();
    let mut reports = Vec::new();
    for (index, operation) in operations.iter().enumerate() {
        match operation {
            PatchOperation::ReplacePptxShapeText {
                target,
                expected_text,
                replacement,
            } => {
                let ArtifactLocator::PptxShape {
                    slide_number,
                    shape_id,
                } = target
                else {
                    reports.push(unsupported_operation(index, operation));
                    continue;
                };
                patched_bytes = officemd_pptx::replace_shape_text(
                    &patched_bytes,
                    *slide_number,
                    *shape_id,
                    expected_text,
                    replacement,
                )
                .map_err(|error| AgentError::PatchPreconditionFailed(error.to_string()))?;
                reports.push(operation_report(
                    index,
                    MutationKind::ReplacePptxShapeText,
                    target.clone(),
                    OperationStatus::Applied,
                    "replaced PPTX shape text",
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
    Ok((patched_bytes, reports))
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
        "operation is not supported for this artifact format",
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

fn validate_output(format: AgentDocumentFormat, bytes: &[u8]) -> AgentResult<()> {
    match format {
        AgentDocumentFormat::Docx => officemd_docx::extract_ir(bytes)
            .map(|_| ())
            .map_err(|error| AgentError::Extraction(error.to_string())),
        AgentDocumentFormat::Xlsx => officemd_xlsx::extract_tables_ir(bytes)
            .map(|_| ())
            .map_err(|error| AgentError::Extraction(error.to_string())),
        AgentDocumentFormat::Pptx => officemd_pptx::extract_ir(bytes)
            .map(|_| ())
            .map_err(|error| AgentError::Extraction(error.to_string())),
        AgentDocumentFormat::Csv | AgentDocumentFormat::Pdf => Err(
            AgentError::UnsupportedCapability(format!("{format} mutation is not supported")),
        ),
    }
}

fn write_atomic(output: &Path, bytes: &[u8], request_id: &str) -> AgentResult<()> {
    let parent = output.parent().unwrap_or_else(|| Path::new("."));
    let file_name = output
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("artifact");
    let safe_request_id = request_id
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || matches!(character, '-' | '_') {
                character
            } else {
                '_'
            }
        })
        .collect::<String>();
    let prefix = format!(".{file_name}.officemd-{safe_request_id}-");
    let mut temporary = tempfile::Builder::new()
        .prefix(&prefix)
        .tempfile_in(parent)
        .map_err(|source| AgentError::Write {
            path: parent.display().to_string(),
            source,
        })?;
    temporary
        .write_all(bytes)
        .and_then(|()| temporary.as_file().sync_all())
        .map_err(|source| AgentError::Write {
            path: temporary.path().display().to_string(),
            source,
        })?;
    temporary
        .persist_noclobber(output)
        .map_err(|error| AgentError::Write {
            path: output.display().to_string(),
            source: error.error,
        })?;
    Ok(())
}

fn write_sidecar(sidecar: &Path, report: &ApplyPatchReport, request_id: &str) -> AgentResult<()> {
    let bytes = serde_json::to_vec_pretty(report)?;
    write_atomic(sidecar, &bytes, request_id)
}
