use std::{
    collections::BTreeMap,
    io::{Cursor, Read, Write},
    path::Path,
    sync::LazyLock,
};

use officemd_core::{
    PptxPatch, PptxTextScope, ScopedPptxReplace, ScopedXlsxReplace, TextReplace, XlsxPatch,
    XlsxSheetRename, XlsxTextScope,
};
use regex::Regex;
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
    let mut reports = Vec::new();
    let mut patched_bytes = source_bytes.to_vec();
    for (index, operation) in operations.iter().enumerate() {
        match operation {
            PatchOperation::ReplaceText {
                target,
                expected_text,
                replacement,
                preserve_formatting: _,
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
) -> AgentResult<usize> {
    let (part, updated, replacements) = {
        let parts = read_zip_parts(content)?;
        let part = docx_locator_part(target)?;
        let xml = parts
            .get(&part)
            .ok_or_else(|| AgentError::InvalidRequest(format!("DOCX part not found: {part}")))?;
        let xml = String::from_utf8_lossy(xml);
        let (updated, replacements) = match target {
            ArtifactLocator::DocxParagraph {
                paragraph_index, ..
            } => replace_docx_paragraph(&xml, *paragraph_index, expected_text, replacement),
            ArtifactLocator::DocxTableCell {
                table_index,
                row_index,
                column_index,
                ..
            } => replace_docx_table_cell(
                &xml,
                *table_index,
                *row_index,
                *column_index,
                expected_text,
                replacement,
            ),
            _ => {
                return Err(AgentError::InvalidRequest(
                    "target must be a DOCX locator".into(),
                ));
            }
        };
        (part, updated, replacements)
    };

    if replacements > 0 {
        replace_zip_part(content, &part, updated.as_bytes())?;
    }
    Ok(replacements)
}

fn docx_locator_part(target: &ArtifactLocator) -> AgentResult<String> {
    let part = match target {
        ArtifactLocator::DocxParagraph { part, .. }
        | ArtifactLocator::DocxTableCell { part, .. } => part.part.as_str(),
        _ => {
            return Err(AgentError::InvalidRequest(
                "target must be a DOCX locator".into(),
            ));
        }
    };
    Ok(match part {
        "body" | "document" => "word/document.xml".to_string(),
        "footnotes" => "word/footnotes.xml".to_string(),
        "endnotes" => "word/endnotes.xml".to_string(),
        path if path.starts_with("word/") => path.to_string(),
        name => format!("word/{name}.xml"),
    })
}

fn replace_docx_paragraph(
    xml: &str,
    paragraph_index: u32,
    expected_text: &str,
    replacement: &str,
) -> (String, usize) {
    let mut visible_index = 0u32;
    replace_indexed_match(xml, &WORD_PARAGRAPH_RE, |block| {
        if docx_text(block).trim().is_empty() {
            return None;
        }
        let is_target = visible_index == paragraph_index;
        visible_index = visible_index.saturating_add(1);
        is_target.then(|| replace_docx_text_nodes(block, expected_text, replacement))
    })
}

fn replace_docx_table_cell(
    xml: &str,
    table_index: u32,
    row_index: u32,
    column_index: u32,
    expected_text: &str,
    replacement: &str,
) -> (String, usize) {
    let mut current_table = 0u32;
    replace_indexed_match(xml, &WORD_TABLE_RE, |table| {
        let is_target_table = current_table == table_index;
        current_table = current_table.saturating_add(1);
        if !is_target_table {
            return None;
        }
        let mut current_row = 0u32;
        Some(replace_indexed_match(table, &WORD_ROW_RE, |row| {
            let is_target_row = current_row == row_index;
            current_row = current_row.saturating_add(1);
            if !is_target_row {
                return None;
            }
            let mut current_col = 0u32;
            Some(replace_indexed_match(row, &WORD_CELL_RE, |cell| {
                let is_target_col = current_col == column_index;
                current_col = current_col.saturating_add(1);
                is_target_col.then(|| replace_docx_text_nodes(cell, expected_text, replacement))
            }))
        }))
    })
}

fn replace_indexed_match<F>(xml: &str, regex: &Regex, mut replacer: F) -> (String, usize)
where
    F: FnMut(&str) -> Option<(String, usize)>,
{
    let mut output = String::with_capacity(xml.len());
    let mut last = 0usize;
    let mut replacements = 0usize;
    for mat in regex.find_iter(xml) {
        output.push_str(&xml[last..mat.start()]);
        if replacements == 0
            && let Some((updated, count)) = replacer(mat.as_str())
        {
            output.push_str(&updated);
            replacements = replacements.saturating_add(count);
        } else {
            output.push_str(mat.as_str());
        }
        last = mat.end();
    }
    output.push_str(&xml[last..]);
    (output, replacements)
}

fn replace_docx_text_nodes(xml: &str, expected_text: &str, replacement: &str) -> (String, usize) {
    let mut output = String::with_capacity(xml.len());
    let mut last = 0usize;
    let mut replacements = 0usize;
    for captures in WORD_TEXT_NODE_RE.captures_iter(xml) {
        let Some(mat) = captures.get(0) else {
            continue;
        };
        let Some(text) = captures.get(1) else {
            continue;
        };
        output.push_str(&xml[last..text.start()]);
        let count = text.as_str().match_indices(expected_text).count();
        if count == 0 {
            output.push_str(text.as_str());
        } else {
            output.push_str(&text.as_str().replace(expected_text, replacement));
            replacements = replacements.saturating_add(count);
        }
        output.push_str(&xml[text.end()..mat.end()]);
        last = mat.end();
    }
    output.push_str(&xml[last..]);
    (output, replacements)
}

fn docx_text(xml: &str) -> String {
    WORD_TEXT_NODE_RE
        .captures_iter(xml)
        .filter_map(|captures| captures.get(1).map(|text| text.as_str()))
        .collect()
}

fn read_zip_parts(content: &[u8]) -> AgentResult<BTreeMap<String, Vec<u8>>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(content))
        .map_err(|e| AgentError::Extraction(e.to_string()))?;
    let mut parts = BTreeMap::new();
    for index in 0..archive.len() {
        let mut file = archive
            .by_index(index)
            .map_err(|e| AgentError::Extraction(e.to_string()))?;
        if file.is_dir() {
            continue;
        }
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)
            .map_err(|e| AgentError::Extraction(e.to_string()))?;
        parts.insert(file.name().to_string(), bytes);
    }
    Ok(parts)
}

fn replace_zip_part(content: &mut Vec<u8>, part: &str, replacement: &[u8]) -> AgentResult<()> {
    let mut parts = read_zip_parts(content)?;
    parts.insert(part.to_string(), replacement.to_vec());
    let mut rewritten = Vec::new();
    {
        let mut writer = zip::ZipWriter::new(Cursor::new(&mut rewritten));
        let options: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();
        for (name, bytes) in parts {
            writer
                .start_file(name, options)
                .map_err(|e| AgentError::Extraction(e.to_string()))?;
            writer
                .write_all(&bytes)
                .map_err(|e| AgentError::Extraction(e.to_string()))?;
        }
        writer
            .finish()
            .map_err(|e| AgentError::Extraction(e.to_string()))?;
    }
    *content = rewritten;
    Ok(())
}

static WORD_PARAGRAPH_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:p(?:\s+[^>]*)?>.*?</w:p>").unwrap());
static WORD_TABLE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:tbl(?:\s+[^>]*)?>.*?</w:tbl>").unwrap());
static WORD_ROW_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:tr(?:\s+[^>]*)?>.*?</w:tr>").unwrap());
static WORD_CELL_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:tc(?:\s+[^>]*)?>.*?</w:tc>").unwrap());
static WORD_TEXT_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<w:t(?:\s+[^>]*)?>(.*?)</w:t>").unwrap());

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
