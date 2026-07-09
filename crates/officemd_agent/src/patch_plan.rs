use std::path::PathBuf;

use serde::{Deserialize, Serialize};

use crate::locator::{ArtifactLocator, XlsxCellLocator, XlsxSheetLocator};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ApplyPatchRequest {
    pub input: PathBuf,
    pub output: PathBuf,
    pub expected_source_sha256: String,
    pub patch: ArtifactPatchPlan,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ArtifactPatchPlan {
    pub version: PatchPlanVersion,
    pub request_id: String,
    pub operations: Vec<PatchOperation>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum PatchPlanVersion {
    V1,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum PatchOperation {
    ReplaceText {
        target: ArtifactLocator,
        expected_text: String,
        replacement: String,
        preserve_formatting: bool,
    },
    SetXlsxCellValue {
        target: XlsxCellLocator,
        value: CellValue,
    },
    SetXlsxCellFormula {
        target: XlsxCellLocator,
        formula: String,
    },
    RenameXlsxSheet {
        target: XlsxSheetLocator,
        new_name: String,
    },
    ReplacePptxShapeText {
        target: ArtifactLocator,
        expected_text: String,
        replacement: String,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(untagged)]
pub enum CellValue {
    Text(String),
    Number(f64),
    Bool(bool),
    Blank,
}
