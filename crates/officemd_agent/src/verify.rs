use std::path::PathBuf;

use serde::{Deserialize, Serialize};

use crate::{
    ArtifactRef, Diagnostic, capability::VerificationCheckKind, locator::ArtifactLocator,
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
