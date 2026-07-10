//! Agent-facing OfficeMD artifact services.

pub mod apply;
pub mod artifact;
pub mod capability;
pub mod diagnostic;
pub mod diff;
pub mod error;
pub mod inspect;
pub mod locator;
pub mod patch_plan;
pub mod render;
pub mod verify;

use std::{path::Path, sync::Arc};

pub use apply::{ApplyPatchReport, ApplyPatchStatus, OperationReport, OperationStatus};
pub use artifact::{AgentDocumentFormat, ArtifactFingerprint, ArtifactRef};
pub use capability::{
    ArtifactCapabilityReport, ArtifactRisk, MutationKind, RenderBackendKind, RenderCapability,
    RenderUnavailableReason, VerificationCheckKind,
};
pub use diagnostic::{Diagnostic, DiagnosticSeverity};
pub use diff::{
    ArtifactDiffReport, DiffArtifactRequest, SemanticDiffReport, VisualDiffReport, VisualImageDiff,
};
pub use error::{AgentError, AgentResult};
pub use inspect::{
    CellPayload, InspectQuery, InspectRequest, InspectionFinding, InspectionPayload,
    InspectionReport, PdfPageInclude, XlsxRangeInclude,
};
pub use locator::{ArtifactLocator, DocxPartLocator, PdfBounds, XlsxCellLocator, XlsxSheetLocator};
pub use patch_plan::{
    ApplyPatchRequest, ArtifactPatchPlan, CellValue, PatchOperation, PatchPlanVersion,
};
pub use render::{
    ArtifactRenderer, PageSelection, RenderReport, RenderRequest, RenderScale, RenderedImage,
};
pub use verify::{CheckReport, VerificationReport, VerificationStatus, VerifyRequest};

pub const AGENT_SCHEMA_VERSION: u32 = 1;

#[derive(Clone)]
pub struct AgentService {
    renderer: Arc<dyn ArtifactRenderer + Send + Sync>,
}

impl std::fmt::Debug for AgentService {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("AgentService")
            .finish_non_exhaustive()
    }
}

impl Default for AgentService {
    fn default() -> Self {
        Self::new()
    }
}

impl AgentService {
    #[must_use]
    pub fn new() -> Self {
        Self {
            renderer: Arc::new(render::UnavailableRenderer),
        }
    }

    #[must_use]
    pub fn with_renderer(renderer: impl ArtifactRenderer + Send + Sync + 'static) -> Self {
        Self {
            renderer: Arc::new(renderer),
        }
    }

    pub fn probe_path(&self, input: &Path) -> AgentResult<InspectionReport> {
        self.probe_path_as(input, None)
    }

    pub fn probe_path_as(
        &self,
        input: &Path,
        format: Option<AgentDocumentFormat>,
    ) -> AgentResult<InspectionReport> {
        let mut report = inspect::probe(input.to_path_buf(), format)?;
        report.capability =
            capability::capability_for(report.artifact.format, self.renderer.capability());
        Ok(report)
    }

    pub fn inspect(&self, request: &InspectRequest) -> AgentResult<InspectionReport> {
        let mut report = inspect::inspect(request)?;
        report.capability =
            capability::capability_for(report.artifact.format, self.renderer.capability());
        Ok(report)
    }

    pub fn apply_patch(&self, request: &ApplyPatchRequest) -> AgentResult<ApplyPatchReport> {
        apply::apply_patch(request)
    }

    pub fn render(&self, request: &RenderRequest) -> AgentResult<RenderReport> {
        self.renderer.render(request)
    }

    pub fn verify(&self, request: &VerifyRequest) -> AgentResult<VerificationReport> {
        verify::verify(request, self.renderer.as_ref())
    }

    pub fn diff_artifacts(&self, request: &DiffArtifactRequest) -> AgentResult<ArtifactDiffReport> {
        diff::diff_artifacts(request, self.renderer.as_ref())
    }
}
