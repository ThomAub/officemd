//! Agent-facing OfficeMD artifact services.

pub mod artifact;
pub mod capability;
pub mod diagnostic;
pub mod error;
pub mod inspect;
pub mod locator;
pub mod patch_plan;
pub mod render;
pub mod verify;

use std::path::Path;

pub use artifact::{AgentDocumentFormat, ArtifactFingerprint, ArtifactRef};
pub use capability::{
    ArtifactCapabilityReport, ArtifactRisk, MutationKind, RenderBackendKind, RenderCapability,
    RenderUnavailableReason, VerificationCheckKind,
};
pub use diagnostic::{Diagnostic, DiagnosticSeverity};
pub use error::{AgentError, AgentResult};
pub use inspect::{
    CellPayload, InspectQuery, InspectRequest, InspectionFinding, InspectionPayload,
    InspectionReport, PdfPageInclude, XlsxRangeInclude,
};
pub use locator::{ArtifactLocator, DocxPartLocator, PdfBounds, XlsxCellLocator, XlsxSheetLocator};
pub use render::{ArtifactRenderer, PageSelection, RenderReport, RenderRequest, RenderScale};
pub use verify::{CheckReport, VerificationReport, VerificationStatus, VerifyRequest};

#[derive(Debug, Clone, Default)]
pub struct AgentService;

impl AgentService {
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    pub fn probe_path(&self, input: &Path) -> AgentResult<InspectionReport> {
        let request = InspectRequest {
            input: input.to_path_buf(),
            format: None,
            query: InspectQuery::DocumentSummary,
        };
        self.inspect(&request)
    }

    pub fn inspect(&self, request: &InspectRequest) -> AgentResult<InspectionReport> {
        inspect::inspect(request)
    }
}
