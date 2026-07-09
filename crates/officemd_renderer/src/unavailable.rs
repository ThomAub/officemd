use officemd_agent::{
    ArtifactRenderer, RenderCapability, RenderReport, RenderRequest,
    capability::RenderUnavailableReason,
    error::{AgentError, AgentResult},
};

#[derive(Debug, Clone, Default)]
pub struct UnavailableRenderer;

impl ArtifactRenderer for UnavailableRenderer {
    fn capability(&self) -> RenderCapability {
        RenderCapability::Unavailable {
            reason: RenderUnavailableReason::BackendNotConfigured,
        }
    }

    fn render(&self, _request: &RenderRequest) -> AgentResult<RenderReport> {
        Err(AgentError::RenderUnavailable(
            "no renderer backend is configured".to_string(),
        ))
    }
}
