use officemd_agent::{
    RenderCapability,
    capability::{RenderBackendKind, RenderUnavailableReason},
};

#[must_use]
pub fn discover() -> RenderCapability {
    RenderCapability::Unavailable {
        reason: RenderUnavailableReason::BackendNotConfigured,
    }
}

#[must_use]
pub fn backend_name() -> RenderBackendKind {
    RenderBackendKind::Unavailable
}
