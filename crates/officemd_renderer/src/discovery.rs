use officemd_agent::{
    AgentDocumentFormat, ArtifactRenderer, RenderCapability, RenderReport, RenderRequest,
    capability::{RenderBackendKind, RenderUnavailableReason},
    error::{AgentError, AgentResult},
};

use crate::{LibreOfficeRenderer, PopplerRenderer};

#[derive(Debug, Clone)]
pub struct SystemRenderer {
    poppler: Option<PopplerRenderer>,
    libreoffice: Option<LibreOfficeRenderer>,
}

impl SystemRenderer {
    #[must_use]
    pub fn discover() -> Self {
        Self {
            poppler: PopplerRenderer::discover(),
            libreoffice: LibreOfficeRenderer::discover(),
        }
    }
}

impl Default for SystemRenderer {
    fn default() -> Self {
        Self::discover()
    }
}

impl ArtifactRenderer for SystemRenderer {
    fn capability(&self) -> RenderCapability {
        match (&self.poppler, &self.libreoffice) {
            (Some(_), Some(_)) => RenderCapability::Available {
                backend: RenderBackendKind::LibreOffice,
                formats: vec![
                    AgentDocumentFormat::Pdf,
                    AgentDocumentFormat::Docx,
                    AgentDocumentFormat::Xlsx,
                    AgentDocumentFormat::Pptx,
                ],
            },
            (Some(_), None) => RenderCapability::Available {
                backend: RenderBackendKind::Poppler,
                formats: vec![AgentDocumentFormat::Pdf],
            },
            _ => RenderCapability::Unavailable {
                reason: RenderUnavailableReason::MissingExecutable,
            },
        }
    }

    fn render(&self, request: &RenderRequest) -> AgentResult<RenderReport> {
        let artifact = crate::artifact_ref_for_path(&request.input)?;
        match artifact.format {
            AgentDocumentFormat::Pdf => self
                .poppler
                .as_ref()
                .ok_or_else(|| AgentError::RenderUnavailable("pdftoppm was not found".to_string()))?
                .render_pdf_request(request, artifact),
            AgentDocumentFormat::Docx | AgentDocumentFormat::Xlsx | AgentDocumentFormat::Pptx => {
                let poppler = self.poppler.as_ref().ok_or_else(|| {
                    AgentError::RenderUnavailable("pdftoppm was not found".to_string())
                })?;
                self.libreoffice
                    .as_ref()
                    .ok_or_else(|| {
                        AgentError::RenderUnavailable(
                            "LibreOffice soffice was not found".to_string(),
                        )
                    })?
                    .render_office_request(request, artifact, poppler)
            }
            AgentDocumentFormat::Csv => Err(AgentError::UnsupportedCapability(
                "CSV visual rendering is not supported".to_string(),
            )),
        }
    }
}

#[must_use]
pub fn discover() -> RenderCapability {
    SystemRenderer::discover().capability()
}

#[must_use]
pub fn backend_name() -> RenderBackendKind {
    match discover() {
        RenderCapability::Available { backend, .. } => backend,
        RenderCapability::Unavailable { .. } => RenderBackendKind::Unavailable,
    }
}
