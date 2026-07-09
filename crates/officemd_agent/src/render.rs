use std::path::PathBuf;

use serde::{Deserialize, Serialize};

use crate::{
    ArtifactRef, Diagnostic,
    capability::{RenderBackendKind, RenderCapability},
    error::AgentResult,
    locator::ArtifactLocator,
};

pub trait ArtifactRenderer {
    fn capability(&self) -> RenderCapability;
    fn render(&self, request: &RenderRequest) -> AgentResult<RenderReport>;
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct RenderRequest {
    pub input: PathBuf,
    pub output_dir: PathBuf,
    pub pages_or_slides: Option<PageSelection>,
    pub scale: RenderScale,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderReport {
    pub artifact: ArtifactRef,
    pub backend: RenderBackendKind,
    pub images: Vec<RenderedImage>,
    pub diagnostics: Vec<Diagnostic>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct RenderedImage {
    pub locator: ArtifactLocator,
    pub path: PathBuf,
    pub pixel_width: u32,
    pub pixel_height: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct PageSelection {
    pub start: u32,
    pub end: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum RenderScale {
    Screen,
    Print,
}
