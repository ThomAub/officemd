use std::{
    ffi::OsString,
    path::{Path, PathBuf},
    process::Command,
};

use officemd_agent::{
    ArtifactRef, ArtifactRenderer, RenderCapability, RenderReport, RenderRequest, RenderedImage,
    capability::RenderBackendKind,
    error::{AgentError, AgentResult},
    locator::ArtifactLocator,
};

#[derive(Debug, Clone)]
pub struct PopplerRenderer {
    executable: PathBuf,
}

impl PopplerRenderer {
    #[must_use]
    pub fn discover() -> Option<Self> {
        find_executable("pdftoppm").map(|executable| Self { executable })
    }

    pub(crate) fn render_pdf_request(
        &self,
        request: &RenderRequest,
        artifact: ArtifactRef,
    ) -> AgentResult<RenderReport> {
        self.render_pdf(
            &request.input,
            artifact,
            &request.output_dir,
            request
                .pages_or_slides
                .as_ref()
                .map(|selection| (selection.start, selection.end)),
        )
    }

    pub(crate) fn render_pdf(
        &self,
        pdf_path: &Path,
        artifact: ArtifactRef,
        output_dir: &Path,
        page_selection: Option<(u32, u32)>,
    ) -> AgentResult<RenderReport> {
        std::fs::create_dir_all(output_dir).map_err(|source| AgentError::Read {
            path: output_dir.display().to_string(),
            source,
        })?;
        let prefix = output_dir.join(render_prefix(pdf_path));
        let mut command = Command::new(&self.executable);
        command.arg("-png").arg("-r").arg("144");
        if let Some((start, end)) = page_selection {
            command
                .arg("-f")
                .arg(start.to_string())
                .arg("-l")
                .arg(end.to_string());
        }
        command.arg(pdf_path).arg(&prefix);
        let output = command.output().map_err(|source| AgentError::Read {
            path: self.executable.display().to_string(),
            source,
        })?;
        if !output.status.success() {
            return Err(AgentError::RenderUnavailable(format!(
                "pdftoppm failed: {}",
                safe_stderr(&output.stderr)
            )));
        }

        let mut images = find_rendered_pngs(output_dir, &prefix)?;
        if images.is_empty() {
            return Err(AgentError::RenderUnavailable(
                "pdftoppm did not produce any PNG files".to_string(),
            ));
        }
        images.sort();
        let rendered = images
            .into_iter()
            .enumerate()
            .map(|(index, path)| {
                let (pixel_width, pixel_height) = png_dimensions(&path)?;
                let page_number = page_selection
                    .map(|(start, _)| start)
                    .unwrap_or(1)
                    .saturating_add(u32::try_from(index).unwrap_or(u32::MAX));
                Ok(RenderedImage {
                    locator: ArtifactLocator::PdfPage { page_number },
                    path,
                    pixel_width,
                    pixel_height,
                })
            })
            .collect::<AgentResult<Vec<_>>>()?;

        Ok(RenderReport {
            artifact,
            backend: RenderBackendKind::Poppler,
            images: rendered,
            diagnostics: Vec::new(),
        })
    }
}

impl ArtifactRenderer for PopplerRenderer {
    fn capability(&self) -> RenderCapability {
        RenderCapability::Available {
            backend: RenderBackendKind::Poppler,
            formats: vec![officemd_agent::AgentDocumentFormat::Pdf],
        }
    }

    fn render(&self, request: &RenderRequest) -> AgentResult<RenderReport> {
        let artifact = crate::artifact_ref_for_pdf_path(&request.input)?;
        self.render_pdf_request(request, artifact)
    }
}

fn find_executable(name: &str) -> Option<PathBuf> {
    let path = std::env::var_os("PATH")?;
    std::env::split_paths(&path)
        .map(|dir| dir.join(name))
        .find(|candidate| candidate.is_file())
}

fn render_prefix(path: &Path) -> OsString {
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("render");
    OsString::from(format!("{stem}-page"))
}

fn find_rendered_pngs(output_dir: &Path, prefix: &Path) -> AgentResult<Vec<PathBuf>> {
    let prefix_name = prefix
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_string();
    let mut paths = Vec::new();
    for entry in std::fs::read_dir(output_dir).map_err(|source| AgentError::Read {
        path: output_dir.display().to_string(),
        source,
    })? {
        let entry = entry.map_err(|source| AgentError::Read {
            path: output_dir.display().to_string(),
            source,
        })?;
        let path = entry.path();
        let is_match = path
            .file_name()
            .and_then(|name| name.to_str())
            .is_some_and(|name| name.starts_with(&prefix_name) && name.ends_with(".png"));
        if is_match {
            paths.push(path);
        }
    }
    Ok(paths)
}

fn png_dimensions(path: &Path) -> AgentResult<(u32, u32)> {
    let bytes = std::fs::read(path).map_err(|source| AgentError::Read {
        path: path.display().to_string(),
        source,
    })?;
    if bytes.len() < 24 || &bytes[..8] != b"\x89PNG\r\n\x1a\n" {
        return Err(AgentError::RenderUnavailable(format!(
            "rendered image is not a PNG: {}",
            path.display()
        )));
    }
    let width = u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]);
    let height = u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]);
    Ok((width, height))
}

fn safe_stderr(stderr: &[u8]) -> String {
    let text = String::from_utf8_lossy(stderr);
    let trimmed = text.trim();
    if trimmed.len() > 400 {
        format!("{}...", &trimmed[..400])
    } else {
        trimmed.to_string()
    }
}
