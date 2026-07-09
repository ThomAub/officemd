use std::{
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

use officemd_agent::{
    AgentDocumentFormat, ArtifactRef, ArtifactRenderer, RenderCapability, RenderReport,
    RenderRequest,
    capability::RenderBackendKind,
    error::{AgentError, AgentResult},
};

use crate::PopplerRenderer;

#[derive(Debug, Clone)]
pub struct LibreOfficeRenderer {
    executable: PathBuf,
}

impl LibreOfficeRenderer {
    #[must_use]
    pub fn discover() -> Option<Self> {
        find_executable("soffice")
            .or_else(|| find_executable("libreoffice"))
            .map(|executable| Self { executable })
    }

    pub(crate) fn render_office_request(
        &self,
        request: &RenderRequest,
        artifact: ArtifactRef,
        poppler: &PopplerRenderer,
    ) -> AgentResult<RenderReport> {
        let temp = TempDir::new("officemd-libreoffice")?;
        let profile = temp.path.join("profile");
        let outdir = temp.path.join("out");
        std::fs::create_dir_all(&profile).map_err(|source| AgentError::Read {
            path: profile.display().to_string(),
            source,
        })?;
        std::fs::create_dir_all(&outdir).map_err(|source| AgentError::Read {
            path: outdir.display().to_string(),
            source,
        })?;

        let output = Command::new(&self.executable)
            .arg("--headless")
            .arg("--nologo")
            .arg("--nodefault")
            .arg("--nofirststartwizard")
            .arg(format!(
                "-env:UserInstallation={}",
                file_url_for_path(&profile)
            ))
            .arg("--convert-to")
            .arg("pdf")
            .arg("--outdir")
            .arg(&outdir)
            .arg(&request.input)
            .output()
            .map_err(|source| AgentError::Read {
                path: self.executable.display().to_string(),
                source,
            })?;

        let pdf = find_converted_pdf(&outdir)?;
        if pdf.is_none() || !output.status.success() {
            return Err(AgentError::RenderUnavailable(format!(
                "LibreOffice conversion failed: {}",
                safe_stderr(&output.stderr)
            )));
        }
        let pdf = pdf.expect("checked is_some");
        let mut report = poppler.render_pdf(
            &pdf,
            artifact,
            &request.output_dir,
            request
                .pages_or_slides
                .as_ref()
                .map(|selection| (selection.start, selection.end)),
        )?;
        report.backend = RenderBackendKind::LibreOffice;
        Ok(report)
    }
}

impl ArtifactRenderer for LibreOfficeRenderer {
    fn capability(&self) -> RenderCapability {
        RenderCapability::Available {
            backend: RenderBackendKind::LibreOffice,
            formats: vec![
                AgentDocumentFormat::Docx,
                AgentDocumentFormat::Xlsx,
                AgentDocumentFormat::Pptx,
            ],
        }
    }

    fn render(&self, _request: &RenderRequest) -> AgentResult<RenderReport> {
        Err(AgentError::RenderUnavailable(
            "LibreOffice rendering requires a Poppler renderer for PDF rasterization".to_string(),
        ))
    }
}

struct TempDir {
    path: PathBuf,
}

impl TempDir {
    fn new(prefix: &str) -> AgentResult<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map_or(0, |duration| duration.as_nanos());
        let path = std::env::temp_dir().join(format!("{prefix}-{}-{nonce}", std::process::id()));
        std::fs::create_dir_all(&path).map_err(|source| AgentError::Read {
            path: path.display().to_string(),
            source,
        })?;
        Ok(Self { path })
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.path);
    }
}

fn find_executable(name: &str) -> Option<PathBuf> {
    let path = std::env::var_os("PATH")?;
    std::env::split_paths(&path)
        .map(|dir| dir.join(name))
        .find(|candidate| candidate.is_file())
}

fn file_url_for_path(path: &Path) -> String {
    let mut value = path.to_string_lossy().replace(' ', "%20");
    if !value.starts_with('/') {
        value.insert(0, '/');
    }
    format!("file://{value}")
}

fn find_converted_pdf(outdir: &Path) -> AgentResult<Option<PathBuf>> {
    let mut pdfs = Vec::new();
    for entry in std::fs::read_dir(outdir).map_err(|source| AgentError::Read {
        path: outdir.display().to_string(),
        source,
    })? {
        let entry = entry.map_err(|source| AgentError::Read {
            path: outdir.display().to_string(),
            source,
        })?;
        let path = entry.path();
        if path
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("pdf"))
        {
            pdfs.push(path);
        }
    }
    pdfs.sort();
    Ok(pdfs.into_iter().next())
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
