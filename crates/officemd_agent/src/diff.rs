use std::path::PathBuf;

use serde::{Deserialize, Serialize};

use crate::{
    AgentDocumentFormat, ArtifactLocator, ArtifactRef, Diagnostic,
    artifact::{fingerprint_bytes, read_artifact, resolve_artifact},
    error::{AgentError, AgentResult},
    render::{ArtifactRenderer, RenderReport, RenderRequest, RenderScale},
};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct DiffArtifactRequest {
    pub left: PathBuf,
    pub right: PathBuf,
    pub semantic: bool,
    pub rendered: bool,
    pub render_output_dir: Option<PathBuf>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ArtifactDiffReport {
    pub schema_version: u32,
    pub left: ArtifactRef,
    pub right: ArtifactRef,
    pub equal: bool,
    pub semantic: Option<SemanticDiffReport>,
    pub visual: Option<VisualDiffReport>,
    pub diagnostics: Vec<Diagnostic>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SemanticDiffReport {
    pub equal: bool,
    pub left_projection_sha256: String,
    pub right_projection_sha256: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VisualDiffReport {
    pub equal: bool,
    pub images: Vec<VisualImageDiff>,
    pub left_evidence: RenderReport,
    pub right_evidence: RenderReport,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VisualImageDiff {
    pub image_index: u32,
    pub left_locator: Option<ArtifactLocator>,
    pub right_locator: Option<ArtifactLocator>,
    pub pixel_match: bool,
}

pub fn diff_artifacts(
    request: &DiffArtifactRequest,
    renderer: &(dyn ArtifactRenderer + Send + Sync),
) -> AgentResult<ArtifactDiffReport> {
    if !request.semantic && !request.rendered {
        return Err(AgentError::InvalidRequest(
            "artifact diff requires semantic, rendered, or both".to_string(),
        ));
    }
    let left_bytes = read_artifact(&request.left)?;
    let right_bytes = read_artifact(&request.right)?;
    let left = resolve_artifact(&request.left, &left_bytes, None)?;
    let right = resolve_artifact(&request.right, &right_bytes, None)?;

    let semantic = request
        .semantic
        .then(|| semantic_diff(&left, &left_bytes, &right, &right_bytes))
        .transpose()?;
    let visual = request
        .rendered
        .then(|| visual_diff(request, renderer))
        .transpose()?;
    let equal = semantic.as_ref().is_none_or(|report| report.equal)
        && visual.as_ref().is_none_or(|report| report.equal);

    Ok(ArtifactDiffReport {
        schema_version: crate::AGENT_SCHEMA_VERSION,
        left,
        right,
        equal,
        semantic,
        visual,
        diagnostics: Vec::new(),
    })
}

fn semantic_diff(
    left: &ArtifactRef,
    left_bytes: &[u8],
    right: &ArtifactRef,
    right_bytes: &[u8],
) -> AgentResult<SemanticDiffReport> {
    let left_projection = semantic_projection(left.format, left_bytes)?;
    let right_projection = semantic_projection(right.format, right_bytes)?;
    Ok(SemanticDiffReport {
        equal: left.format == right.format && left_projection == right_projection,
        left_projection_sha256: fingerprint_bytes(&left_projection).sha256,
        right_projection_sha256: fingerprint_bytes(&right_projection).sha256,
    })
}

fn semantic_projection(format: AgentDocumentFormat, bytes: &[u8]) -> AgentResult<Vec<u8>> {
    let document = match format {
        AgentDocumentFormat::Docx => officemd_docx::extract_ir(bytes)
            .map_err(|error| AgentError::Extraction(error.to_string()))?,
        AgentDocumentFormat::Xlsx => officemd_xlsx::extract_tables_ir(bytes)
            .map_err(|error| AgentError::Extraction(error.to_string()))?,
        AgentDocumentFormat::Csv => officemd_csv::extract_tables_ir(bytes)
            .map_err(|error| AgentError::Extraction(error.to_string()))?,
        AgentDocumentFormat::Pptx => officemd_pptx::extract_ir(bytes)
            .map_err(|error| AgentError::Extraction(error.to_string()))?,
        AgentDocumentFormat::Pdf => officemd_pdf::extract_ir(bytes)
            .map_err(|error| AgentError::Extraction(error.to_string()))?,
    };
    serde_json::to_vec(&document).map_err(AgentError::from)
}

fn visual_diff(
    request: &DiffArtifactRequest,
    renderer: &(dyn ArtifactRenderer + Send + Sync),
) -> AgentResult<VisualDiffReport> {
    let root = request.render_output_dir.as_ref().ok_or_else(|| {
        AgentError::InvalidRequest("rendered diff requires render_output_dir".to_string())
    })?;
    let left_evidence = renderer.render(&RenderRequest {
        input: request.left.clone(),
        output_dir: root.join("left"),
        pages_or_slides: None,
        scale: RenderScale::Screen,
    })?;
    let right_evidence = renderer.render(&RenderRequest {
        input: request.right.clone(),
        output_dir: root.join("right"),
        pages_or_slides: None,
        scale: RenderScale::Screen,
    })?;
    let image_count = left_evidence.images.len().max(right_evidence.images.len());
    let mut images = Vec::with_capacity(image_count);
    for index in 0..image_count {
        let left_image = left_evidence.images.get(index);
        let right_image = right_evidence.images.get(index);
        let pixel_match = match (left_image, right_image) {
            (Some(left), Some(right))
                if left.pixel_width == right.pixel_width
                    && left.pixel_height == right.pixel_height =>
            {
                read_png_pixels(&left.path)? == read_png_pixels(&right.path)?
            }
            _ => false,
        };
        images.push(VisualImageDiff {
            image_index: u32::try_from(index).unwrap_or(u32::MAX),
            left_locator: left_image.map(|image| image.locator.clone()),
            right_locator: right_image.map(|image| image.locator.clone()),
            pixel_match,
        });
    }
    let equal = images.iter().all(|image| image.pixel_match);
    Ok(VisualDiffReport {
        equal,
        images,
        left_evidence,
        right_evidence,
    })
}

fn read_png_pixels(path: &std::path::Path) -> AgentResult<Vec<u8>> {
    let file = std::fs::File::open(path).map_err(|source| AgentError::Read {
        path: path.display().to_string(),
        source,
    })?;
    let mut decoder = png::Decoder::new(std::io::BufReader::new(file));
    decoder.set_transformations(png::Transformations::EXPAND | png::Transformations::STRIP_16);
    let mut reader = decoder
        .read_info()
        .map_err(|error| AgentError::RenderBackendFailed(error.to_string()))?;
    let mut pixels = vec![0; reader.output_buffer_size()];
    let info = reader
        .next_frame(&mut pixels)
        .map_err(|error| AgentError::RenderBackendFailed(error.to_string()))?;
    pixels.truncate(info.buffer_size());
    Ok(pixels)
}
