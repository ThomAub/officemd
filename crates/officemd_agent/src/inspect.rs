use std::path::PathBuf;

use officemd_core::ir::{Block, Inline, OoxmlDocument, Paragraph, TableCell};
use officemd_pptx::PptxExtractOptions;
use serde::{Deserialize, Serialize};

use crate::{
    ArtifactRef,
    artifact::{AgentDocumentFormat, read_artifact, resolve_artifact},
    capability::{
        ArtifactCapabilityReport, RenderCapability, RenderUnavailableReason, capability_for,
    },
    diagnostic::Diagnostic,
    error::{AgentError, AgentResult},
    locator::{ArtifactLocator, DocxPartLocator},
};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct InspectRequest {
    pub input: PathBuf,
    pub format: Option<AgentDocumentFormat>,
    pub query: InspectQuery,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum InspectQuery {
    DocumentSummary,
    DocxOutline,
    DocxTables {
        max_tables: u32,
        max_rows_per_table: u32,
    },
    XlsxRange {
        sheet: String,
        range: String,
        include: XlsxRangeInclude,
    },
    XlsxSheets,
    PptxSlides {
        start: u32,
        end: u32,
    },
    PdfPages {
        pages: Vec<u32>,
        include: PdfPageInclude,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct XlsxRangeInclude {
    #[serde(default)]
    pub values: bool,
    #[serde(default)]
    pub formulas: bool,
    #[serde(default)]
    pub number_formats: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct PdfPageInclude {
    #[serde(default)]
    pub markdown: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InspectionReport {
    pub schema_version: u32,
    pub artifact: ArtifactRef,
    pub capability: ArtifactCapabilityReport,
    pub findings: Vec<InspectionFinding>,
    pub diagnostics: Vec<Diagnostic>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InspectionFinding {
    pub locator: ArtifactLocator,
    pub kind: String,
    pub payload: InspectionPayload,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum InspectionPayload {
    Summary {
        format: AgentDocumentFormat,
        sections: usize,
        sheets: usize,
        slides: usize,
        pages: usize,
    },
    Text {
        text: String,
    },
    Table {
        rows: Vec<Vec<String>>,
    },
    Sheet {
        name: String,
        rows: usize,
        cols: usize,
    },
    Cell {
        address: String,
        value: Option<String>,
        formula: Option<String>,
        number_format: Option<String>,
    },
    Slide {
        number: u32,
        title: Option<String>,
        has_notes: bool,
        comment_count: usize,
    },
    Shape {
        slide_number: u32,
        shape_id: u32,
        text: String,
    },
    PdfPage {
        number: u32,
        markdown: Option<String>,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct CellPayload {
    pub address: String,
    pub value: Option<String>,
    pub formula: Option<String>,
    pub number_format: Option<String>,
}

pub fn inspect(request: &InspectRequest) -> AgentResult<InspectionReport> {
    let bytes = read_artifact(&request.input)?;
    let artifact = resolve_artifact(&request.input, &bytes, request.format)?;
    let capability = capability_for(
        artifact.format,
        RenderCapability::Unavailable {
            reason: RenderUnavailableReason::BackendNotConfigured,
        },
    );
    let findings = inspect_bytes(&bytes, &artifact, &request.query)?;

    Ok(InspectionReport {
        schema_version: crate::AGENT_SCHEMA_VERSION,
        artifact,
        capability,
        findings,
        diagnostics: Vec::new(),
    })
}

pub fn probe(input: PathBuf, format: Option<AgentDocumentFormat>) -> AgentResult<InspectionReport> {
    let bytes = read_artifact(&input)?;
    let artifact = resolve_artifact(&input, &bytes, format)?;
    validate_probe_structure(artifact.format, &bytes)?;
    let capability = capability_for(
        artifact.format,
        RenderCapability::Unavailable {
            reason: RenderUnavailableReason::BackendNotConfigured,
        },
    );
    Ok(InspectionReport {
        schema_version: crate::AGENT_SCHEMA_VERSION,
        artifact,
        capability,
        findings: Vec::new(),
        diagnostics: Vec::new(),
    })
}

fn validate_probe_structure(format: AgentDocumentFormat, bytes: &[u8]) -> AgentResult<()> {
    match format {
        AgentDocumentFormat::Docx | AgentDocumentFormat::Xlsx | AgentDocumentFormat::Pptx => {
            let mut package = officemd_core::opc::OpcPackage::from_bytes(bytes)
                .map_err(|error| AgentError::Extraction(error.to_string()))?;
            let required_part = match format {
                AgentDocumentFormat::Docx => "word/document.xml",
                AgentDocumentFormat::Xlsx => "xl/workbook.xml",
                AgentDocumentFormat::Pptx => "ppt/presentation.xml",
                _ => unreachable!(),
            };
            if !package.has_part(required_part) {
                return Err(AgentError::Extraction(format!(
                    "required package part is missing: {required_part}"
                )));
            }
            Ok(())
        }
        AgentDocumentFormat::Pdf if !officemd_pdf::looks_like_pdf_header(bytes) => Err(
            AgentError::Extraction("artifact does not contain a PDF header".to_string()),
        ),
        AgentDocumentFormat::Csv | AgentDocumentFormat::Pdf => Ok(()),
    }
}

fn inspect_bytes(
    bytes: &[u8],
    artifact: &ArtifactRef,
    query: &InspectQuery,
) -> AgentResult<Vec<InspectionFinding>> {
    match query {
        InspectQuery::DocumentSummary => inspect_summary(bytes, artifact),
        InspectQuery::DocxOutline => require_format(artifact, AgentDocumentFormat::Docx)
            .and_then(|()| inspect_docx_outline(bytes)),
        InspectQuery::DocxTables {
            max_tables,
            max_rows_per_table,
        } => require_format(artifact, AgentDocumentFormat::Docx)
            .and_then(|()| inspect_docx_tables(bytes, *max_tables, *max_rows_per_table)),
        InspectQuery::XlsxSheets => require_format(artifact, AgentDocumentFormat::Xlsx)
            .and_then(|()| inspect_xlsx_sheets(bytes)),
        InspectQuery::XlsxRange {
            sheet,
            range,
            include,
        } => require_format(artifact, AgentDocumentFormat::Xlsx)
            .and_then(|()| inspect_xlsx_range(bytes, sheet, range, include)),
        InspectQuery::PptxSlides { start, end } => {
            require_format(artifact, AgentDocumentFormat::Pptx)
                .and_then(|()| inspect_pptx_slides(bytes, *start, *end))
        }
        InspectQuery::PdfPages { pages, include } => {
            require_format(artifact, AgentDocumentFormat::Pdf)
                .and_then(|()| inspect_pdf_pages(bytes, pages, include))
        }
    }
}

fn require_format(artifact: &ArtifactRef, expected: AgentDocumentFormat) -> AgentResult<()> {
    if artifact.format == expected {
        Ok(())
    } else {
        Err(AgentError::InvalidRequest(format!(
            "query requires {expected}, artifact is {}",
            artifact.format
        )))
    }
}

fn inspect_summary(bytes: &[u8], artifact: &ArtifactRef) -> AgentResult<Vec<InspectionFinding>> {
    let doc = extract_ir(bytes, artifact.format)?;
    let pages = doc.pdf.as_ref().map_or(0, |pdf| pdf.diagnostics.page_count);
    let locator = match artifact.format {
        AgentDocumentFormat::Docx => ArtifactLocator::DocxParagraph {
            part: DocxPartLocator {
                part: "document".to_string(),
            },
            paragraph_index: 0,
        },
        AgentDocumentFormat::Xlsx | AgentDocumentFormat::Csv => ArtifactLocator::XlsxSheet {
            name: doc
                .sheets
                .first()
                .map_or_else(|| "sheet".to_string(), |sheet| sheet.name.clone()),
        },
        AgentDocumentFormat::Pptx => ArtifactLocator::PptxSlide { slide_number: 1 },
        AgentDocumentFormat::Pdf => ArtifactLocator::PdfPage { page_number: 1 },
    };

    Ok(vec![InspectionFinding {
        locator,
        kind: "document_summary".to_string(),
        payload: InspectionPayload::Summary {
            format: artifact.format,
            sections: doc.sections.len(),
            sheets: doc.sheets.len(),
            slides: doc.slides.len(),
            pages,
        },
    }])
}

fn inspect_docx_outline(bytes: &[u8]) -> AgentResult<Vec<InspectionFinding>> {
    let doc =
        officemd_docx::extract_ir(bytes).map_err(|e| AgentError::Extraction(e.to_string()))?;
    let mut findings = Vec::new();
    for section in doc.sections {
        let mut paragraph_index = 0u32;
        for block in section.blocks {
            if let Block::Paragraph(paragraph) = block {
                let text = paragraph_text(&paragraph);
                if text.trim().is_empty() {
                    continue;
                }
                findings.push(InspectionFinding {
                    locator: ArtifactLocator::DocxParagraph {
                        part: DocxPartLocator {
                            part: section.name.clone(),
                        },
                        paragraph_index,
                    },
                    kind: "paragraph".to_string(),
                    payload: InspectionPayload::Text { text },
                });
                paragraph_index = paragraph_index.saturating_add(1);
            }
        }
    }
    Ok(findings)
}

fn inspect_docx_tables(
    bytes: &[u8],
    max_tables: u32,
    max_rows_per_table: u32,
) -> AgentResult<Vec<InspectionFinding>> {
    let doc =
        officemd_docx::extract_ir(bytes).map_err(|e| AgentError::Extraction(e.to_string()))?;
    let mut findings = Vec::new();
    for section in doc.sections {
        let mut table_index = 0u32;
        for block in section.blocks {
            let Block::Table(table) = block else {
                continue;
            };
            if table_index >= max_tables {
                break;
            }
            let rows = table
                .rows
                .iter()
                .take(usize::try_from(max_rows_per_table).unwrap_or(usize::MAX))
                .map(|row| row.iter().map(cell_text).collect::<Vec<_>>())
                .collect::<Vec<_>>();
            findings.push(InspectionFinding {
                locator: ArtifactLocator::DocxTableCell {
                    part: DocxPartLocator {
                        part: section.name.clone(),
                    },
                    table_index,
                    row_index: 0,
                    column_index: 0,
                },
                kind: "table".to_string(),
                payload: InspectionPayload::Table { rows },
            });
            table_index = table_index.saturating_add(1);
        }
    }
    Ok(findings)
}

fn inspect_xlsx_sheets(bytes: &[u8]) -> AgentResult<Vec<InspectionFinding>> {
    let summaries = officemd_xlsx::inspect_sheet_summaries(bytes, None)
        .map_err(|e| AgentError::Extraction(e.to_string()))?;
    Ok(summaries
        .into_iter()
        .map(|sheet| InspectionFinding {
            locator: ArtifactLocator::XlsxSheet {
                name: sheet.name.clone(),
            },
            kind: "sheet".to_string(),
            payload: InspectionPayload::Sheet {
                name: sheet.name,
                rows: sheet.rows,
                cols: sheet.cols,
            },
        })
        .collect())
}

fn inspect_xlsx_range(
    bytes: &[u8],
    sheet: &str,
    range: &str,
    include: &XlsxRangeInclude,
) -> AgentResult<Vec<InspectionFinding>> {
    let (start_row, start_col, end_row, end_col) = parse_range(range)?;
    let mut addresses = Vec::new();
    for row_idx in start_row..=end_row {
        for col_idx in start_col..=end_col {
            addresses.push(format!("{}{}", column_name(col_idx), row_idx + 1));
        }
    }
    let cells = officemd_xlsx::inspect_cells(bytes, sheet, &addresses)
        .map_err(|e| AgentError::Extraction(e.to_string()))?;

    let mut findings = Vec::new();
    for cell in cells {
        let value = include.values.then_some(cell.value).flatten();
        let formula = include.formulas.then_some(cell.formula).flatten();
        let number_format = include
            .number_formats
            .then_some(cell.number_format)
            .flatten();
        findings.push(InspectionFinding {
            locator: ArtifactLocator::XlsxCell {
                sheet: sheet.to_string(),
                address: cell.address.clone(),
            },
            kind: "cell".to_string(),
            payload: InspectionPayload::Cell {
                address: cell.address,
                value,
                formula,
                number_format,
            },
        });
    }
    Ok(findings)
}

fn inspect_pptx_slides(bytes: &[u8], start: u32, end: u32) -> AgentResult<Vec<InspectionFinding>> {
    if start == 0 || end < start {
        return Err(AgentError::InvalidRequest(
            "slide range must be 1-based and ordered".to_string(),
        ));
    }
    let slide_numbers = (start..=end)
        .map(|n| {
            usize::try_from(n)
                .map_err(|_| AgentError::InvalidRequest(format!("slide number {n} is too large")))
        })
        .collect::<AgentResult<std::collections::HashSet<_>>>()?;
    let doc = officemd_pptx::extract_ir_with_options(
        bytes,
        PptxExtractOptions {
            slide_numbers: Some(slide_numbers),
        },
    )
    .map_err(|e| AgentError::Extraction(e.to_string()))?;
    let mut findings = doc
        .slides
        .into_iter()
        .map(|slide| InspectionFinding {
            locator: ArtifactLocator::PptxSlide {
                slide_number: u32::try_from(slide.number).unwrap_or(u32::MAX),
            },
            kind: "slide".to_string(),
            payload: InspectionPayload::Slide {
                number: u32::try_from(slide.number).unwrap_or(u32::MAX),
                title: slide.title,
                has_notes: slide.notes.is_some_and(|notes| !notes.is_empty()),
                comment_count: slide.comments.len(),
            },
        })
        .collect::<Vec<_>>();
    findings.extend(
        officemd_pptx::inspect_shapes(bytes, start, end)
            .map_err(|error| AgentError::Extraction(error.to_string()))?
            .into_iter()
            .map(|shape| InspectionFinding {
                locator: ArtifactLocator::PptxShape {
                    slide_number: shape.slide_number,
                    shape_id: shape.shape_id,
                },
                kind: "shape".to_string(),
                payload: InspectionPayload::Shape {
                    slide_number: shape.slide_number,
                    shape_id: shape.shape_id,
                    text: shape.text,
                },
            }),
    );
    Ok(findings)
}

fn inspect_pdf_pages(
    bytes: &[u8],
    pages: &[u32],
    include: &PdfPageInclude,
) -> AgentResult<Vec<InspectionFinding>> {
    let doc = officemd_pdf::extract_ir_force_pages(bytes, false, Some(pages))
        .map_err(|e| AgentError::Extraction(e.to_string()))?;
    let Some(pdf) = doc.pdf else {
        return Ok(Vec::new());
    };
    Ok(pdf
        .pages
        .into_iter()
        .map(|page| InspectionFinding {
            locator: ArtifactLocator::PdfPage {
                page_number: u32::try_from(page.number).unwrap_or(u32::MAX),
            },
            kind: "pdf_page".to_string(),
            payload: InspectionPayload::PdfPage {
                number: u32::try_from(page.number).unwrap_or(u32::MAX),
                markdown: include.markdown.then_some(page.markdown),
            },
        })
        .collect())
}

fn extract_ir(bytes: &[u8], format: AgentDocumentFormat) -> AgentResult<OoxmlDocument> {
    match format {
        AgentDocumentFormat::Docx => {
            officemd_docx::extract_ir(bytes).map_err(|e| AgentError::Extraction(e.to_string()))
        }
        AgentDocumentFormat::Xlsx => officemd_xlsx::extract_tables_ir(bytes)
            .map_err(|e| AgentError::Extraction(e.to_string())),
        AgentDocumentFormat::Csv => officemd_csv::extract_tables_ir(bytes)
            .map_err(|e| AgentError::Extraction(e.to_string())),
        AgentDocumentFormat::Pptx => {
            officemd_pptx::extract_ir(bytes).map_err(|e| AgentError::Extraction(e.to_string()))
        }
        AgentDocumentFormat::Pdf => {
            officemd_pdf::extract_ir(bytes).map_err(|e| AgentError::Extraction(e.to_string()))
        }
    }
}

fn paragraph_text(paragraph: &Paragraph) -> String {
    paragraph
        .inlines
        .iter()
        .map(|inline| match inline {
            Inline::Text(text) => text.as_str(),
            Inline::Link(link) => link.display.as_str(),
        })
        .collect::<Vec<_>>()
        .join("")
}

fn cell_text(cell: &TableCell) -> String {
    cell.content
        .iter()
        .map(paragraph_text)
        .collect::<Vec<_>>()
        .join("\n")
}

fn parse_range(range: &str) -> AgentResult<(usize, usize, usize, usize)> {
    let mut parts = range.splitn(2, ':');
    let start = parse_cell_ref(parts.next().unwrap_or_default())?;
    let end = parts.next().map_or(Ok(start), parse_cell_ref)?;
    Ok((
        start.0.min(end.0),
        start.1.min(end.1),
        start.0.max(end.0),
        start.1.max(end.1),
    ))
}

fn parse_cell_ref(value: &str) -> AgentResult<(usize, usize)> {
    let trimmed = value.trim().trim_matches('$');
    let split = trimmed
        .find(|ch: char| ch.is_ascii_digit())
        .ok_or_else(|| AgentError::InvalidRequest(format!("invalid cell reference: {value}")))?;
    let (letters, digits) = trimmed.split_at(split);
    if letters.is_empty() || digits.is_empty() {
        return Err(AgentError::InvalidRequest(format!(
            "invalid cell reference: {value}"
        )));
    }
    let mut col = 0usize;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(AgentError::InvalidRequest(format!(
                "invalid cell reference: {value}"
            )));
        }
        col = col
            .saturating_mul(26)
            .saturating_add((ch.to_ascii_uppercase() as usize) - ('A' as usize) + 1);
    }
    let row = digits
        .parse::<usize>()
        .map_err(|_| AgentError::InvalidRequest(format!("invalid cell reference: {value}")))?;
    if row == 0 || col == 0 {
        return Err(AgentError::InvalidRequest(format!(
            "cell reference must be 1-based: {value}"
        )));
    }
    Ok((row - 1, col - 1))
}

fn column_name(mut col_zero_based: usize) -> String {
    let mut chars = Vec::new();
    loop {
        let rem = col_zero_based % 26;
        chars.push(char::from(b'A' + u8::try_from(rem).unwrap_or(0)));
        if col_zero_based < 26 {
            break;
        }
        col_zero_based = (col_zero_based / 26) - 1;
    }
    chars.iter().rev().collect()
}
