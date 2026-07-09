use clap::{CommandFactory, Parser, Subcommand, ValueEnum};
use officemd_agent::{
    AgentDocumentFormat, AgentService, ApplyPatchRequest, ArtifactPatchPlan, ArtifactRenderer,
    InspectQuery, InspectRequest, RenderRequest, RenderScale, VerificationCheckKind,
    VerificationStatus, VerifyRequest,
};
use officemd_core::ir::OoxmlDocument;
use officemd_core::opc::OpcPackage;
use officemd_pptx::PptxExtractOptions;
use officemd_xlsx::{SheetFilter, XlsxExtractOptions, inspect_sheet_summaries};
use serde::Serialize;
use similar::{ChangeTag, TextDiff};
use std::fmt::{Display, Formatter, Write as _};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

#[derive(Parser, Debug)]
#[command(
    name = "officemd",
    version,
    about = "Fast Office document extraction for LLMs and agents"
)]
struct Cli {
    #[command(subcommand)]
    command: Option<Command>,

    /// Show commands and options in a tree format. Depth 1 shows commands only,
    /// depth 2 includes arguments and options.
    #[arg(long, global = true, value_name = "DEPTH", default_missing_value = "2", num_args = 0..=1)]
    help_tree: Option<u8>,
}

/// Shared options used by convert, stream, and inspect.
#[derive(clap::Args, Debug, Clone)]
struct CommonOptions {
    /// Explicitly set the document format.
    #[arg(long, value_enum)]
    format: Option<FormatArg>,

    #[command(flatten)]
    output: CommonOutputOptions,

    /// Filter XLSX sheets by name or 1-based index (comma-separated).
    #[arg(long)]
    sheets: Option<String>,

    /// Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5").
    #[arg(long)]
    pages: Option<String>,

    /// Filter PPTX slides by number or range (e.g. "1-3,5").
    #[arg(long)]
    slides: Option<String>,

    #[command(flatten)]
    pdf: PdfCliOptions,

    #[command(flatten)]
    xlsx: XlsxCliOptions,

    #[command(flatten)]
    include: MarkdownIncludeCliOptions,

    #[command(flatten)]
    table: MarkdownTableCliOptions,

    /// Markdown style profile.
    #[arg(long, value_enum, default_value_t = MarkdownStyleArg::Compact)]
    markdown_style: MarkdownStyleArg,
}

#[derive(clap::Args, Debug, Clone)]
struct CommonOutputOptions {
    /// Output format: markdown (default) or json.
    #[arg(long, value_enum)]
    output_format: Option<OutputFormatArg>,

    /// Pretty-print JSON output.
    #[arg(long, default_value_t = false)]
    pretty: bool,
}

#[derive(clap::Args, Debug, Clone)]
struct PdfCliOptions {
    /// Force extraction even for scanned/image-based PDFs.
    #[arg(long, default_value_t = false)]
    force: bool,
}

#[derive(clap::Args, Debug, Clone)]
struct XlsxCliOptions {
    /// Use style-aware cell values for XLSX.
    #[arg(long, default_value_t = false)]
    style_aware: bool,

    /// Use streaming row parser for XLSX.
    #[arg(long, default_value_t = false)]
    streaming: bool,
}

#[derive(clap::Args, Debug, Clone)]
struct MarkdownIncludeCliOptions {
    /// Include document properties in markdown output.
    #[arg(
        long = "include-document-properties",
        alias = "document-properties",
        default_value_t = false
    )]
    document_properties: bool,

    /// Omit DOCX header/footer sections from markdown output.
    #[arg(long, default_value_t = false)]
    no_headers_footers: bool,

    /// Omit XLSX formula footnotes from markdown output.
    #[arg(long, default_value_t = false)]
    no_formulas: bool,

    /// Omit the leading `<!-- officemd: ... -->` frontmatter comment.
    #[arg(long, default_value_t = false)]
    no_frontmatter: bool,
}

#[derive(clap::Args, Debug, Clone)]
struct MarkdownTableCliOptions {
    /// Use synthetic Col1/Col2 headers instead of first data row.
    #[arg(long, default_value_t = false)]
    no_first_row_header: bool,
}

#[derive(Subcommand, Debug)]
enum Command {
    /// Probe artifact identity, capabilities, and operational risks.
    Probe {
        /// Input document path (.docx/.xlsx/.csv/.pptx/.pdf).
        input: PathBuf,

        #[command(flatten)]
        output: CommonOutputOptions,

        /// Explicitly set the document format.
        #[arg(long, value_enum)]
        format: Option<FormatArg>,
    },

    /// Apply a typed agent patch plan to a new output artifact.
    Apply {
        /// Input document path (.docx/.xlsx/.pptx).
        input: PathBuf,

        /// JSON file containing an ArtifactPatchPlan.
        #[arg(long)]
        patch: PathBuf,

        /// Output artifact path. Must not already exist.
        #[arg(short, long)]
        output: PathBuf,

        /// Expected SHA-256 fingerprint of the source artifact.
        #[arg(long)]
        expected_source_sha256: String,

        #[command(flatten)]
        output_options: CommonOutputOptions,
    },

    /// Render an artifact into visual evidence images when a backend is available.
    RenderArtifact {
        /// Input document path (.docx/.xlsx/.pptx/.pdf).
        input: PathBuf,

        /// Directory for rendered image evidence.
        #[arg(long)]
        output_dir: PathBuf,

        #[command(flatten)]
        output: CommonOutputOptions,
    },

    /// Verify semantic, structural, or visual artifact invariants.
    Verify {
        /// Input document path (.docx/.xlsx/.csv/.pptx/.pdf).
        input: PathBuf,

        /// Comma-separated checks: structure, formula-references, visual-render.
        #[arg(long)]
        checks: Option<String>,

        /// Directory for renderer-backed verification evidence.
        #[arg(long)]
        render_output_dir: Option<PathBuf>,

        #[command(flatten)]
        output: CommonOutputOptions,
    },

    /// Extract markdown, print to stdout.
    Markdown {
        /// Path to an input document.
        file: PathBuf,

        #[command(flatten)]
        common: CommonOptions,
    },

    /// Extract markdown, render to terminal with ANSI formatting.
    Render {
        /// Path to an input document.
        file: PathBuf,

        #[command(flatten)]
        common: CommonOptions,
    },

    /// Diff markdown output of two documents.
    Diff {
        /// Path to first input document.
        file_a: PathBuf,

        /// Path to second input document.
        file_b: PathBuf,

        #[command(flatten)]
        common: CommonOptions,
    },

    /// Convert a document file to an output file.
    Convert {
        /// Input document path (.docx/.xlsx/.csv/.pptx/.pdf).
        input: PathBuf,

        /// Output file path. Defaults to <input>.md or <input>.json.
        #[arg(short, long)]
        output: Option<PathBuf>,

        #[command(flatten)]
        common: CommonOptions,
    },

    /// Stream output to stdout from a file path or stdin.
    Stream {
        /// Input path or '-' for stdin.
        #[arg(default_value = "-")]
        input: PathBuf,

        #[command(flatten)]
        common: CommonOptions,
    },

    /// Inspect document metadata without full content rendering.
    Inspect {
        /// Input document path (.docx/.xlsx/.csv/.pptx/.pdf).
        input: PathBuf,

        #[command(flatten)]
        common: CommonOptions,

        /// JSON file containing an agent InspectQuery.
        #[arg(long)]
        agent_query: Option<PathBuf>,
    },

    /// Emit an agent-friendly parsing plan with follow-up commands.
    Plan {
        /// Input document path (.docx/.xlsx/.csv/.pptx/.pdf).
        input: PathBuf,

        #[command(flatten)]
        common: CommonOptions,
    },

    /// Create an Office document from markdown input.
    ///
    /// Reads officemd-flavored markdown and generates a .docx, .xlsx, or .pptx
    /// file. The output format is detected from the file extension.
    ///
    /// Examples:
    ///   officemd create report.docx < input.md
    ///   officemd create data.xlsx --input table.md
    ///   officemd create slides.pptx --input deck.md
    Create {
        /// Output file path (.docx, .xlsx, .pptx).
        output: PathBuf,

        /// Input markdown file, or '-' for stdin (default).
        #[arg(short, long, default_value = "-")]
        input: PathBuf,
    },
}

#[derive(Debug, Clone, Copy, ValueEnum, PartialEq, Eq)]
enum FormatArg {
    Docx,
    Xlsx,
    Csv,
    Pptx,
    Pdf,
}

#[derive(Debug, Clone, Copy, ValueEnum, PartialEq, Eq)]
enum OutputFormatArg {
    Markdown,
    Json,
}

#[derive(Debug, Clone, Copy, ValueEnum, PartialEq, Eq)]
enum MarkdownStyleArg {
    Compact,
    Human,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DocumentFormat {
    Docx,
    Xlsx,
    Csv,
    Pptx,
    Pdf,
}

impl From<FormatArg> for DocumentFormat {
    fn from(value: FormatArg) -> Self {
        match value {
            FormatArg::Docx => Self::Docx,
            FormatArg::Xlsx => Self::Xlsx,
            FormatArg::Csv => Self::Csv,
            FormatArg::Pptx => Self::Pptx,
            FormatArg::Pdf => Self::Pdf,
        }
    }
}

impl From<FormatArg> for AgentDocumentFormat {
    fn from(value: FormatArg) -> Self {
        match value {
            FormatArg::Docx => Self::Docx,
            FormatArg::Xlsx => Self::Xlsx,
            FormatArg::Csv => Self::Csv,
            FormatArg::Pptx => Self::Pptx,
            FormatArg::Pdf => Self::Pdf,
        }
    }
}

impl Display for DocumentFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Docx => write!(f, "docx"),
            Self::Xlsx => write!(f, "xlsx"),
            Self::Csv => write!(f, "csv"),
            Self::Pptx => write!(f, "pptx"),
            Self::Pdf => write!(f, "pdf"),
        }
    }
}

#[derive(Debug, Serialize)]
struct InspectInfo {
    format: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    sections: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    sheets: Option<Vec<SheetInfo>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    slides: Option<Vec<SlideInfo>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pdf: Option<PdfInfo>,
}

#[derive(Debug, Serialize)]
struct SheetInfo {
    name: String,
    rows: usize,
    cols: usize,
}

#[derive(Debug, Serialize)]
struct SlideInfo {
    number: usize,
    title: Option<String>,
    has_notes: bool,
    comment_count: usize,
}

#[derive(Debug, Serialize)]
struct PdfInfo {
    classification: String,
    confidence: f32,
    page_count: usize,
    pages_needing_ocr: Vec<usize>,
    has_encoding_issues: bool,
}

#[derive(Debug, Serialize)]
struct AgentParsePlan {
    format: String,
    input: String,
    inspect: InspectInfo,
    commands: AgentCommandSet,
    selectors: Vec<AgentSelector>,
    warnings: Vec<String>,
}

#[derive(Debug, Serialize)]
struct AgentCommandSet {
    inspect_json: Vec<String>,
    extract_markdown: Vec<String>,
    extract_json: Vec<String>,
}

#[derive(Debug, Serialize)]
struct AgentSelector {
    kind: String,
    flag: String,
    values: Vec<AgentSelectorValue>,
}

#[derive(Debug, Serialize)]
struct AgentSelectorValue {
    value: String,
    label: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    rows: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    cols: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    has_notes: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    comment_count: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    needs_ocr: Option<bool>,
    extract_markdown: Vec<String>,
    extract_json: Vec<String>,
}

impl CommonOptions {
    /// Effective output format: explicit choice, or Markdown by default.
    fn output_format(&self) -> OutputFormatArg {
        self.output
            .output_format
            .unwrap_or(OutputFormatArg::Markdown)
    }
}

impl clap_ai::AiDefaults for CommonOptions {
    fn apply_ai_defaults(&mut self) {
        // Only override when the user did not pass --output-format explicitly.
        if self.output.output_format.is_none() {
            self.output.output_format = Some(OutputFormatArg::Json);
        }
        if !self.output.pretty {
            self.output.pretty = true;
        }
    }
}

fn detect_format_from_path(path: &Path) -> Option<DocumentFormat> {
    path.extension().and_then(
        |ext| match ext.to_string_lossy().to_ascii_lowercase().as_str() {
            "docx" => Some(DocumentFormat::Docx),
            "xlsx" => Some(DocumentFormat::Xlsx),
            "csv" => Some(DocumentFormat::Csv),
            "pptx" => Some(DocumentFormat::Pptx),
            "pdf" => Some(DocumentFormat::Pdf),
            _ => None,
        },
    )
}

fn detect_format_from_bytes(content: &[u8]) -> Result<DocumentFormat, String> {
    if officemd_pdf::looks_like_pdf_header(content) {
        return Ok(DocumentFormat::Pdf);
    }

    let mut package = OpcPackage::from_bytes(content).map_err(|e| e.to_string())?;

    if package.has_part("word/document.xml") {
        return Ok(DocumentFormat::Docx);
    }
    if package.has_part("xl/workbook.xml") {
        return Ok(DocumentFormat::Xlsx);
    }
    if package.has_part("ppt/presentation.xml") {
        return Ok(DocumentFormat::Pptx);
    }

    Err("Could not detect format from file content (supported: docx, xlsx, csv, pptx, pdf; csv requires --format csv or .csv extension)".to_string())
}

fn resolve_format(
    content: &[u8],
    path_hint: Option<&Path>,
    explicit: Option<FormatArg>,
) -> Result<DocumentFormat, String> {
    if let Some(format) = explicit {
        return Ok(format.into());
    }

    if let Some(path) = path_hint
        && let Some(format) = detect_format_from_path(path)
    {
        return Ok(format);
    }

    detect_format_from_bytes(content)
}

fn extract_ir_document(
    content: &[u8],
    format: DocumentFormat,
    common: &CommonOptions,
) -> Result<OoxmlDocument, String> {
    if common.sheets.is_some() && format != DocumentFormat::Xlsx {
        if format == DocumentFormat::Pdf {
            eprintln!("Warning: --sheets is ignored for PDF files");
        } else {
            return Err("--sheets can only be used with XLSX files".to_string());
        }
    }
    if common.slides.is_some() && format != DocumentFormat::Pptx {
        if format == DocumentFormat::Pdf {
            eprintln!("Warning: --slides is ignored for PDF files");
        } else {
            return Err("--slides can only be used with PPTX files".to_string());
        }
    }

    let doc = match format {
        DocumentFormat::Docx => officemd_docx::extract_ir(content).map_err(|e| e.to_string())?,
        DocumentFormat::Xlsx => {
            let mut options = XlsxExtractOptions::default();
            options.text.style_aware_values = common.xlsx.style_aware;
            options.text.streaming_rows = common.xlsx.streaming;
            options.sheet_filter =
                build_sheet_filter(common.sheets.as_deref(), common.pages.as_deref())?;
            options.include.document_properties = common.include.document_properties
                || common.output_format() == OutputFormatArg::Json;
            options.trim.empty_edges = matches!(common.markdown_style, MarkdownStyleArg::Compact);

            officemd_xlsx::extract_tables_ir_with_options(content, &options)
                .map_err(|e| e.to_string())?
        }
        DocumentFormat::Csv => {
            let options = officemd_csv::table_ir::CsvExtractOptions {
                include_document_properties: common.include.document_properties
                    || common.output_format() == OutputFormatArg::Json,
                ..Default::default()
            };
            officemd_csv::extract_tables_ir_with_options(content, options)
                .map_err(|e| e.to_string())?
        }
        DocumentFormat::Pptx => {
            // --pages and --slides both select slides for PPTX
            let slides_spec = common.slides.as_deref().or(common.pages.as_deref());
            let slide_numbers = slides_spec
                .map(parse_number_ranges)
                .transpose()?
                .map(|values| values.into_iter().collect());
            let options = PptxExtractOptions { slide_numbers };
            officemd_pptx::extract_ir_with_options(content, options).map_err(|e| e.to_string())?
        }
        DocumentFormat::Pdf => {
            let page_numbers = common
                .pages
                .as_deref()
                .map(parse_positive_u32_ranges)
                .transpose()?;
            officemd_pdf::extract_ir_force_pages(content, common.pdf.force, page_numbers.as_deref())
                .map_err(|e| e.to_string())?
        }
    };

    // Warn about XLSX-specific flags used with non-XLSX formats
    if format != DocumentFormat::Xlsx {
        if common.xlsx.style_aware {
            eprintln!("Warning: --style-aware is only effective with XLSX files");
        }
        if common.xlsx.streaming {
            eprintln!("Warning: --streaming is only effective with XLSX files");
        }
    }

    Ok(doc)
}

fn parse_sheet_filter(spec: &str) -> SheetFilter {
    let mut filter = SheetFilter::default();
    for part in spec.split(',') {
        let value = part.trim();
        if value.is_empty() {
            continue;
        }
        if let Ok(idx) = value.parse::<usize>() {
            filter.indices_1_based.insert(idx);
        } else {
            filter.names.insert(value.to_string());
        }
    }
    filter
}

fn build_sheet_filter(
    sheets_spec: Option<&str>,
    pages_spec: Option<&str>,
) -> Result<Option<SheetFilter>, String> {
    let mut filter = sheets_spec.map(parse_sheet_filter).unwrap_or_default();

    if let Some(pages_spec) = pages_spec {
        for index in parse_number_ranges(pages_spec)? {
            filter.indices_1_based.insert(index);
        }
    }

    if filter.names.is_empty() && filter.indices_1_based.is_empty() {
        Ok(None)
    } else {
        Ok(Some(filter))
    }
}

fn parse_number_ranges(spec: &str) -> Result<Vec<usize>, String> {
    const MAX_EXPANDED_NUMBERS: usize = 100_000;
    let mut numbers = Vec::new();
    for part in spec.split(',') {
        let part = part.trim();
        if part.contains('-') {
            let bounds: Vec<&str> = part.splitn(2, '-').collect();
            let start: usize = bounds[0]
                .trim()
                .parse()
                .map_err(|_| format!("invalid range start: '{}'", bounds[0].trim()))?;
            let end: usize = bounds[1]
                .trim()
                .parse()
                .map_err(|_| format!("invalid range end: '{}'", bounds[1].trim()))?;
            if start > end {
                return Err(format!("invalid range: {start}-{end}"));
            }
            let width = end - start + 1;
            if width > MAX_EXPANDED_NUMBERS {
                return Err(format!(
                    "range {start}-{end} is too large (max {MAX_EXPANDED_NUMBERS} values)"
                ));
            }
            if numbers.len().saturating_add(width) > MAX_EXPANDED_NUMBERS {
                return Err(format!(
                    "too many slide values (max {MAX_EXPANDED_NUMBERS})"
                ));
            }
            for n in start..=end {
                numbers.push(n);
            }
        } else {
            let n: usize = part
                .parse()
                .map_err(|_| format!("invalid number: '{part}'"))?;
            if numbers.len() >= MAX_EXPANDED_NUMBERS {
                return Err(format!(
                    "too many slide values (max {MAX_EXPANDED_NUMBERS})"
                ));
            }
            numbers.push(n);
        }
    }
    Ok(numbers)
}

fn parse_positive_u32_ranges(spec: &str) -> Result<Vec<u32>, String> {
    parse_number_ranges(spec)?
        .into_iter()
        .map(|n| {
            if n == 0 {
                Err("page numbers must be >= 1".to_string())
            } else {
                u32::try_from(n).map_err(|_| format!("page number {n} is too large"))
            }
        })
        .collect()
}

fn render_output(doc: &OoxmlDocument, common: &CommonOptions) -> Result<String, String> {
    match common.output_format() {
        OutputFormatArg::Markdown => {
            let markdown_profile = match common.markdown_style {
                MarkdownStyleArg::Compact => officemd_markdown::MarkdownProfile::LlmCompact,
                MarkdownStyleArg::Human => officemd_markdown::MarkdownProfile::Human,
            };
            let options = officemd_markdown::RenderOptions {
                include: officemd_markdown::RenderIncludeOptions {
                    document_properties: common.include.document_properties,
                    headers_footers: !common.include.no_headers_footers,
                    formulas: !common.include.no_formulas,
                    frontmatter: !common.include.no_frontmatter,
                },
                table: officemd_markdown::RenderTableOptions {
                    first_row_as_header: !common.table.no_first_row_header,
                },
                markdown_profile,
            };
            Ok(officemd_markdown::render_document_with_options(
                doc, options,
            ))
        }
        OutputFormatArg::Json => {
            if common.output.pretty {
                serde_json::to_string_pretty(doc).map_err(|e| e.to_string())
            } else {
                serde_json::to_string(doc).map_err(|e| e.to_string())
            }
        }
    }
}

fn build_xlsx_inspect_info(
    content: &[u8],
    sheet_filter_spec: Option<&str>,
) -> Result<InspectInfo, String> {
    let sheet_filter = sheet_filter_spec.map(parse_sheet_filter);
    let sheet_summaries =
        inspect_sheet_summaries(content, sheet_filter.as_ref()).map_err(|e| e.to_string())?;
    Ok(InspectInfo {
        format: "xlsx".to_string(),
        sections: None,
        sheets: Some(
            sheet_summaries
                .into_iter()
                .map(|summary| SheetInfo {
                    name: summary.name,
                    rows: summary.rows,
                    cols: summary.cols,
                })
                .collect(),
        ),
        slides: None,
        pdf: None,
    })
}

fn build_pdf_inspect_info(content: &[u8]) -> Result<InspectInfo, String> {
    let diagnostics = officemd_pdf::inspect_pdf(content).map_err(|e| e.to_string())?;
    Ok(InspectInfo {
        format: "pdf".to_string(),
        sections: None,
        sheets: None,
        slides: None,
        pdf: Some(PdfInfo {
            classification: format!("{:?}", diagnostics.classification),
            confidence: diagnostics.confidence,
            page_count: diagnostics.page_count,
            pages_needing_ocr: diagnostics.pages_needing_ocr,
            has_encoding_issues: diagnostics.has_encoding_issues,
        }),
    })
}

fn build_inspect_info(doc: &OoxmlDocument, format: DocumentFormat) -> InspectInfo {
    match format {
        DocumentFormat::Docx => InspectInfo {
            format: "docx".to_string(),
            sections: Some(doc.sections.iter().map(|s| s.name.clone()).collect()),
            sheets: None,
            slides: None,
            pdf: None,
        },
        DocumentFormat::Xlsx => InspectInfo {
            format: "xlsx".to_string(),
            sections: None,
            sheets: Some(
                doc.sheets
                    .iter()
                    .map(|s| {
                        let (rows, cols) = s
                            .tables
                            .first()
                            .map_or((0, 0), |t| (t.rows.len(), t.headers.len()));
                        SheetInfo {
                            name: s.name.clone(),
                            rows,
                            cols,
                        }
                    })
                    .collect(),
            ),
            slides: None,
            pdf: None,
        },
        DocumentFormat::Csv => InspectInfo {
            format: "csv".to_string(),
            sections: None,
            sheets: Some(
                doc.sheets
                    .iter()
                    .map(|s| {
                        let (rows, cols) = s
                            .tables
                            .first()
                            .map_or((0, 0), |t| (t.rows.len(), t.headers.len()));
                        SheetInfo {
                            name: s.name.clone(),
                            rows,
                            cols,
                        }
                    })
                    .collect(),
            ),
            slides: None,
            pdf: None,
        },
        DocumentFormat::Pptx => InspectInfo {
            format: "pptx".to_string(),
            sections: None,
            sheets: None,
            slides: Some(
                doc.slides
                    .iter()
                    .map(|s| SlideInfo {
                        number: s.number,
                        title: s.title.clone(),
                        has_notes: s.notes.as_ref().is_some_and(|n| !n.is_empty()),
                        comment_count: s.comments.len(),
                    })
                    .collect(),
            ),
            pdf: None,
        },
        DocumentFormat::Pdf => {
            let info = doc.pdf.as_ref().map_or(
                PdfInfo {
                    classification: "Unknown".to_string(),
                    confidence: 0.0,
                    page_count: 0,
                    pages_needing_ocr: vec![],
                    has_encoding_issues: false,
                },
                |pdf| PdfInfo {
                    classification: format!("{:?}", pdf.diagnostics.classification),
                    confidence: pdf.diagnostics.confidence,
                    page_count: pdf.diagnostics.page_count,
                    pages_needing_ocr: pdf.diagnostics.pages_needing_ocr.clone(),
                    has_encoding_issues: pdf.diagnostics.has_encoding_issues,
                },
            );

            InspectInfo {
                format: "pdf".to_string(),
                sections: None,
                sheets: None,
                slides: None,
                pdf: Some(info),
            }
        }
    }
}

fn render_inspect_text(info: &InspectInfo) -> String {
    let mut out = String::new();
    let _ = writeln!(out, "Format: {}", info.format);

    if let Some(sections) = &info.sections {
        let _ = writeln!(out, "Sections: {}", sections.join(", "));
    }

    if let Some(sheets) = &info.sheets {
        let _ = writeln!(out, "Sheets ({}):", sheets.len());
        for (i, sheet) in sheets.iter().enumerate() {
            let _ = writeln!(
                out,
                "  {}. {} ({} rows x {} cols)",
                i + 1,
                sheet.name,
                sheet.rows,
                sheet.cols
            );
        }
    }

    if let Some(slides) = &info.slides {
        let _ = writeln!(out, "Slides ({}):", slides.len());
        for slide in slides {
            let title = slide.title.as_deref().unwrap_or("(untitled)");
            let mut annotations = Vec::new();
            if slide.has_notes {
                annotations.push("notes".to_string());
            }
            if slide.comment_count > 0 {
                annotations.push(format!("{} comments", slide.comment_count));
            }
            if annotations.is_empty() {
                let _ = writeln!(out, "  {}. {}", slide.number, title);
            } else {
                let _ = writeln!(
                    out,
                    "  {}. {} [{}]",
                    slide.number,
                    title,
                    annotations.join("] [")
                );
            }
        }
    }

    if let Some(pdf) = &info.pdf {
        out.push_str("PDF Diagnostics:\n");
        let _ = writeln!(out, "  Classification: {}", pdf.classification);
        let _ = writeln!(out, "  Confidence: {:.4}", pdf.confidence);
        let _ = writeln!(out, "  Page count: {}", pdf.page_count);
        if pdf.pages_needing_ocr.is_empty() {
            out.push_str("  Pages needing OCR: none\n");
        } else {
            let _ = writeln!(
                out,
                "  Pages needing OCR: {}",
                pdf.pages_needing_ocr
                    .iter()
                    .map(usize::to_string)
                    .collect::<Vec<_>>()
                    .join(", ")
            );
        }
        let _ = writeln!(out, "  Has encoding issues: {}", pdf.has_encoding_issues);
    }

    out
}

fn build_agent_plan(
    input: &Path,
    resolved: DocumentFormat,
    common: &CommonOptions,
    inspect: InspectInfo,
) -> AgentParsePlan {
    let input_label = input.display().to_string();
    let format_label = resolved.to_string();
    let commands = AgentCommandSet {
        inspect_json: build_extract_command(
            "inspect",
            &input_label,
            resolved,
            common,
            OutputFormatArg::Json,
            None,
        ),
        extract_markdown: build_extract_command(
            "stream",
            &input_label,
            resolved,
            common,
            OutputFormatArg::Markdown,
            None,
        ),
        extract_json: build_extract_command(
            "stream",
            &input_label,
            resolved,
            common,
            OutputFormatArg::Json,
            None,
        ),
    };

    let selectors = build_agent_selectors(&input_label, resolved, common, &inspect);
    let warnings = build_agent_warnings(&inspect);

    AgentParsePlan {
        format: format_label,
        input: input_label,
        inspect,
        commands,
        selectors,
        warnings,
    }
}

fn build_agent_selectors(
    input_label: &str,
    resolved: DocumentFormat,
    common: &CommonOptions,
    inspect: &InspectInfo,
) -> Vec<AgentSelector> {
    match resolved {
        DocumentFormat::Xlsx => inspect.sheets.as_ref().map_or_else(Vec::new, |sheets| {
            vec![AgentSelector {
                kind: "sheet".to_string(),
                flag: "--sheets".to_string(),
                values: sheets
                    .iter()
                    .enumerate()
                    .map(|(index, sheet)| {
                        let value = (index + 1).to_string();
                        AgentSelectorValue {
                            value: value.clone(),
                            label: sheet.name.clone(),
                            rows: Some(sheet.rows),
                            cols: Some(sheet.cols),
                            title: None,
                            has_notes: None,
                            comment_count: None,
                            needs_ocr: None,
                            extract_markdown: build_extract_command(
                                "stream",
                                input_label,
                                resolved,
                                common,
                                OutputFormatArg::Markdown,
                                Some(("--sheets", &value)),
                            ),
                            extract_json: build_extract_command(
                                "stream",
                                input_label,
                                resolved,
                                common,
                                OutputFormatArg::Json,
                                Some(("--sheets", &value)),
                            ),
                        }
                    })
                    .collect(),
            }]
        }),
        DocumentFormat::Pptx => inspect.slides.as_ref().map_or_else(Vec::new, |slides| {
            vec![AgentSelector {
                kind: "slide".to_string(),
                flag: "--slides".to_string(),
                values: slides
                    .iter()
                    .map(|slide| {
                        let value = slide.number.to_string();
                        AgentSelectorValue {
                            value: value.clone(),
                            label: slide
                                .title
                                .clone()
                                .unwrap_or_else(|| format!("Slide {}", slide.number)),
                            rows: None,
                            cols: None,
                            title: slide.title.clone(),
                            has_notes: Some(slide.has_notes),
                            comment_count: Some(slide.comment_count),
                            needs_ocr: None,
                            extract_markdown: build_extract_command(
                                "stream",
                                input_label,
                                resolved,
                                common,
                                OutputFormatArg::Markdown,
                                Some(("--slides", &value)),
                            ),
                            extract_json: build_extract_command(
                                "stream",
                                input_label,
                                resolved,
                                common,
                                OutputFormatArg::Json,
                                Some(("--slides", &value)),
                            ),
                        }
                    })
                    .collect(),
            }]
        }),
        DocumentFormat::Pdf => inspect.pdf.as_ref().map_or_else(Vec::new, |pdf| {
            if pdf.page_count == 0 {
                return Vec::new();
            }

            let mut values = vec![pdf_page_selector_value(
                input_label,
                resolved,
                common,
                "1",
                "first page",
                pdf.pages_needing_ocr.contains(&1),
            )];

            if pdf.page_count > 1 {
                let all_pages = format!("1-{}", pdf.page_count);
                values.push(pdf_page_selector_value(
                    input_label,
                    resolved,
                    common,
                    &all_pages,
                    "all pages",
                    false,
                ));
            }

            if !pdf.pages_needing_ocr.is_empty() {
                let ocr_pages = pdf
                    .pages_needing_ocr
                    .iter()
                    .map(usize::to_string)
                    .collect::<Vec<_>>()
                    .join(",");
                values.push(pdf_page_selector_value(
                    input_label,
                    resolved,
                    common,
                    &ocr_pages,
                    "pages needing OCR",
                    true,
                ));
            }

            vec![AgentSelector {
                kind: "page".to_string(),
                flag: "--pages".to_string(),
                values,
            }]
        }),
        DocumentFormat::Docx | DocumentFormat::Csv => Vec::new(),
    }
}

fn pdf_page_selector_value(
    input_label: &str,
    resolved: DocumentFormat,
    common: &CommonOptions,
    value: &str,
    label: &str,
    needs_ocr: bool,
) -> AgentSelectorValue {
    AgentSelectorValue {
        value: value.to_string(),
        label: label.to_string(),
        rows: None,
        cols: None,
        title: None,
        has_notes: None,
        comment_count: None,
        needs_ocr: Some(needs_ocr),
        extract_markdown: build_extract_command(
            "stream",
            input_label,
            resolved,
            common,
            OutputFormatArg::Markdown,
            Some(("--pages", value)),
        ),
        extract_json: build_extract_command(
            "stream",
            input_label,
            resolved,
            common,
            OutputFormatArg::Json,
            Some(("--pages", value)),
        ),
    }
}

fn build_agent_warnings(inspect: &InspectInfo) -> Vec<String> {
    let mut warnings = Vec::new();
    if let Some(pdf) = &inspect.pdf {
        if !pdf.pages_needing_ocr.is_empty() {
            warnings.push(format!(
                "PDF has pages that likely need OCR: {}",
                pdf.pages_needing_ocr
                    .iter()
                    .map(usize::to_string)
                    .collect::<Vec<_>>()
                    .join(", ")
            ));
        }
        if pdf.has_encoding_issues {
            warnings.push("PDF has font encoding issues; extracted text may be unreliable".into());
        }
    }
    warnings
}

fn build_extract_command(
    command: &str,
    input_label: &str,
    resolved: DocumentFormat,
    common: &CommonOptions,
    output_format: OutputFormatArg,
    selector: Option<(&str, &str)>,
) -> Vec<String> {
    let mut args = vec![
        "officemd".to_string(),
        command.to_string(),
        input_label.to_string(),
        "--format".to_string(),
        resolved.to_string(),
        "--output-format".to_string(),
        match output_format {
            OutputFormatArg::Markdown => "markdown".to_string(),
            OutputFormatArg::Json => "json".to_string(),
        },
    ];

    if output_format == OutputFormatArg::Json || common.output.pretty {
        args.push("--pretty".to_string());
    }

    push_agent_common_flags(&mut args, resolved, common, selector);
    args
}

fn push_agent_common_flags(
    args: &mut Vec<String>,
    resolved: DocumentFormat,
    common: &CommonOptions,
    selector: Option<(&str, &str)>,
) {
    if let Some((flag, value)) = selector {
        args.push(flag.to_string());
        args.push(value.to_string());
    } else {
        match resolved {
            DocumentFormat::Xlsx => {
                if let Some(sheets) = &common.sheets {
                    args.push("--sheets".to_string());
                    args.push(sheets.clone());
                }
                if let Some(pages) = &common.pages {
                    args.push("--pages".to_string());
                    args.push(pages.clone());
                }
            }
            DocumentFormat::Pptx => {
                if let Some(slides) = &common.slides {
                    args.push("--slides".to_string());
                    args.push(slides.clone());
                } else if let Some(pages) = &common.pages {
                    args.push("--pages".to_string());
                    args.push(pages.clone());
                }
            }
            DocumentFormat::Pdf => {
                if let Some(pages) = &common.pages {
                    args.push("--pages".to_string());
                    args.push(pages.clone());
                }
            }
            DocumentFormat::Docx | DocumentFormat::Csv => {}
        }
    }

    if common.pdf.force && resolved == DocumentFormat::Pdf {
        args.push("--force".to_string());
    }
    if common.xlsx.style_aware && resolved == DocumentFormat::Xlsx {
        args.push("--style-aware".to_string());
    }
    if common.xlsx.streaming && resolved == DocumentFormat::Xlsx {
        args.push("--streaming".to_string());
    }
    if common.include.document_properties {
        args.push("--include-document-properties".to_string());
    }
    if common.include.no_headers_footers {
        args.push("--no-headers-footers".to_string());
    }
    if common.include.no_formulas {
        args.push("--no-formulas".to_string());
    }
    if common.include.no_frontmatter {
        args.push("--no-frontmatter".to_string());
    }
    if common.table.no_first_row_header {
        args.push("--no-first-row-header".to_string());
    }
    if common.markdown_style != MarkdownStyleArg::Compact {
        args.push("--markdown-style".to_string());
        args.push(
            match common.markdown_style {
                MarkdownStyleArg::Compact => "compact",
                MarkdownStyleArg::Human => "human",
            }
            .to_string(),
        );
    }
}

fn read_all_from_stdin() -> Result<Vec<u8>, String> {
    let mut bytes = Vec::new();
    std::io::stdin()
        .read_to_end(&mut bytes)
        .map_err(|e| format!("failed to read stdin: {e}"))?;
    Ok(bytes)
}

fn default_output_path(input: &Path, output_format: OutputFormatArg) -> PathBuf {
    let mut out = input.to_path_buf();
    match output_format {
        OutputFormatArg::Markdown => out.set_extension("md"),
        OutputFormatArg::Json => out.set_extension("json"),
    };
    out
}

fn validate_xlsx_markdown_fast_path(common: &CommonOptions) -> Result<(), String> {
    if common.slides.is_some() {
        return Err("--slides can only be used with PPTX files".to_string());
    }

    Ok(())
}

fn extract_markdown_from_file(path: &Path, common: &CommonOptions) -> Result<String, String> {
    let bytes =
        std::fs::read(path).map_err(|e| format!("failed to read '{}': {e}", path.display()))?;
    let resolved = resolve_format(&bytes, Some(path), common.format)?;

    // Force markdown output regardless of --output-format
    let markdown_profile = match common.markdown_style {
        MarkdownStyleArg::Compact => officemd_markdown::MarkdownProfile::LlmCompact,
        MarkdownStyleArg::Human => officemd_markdown::MarkdownProfile::Human,
    };
    let options = officemd_markdown::RenderOptions {
        include: officemd_markdown::RenderIncludeOptions {
            document_properties: common.include.document_properties,
            headers_footers: !common.include.no_headers_footers,
            formulas: !common.include.no_formulas,
            frontmatter: !common.include.no_frontmatter,
        },
        table: officemd_markdown::RenderTableOptions {
            first_row_as_header: !common.table.no_first_row_header,
        },
        markdown_profile,
    };
    let md = if resolved == DocumentFormat::Xlsx {
        validate_xlsx_markdown_fast_path(common)?;

        let mut extract_options = XlsxExtractOptions::default();
        extract_options.text.style_aware_values = common.xlsx.style_aware;
        extract_options.text.streaming_rows = common.xlsx.streaming;
        extract_options.sheet_filter =
            build_sheet_filter(common.sheets.as_deref(), common.pages.as_deref())?;
        extract_options.include.document_properties = common.include.document_properties;
        extract_options.trim.empty_edges =
            matches!(common.markdown_style, MarkdownStyleArg::Compact);

        officemd_xlsx::markdown_from_bytes_with_extract_options(&bytes, options, &extract_options)
            .map_err(|e| e.to_string())?
    } else {
        let doc = extract_ir_document(&bytes, resolved, common)?;
        officemd_markdown::render_document_with_options(&doc, options)
    };

    if common.pages.is_some() && resolved == DocumentFormat::Csv {
        eprintln!("Hint: use --sheets for sheet selection with csv files");
    }

    // Warn about scanned PDFs
    if resolved == DocumentFormat::Pdf && md.trim().len() < 50 {
        warn_scanned_pdf(&bytes, common.pdf.force);
    }

    Ok(md)
}

fn warn_scanned_pdf(content: &[u8], force: bool) {
    let Ok(diagnostics) = officemd_pdf::inspect_pdf(content) else {
        return;
    };

    let class = format!("{:?}", diagnostics.classification);
    if class != "Scanned" && class != "ImageBased" {
        return;
    }

    if force {
        eprintln!(
            "Info: PDF classified as {} (confidence: {:.0}%, {} page(s)). \
             Forced extraction attempted - output may be empty or incomplete.",
            class,
            diagnostics.confidence * 100.0,
            diagnostics.page_count,
        );
    } else {
        let ocr_summary = if diagnostics.pages_needing_ocr.is_empty() {
            format!("{} page(s)", diagnostics.page_count)
        } else {
            format!(
                "pages needing OCR: {}",
                diagnostics
                    .pages_needing_ocr
                    .iter()
                    .map(usize::to_string)
                    .collect::<Vec<_>>()
                    .join(", ")
            )
        };
        eprintln!(
            "Warning: PDF classified as {} (confidence: {:.0}%, {}). \
             No text could be extracted - this document likely needs OCR.\n\
             Hint: use --force to attempt extraction anyway.",
            class,
            diagnostics.confidence * 100.0,
            ocr_summary,
        );
    }
}

const DIFF_RED: &str = "\x1b[31m";
const DIFF_GREEN: &str = "\x1b[32m";
const DIFF_CYAN: &str = "\x1b[36m";
const DIFF_RESET: &str = "\x1b[0m";

fn render_colored_diff(text_a: &str, text_b: &str, label_a: &str, label_b: &str) -> String {
    let diff = TextDiff::from_lines(text_a, text_b);
    let mut out = String::new();

    for hunk in diff.unified_diff().header(label_a, label_b).iter_hunks() {
        for change in hunk.iter_changes() {
            let (color, sign) = match change.tag() {
                ChangeTag::Delete => (DIFF_RED, "-"),
                ChangeTag::Insert => (DIFF_GREEN, "+"),
                ChangeTag::Equal => ("", " "),
            };
            if color.is_empty() {
                let _ = write!(out, "{sign}{change}");
            } else {
                let _ = write!(out, "{color}{sign}{change}{DIFF_RESET}");
            }
            if change.missing_newline() {
                out.push('\n');
            }
        }
        let _ = writeln!(out, "{DIFF_CYAN}---{DIFF_RESET}");
    }

    if out.is_empty() {
        out.push_str("(no differences)\n");
    }

    out
}

fn print_help_tree(depth: u8) {
    let cmd = Cli::command();
    let opts = clap_ai::HelpTreeOptions {
        depth,
        root_suffix: if clap_ai::is_ai_mode() {
            Some(" [AI mode]".into())
        } else {
            None
        },
        footer_lines: if clap_ai::is_ai_mode() {
            vec![
                String::new(),
                "AI mode active (AI=True): defaults to --output-format json --pretty".into(),
            ]
        } else {
            vec![]
        },
    };
    clap_ai::print_help_tree(&cmd, &opts);
}

fn write_stdout(output: &str) -> Result<(), String> {
    let mut stdout = std::io::stdout().lock();
    stdout
        .write_all(output.as_bytes())
        .map_err(|e| format!("failed to write stdout: {e}"))
}

fn run_markdown_command(file: &Path, common: &CommonOptions) -> Result<(), String> {
    write_stdout(&extract_markdown_from_file(file, common)?)
}

fn run_diff_command(file_a: &Path, file_b: &Path, common: &CommonOptions) -> Result<(), String> {
    let md_a = extract_markdown_from_file(file_a, common)?;
    let md_b = extract_markdown_from_file(file_b, common)?;
    let label_a = file_a.display().to_string();
    let label_b = file_b.display().to_string();
    write_stdout(&render_colored_diff(&md_a, &md_b, &label_a, &label_b))
}

fn run_convert_command(
    input: &Path,
    output: Option<PathBuf>,
    mut common: CommonOptions,
) -> Result<(), String> {
    clap_ai::maybe_apply_ai_defaults(&mut common);
    let bytes = std::fs::read(input)
        .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?;
    let resolved = resolve_format(&bytes, Some(input), common.format)?;
    let doc = extract_ir_document(&bytes, resolved, &common)?;
    let rendered = render_output(&doc, &common)?;
    let output_path = output.unwrap_or_else(|| default_output_path(input, common.output_format()));
    std::fs::write(&output_path, &rendered)
        .map_err(|e| format!("failed to write output '{}': {e}", output_path.display()))?;
    let format_label = match common.output_format() {
        OutputFormatArg::Markdown => "markdown",
        OutputFormatArg::Json => "JSON",
    };
    eprintln!(
        "Wrote {} for {} document to {}",
        format_label,
        resolved,
        output_path.display()
    );
    Ok(())
}

fn run_stream_command(input: &Path, mut common: CommonOptions) -> Result<(), String> {
    clap_ai::maybe_apply_ai_defaults(&mut common);
    let use_stdin = input == Path::new("-");
    let bytes = if use_stdin {
        read_all_from_stdin()?
    } else {
        std::fs::read(input)
            .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?
    };

    let path_hint = if use_stdin { None } else { Some(input) };
    let resolved = resolve_format(&bytes, path_hint, common.format)?;
    let doc = extract_ir_document(&bytes, resolved, &common)?;
    drop(bytes);
    write_stdout(&render_output(&doc, &common)?)
}

fn run_inspect_command(input: &Path, mut common: CommonOptions) -> Result<(), String> {
    clap_ai::maybe_apply_ai_defaults(&mut common);
    let bytes = std::fs::read(input)
        .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?;
    let resolved = resolve_format(&bytes, Some(input), common.format)?;
    let info = inspect_input(&bytes, resolved, &common)?;
    let output = match common.output_format() {
        OutputFormatArg::Json if common.output.pretty => {
            serde_json::to_string_pretty(&info).map_err(|e| e.to_string())?
        }
        OutputFormatArg::Json => serde_json::to_string(&info).map_err(|e| e.to_string())?,
        OutputFormatArg::Markdown => render_inspect_text(&info),
    };
    write_stdout(&output)
}

fn run_probe_command(
    input: &Path,
    output: CommonOutputOptions,
    format: Option<FormatArg>,
) -> Result<(), String> {
    let service = AgentService::new();
    let request = InspectRequest {
        input: input.to_path_buf(),
        format: format.map(AgentDocumentFormat::from),
        query: InspectQuery::DocumentSummary,
    };
    let mut report = service.inspect(&request).map_err(|e| e.to_string())?;
    report.capability.render_capability = officemd_renderer::discover();
    let output_text = match output.output_format.unwrap_or(OutputFormatArg::Json) {
        OutputFormatArg::Json if output.pretty => {
            serde_json::to_string_pretty(&report).map_err(|e| e.to_string())?
        }
        OutputFormatArg::Json => serde_json::to_string(&report).map_err(|e| e.to_string())?,
        OutputFormatArg::Markdown => render_agent_probe_text(&report),
    };
    write_stdout(&output_text)
}

fn run_agent_inspect_command(
    input: &Path,
    common: CommonOptions,
    agent_query: &Path,
) -> Result<(), String> {
    let query_bytes = std::fs::read(agent_query).map_err(|e| {
        format!(
            "failed to read agent query '{}': {e}",
            agent_query.display()
        )
    })?;
    let query: InspectQuery = serde_json::from_slice(&query_bytes)
        .map_err(|e| format!("failed to parse agent query JSON: {e}"))?;
    let service = AgentService::new();
    let request = InspectRequest {
        input: input.to_path_buf(),
        format: common.format.map(AgentDocumentFormat::from),
        query,
    };
    let report = service.inspect(&request).map_err(|e| e.to_string())?;
    let output = match common.output_format() {
        OutputFormatArg::Json if common.output.pretty => {
            serde_json::to_string_pretty(&report).map_err(|e| e.to_string())?
        }
        OutputFormatArg::Json => serde_json::to_string(&report).map_err(|e| e.to_string())?,
        OutputFormatArg::Markdown => render_agent_probe_text(&report),
    };
    write_stdout(&output)
}

fn run_apply_command(
    input: &Path,
    patch_path: &Path,
    output: &Path,
    expected_source_sha256: String,
    output_options: CommonOutputOptions,
) -> Result<(), String> {
    let patch_bytes = std::fs::read(patch_path)
        .map_err(|e| format!("failed to read patch '{}': {e}", patch_path.display()))?;
    let patch: ArtifactPatchPlan = serde_json::from_slice(&patch_bytes)
        .map_err(|e| format!("failed to parse patch plan JSON: {e}"))?;
    let request = ApplyPatchRequest {
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        expected_source_sha256,
        patch,
    };
    let report = AgentService::new()
        .apply_patch(&request)
        .map_err(|e| e.to_string())?;
    let output_text = match output_options
        .output_format
        .unwrap_or(OutputFormatArg::Json)
    {
        OutputFormatArg::Json if output_options.pretty => {
            serde_json::to_string_pretty(&report).map_err(|e| e.to_string())?
        }
        OutputFormatArg::Json => serde_json::to_string(&report).map_err(|e| e.to_string())?,
        OutputFormatArg::Markdown => render_apply_text(&report),
    };
    write_stdout(&output_text)
}

fn run_render_artifact_command(
    input: &Path,
    output_dir: &Path,
    output: CommonOutputOptions,
) -> Result<(), String> {
    let request = RenderRequest {
        input: input.to_path_buf(),
        output_dir: output_dir.to_path_buf(),
        pages_or_slides: None,
        scale: RenderScale::Screen,
    };
    let report = officemd_renderer::SystemRenderer::discover()
        .render(&request)
        .map_err(|e| e.to_string())?;
    let output_text = match output.output_format.unwrap_or(OutputFormatArg::Json) {
        OutputFormatArg::Json if output.pretty => {
            serde_json::to_string_pretty(&report).map_err(|e| e.to_string())?
        }
        OutputFormatArg::Json => serde_json::to_string(&report).map_err(|e| e.to_string())?,
        OutputFormatArg::Markdown => render_render_report_text(&report),
    };
    write_stdout(&output_text)
}

fn run_verify_command(
    input: &Path,
    checks: Option<String>,
    render_output_dir: Option<PathBuf>,
    output: CommonOutputOptions,
) -> Result<(), String> {
    let request = VerifyRequest {
        input: input.to_path_buf(),
        checks: parse_verification_checks(checks.as_deref())?,
        render_output_dir,
    };
    let mut report = AgentService::new()
        .verify(&request)
        .map_err(|e| e.to_string())?;
    if request
        .checks
        .contains(&VerificationCheckKind::VisualRender)
        && let Some(render_output_dir) = &request.render_output_dir
    {
        let render_request = RenderRequest {
            input: input.to_path_buf(),
            output_dir: render_output_dir.clone(),
            pages_or_slides: None,
            scale: RenderScale::Screen,
        };
        match officemd_renderer::SystemRenderer::discover().render(&render_request) {
            Ok(rendered) => {
                for check in &mut report.checks {
                    if check.kind == VerificationCheckKind::VisualRender {
                        check.status = VerificationStatus::Passed;
                        check.message = format!(
                            "rendered {} visual evidence image(s)",
                            rendered.images.len()
                        );
                    }
                }
                report.rendered_evidence = Some(rendered);
                report.status = aggregate_cli_verify_status(&report.checks);
            }
            Err(err) => {
                for check in &mut report.checks {
                    if check.kind == VerificationCheckKind::VisualRender {
                        check.status = VerificationStatus::NotRunnable;
                        check.message = err.to_string();
                    }
                }
                report.status = aggregate_cli_verify_status(&report.checks);
            }
        }
    }
    let output_text = match output.output_format.unwrap_or(OutputFormatArg::Json) {
        OutputFormatArg::Json if output.pretty => {
            serde_json::to_string_pretty(&report).map_err(|e| e.to_string())?
        }
        OutputFormatArg::Json => serde_json::to_string(&report).map_err(|e| e.to_string())?,
        OutputFormatArg::Markdown => render_verify_text(&report),
    };
    write_stdout(&output_text)
}

fn aggregate_cli_verify_status(checks: &[officemd_agent::CheckReport]) -> VerificationStatus {
    if checks
        .iter()
        .any(|check| check.status == VerificationStatus::Failed)
    {
        VerificationStatus::Failed
    } else if checks
        .iter()
        .any(|check| check.status == VerificationStatus::NotRunnable)
    {
        VerificationStatus::PassedWithWarnings
    } else {
        VerificationStatus::Passed
    }
}

fn parse_verification_checks(spec: Option<&str>) -> Result<Vec<VerificationCheckKind>, String> {
    let Some(spec) = spec else {
        return Ok(vec![VerificationCheckKind::Structure]);
    };
    spec.split(',')
        .map(|raw| match raw.trim() {
            "structure" => Ok(VerificationCheckKind::Structure),
            "formula-references" | "formula_references" => {
                Ok(VerificationCheckKind::FormulaReferences)
            }
            "visual-render" | "visual_render" => Ok(VerificationCheckKind::VisualRender),
            value => Err(format!(
                "unknown verification check '{value}' (expected structure, formula-references, visual-render)"
            )),
        })
        .collect()
}

fn render_apply_text(report: &officemd_agent::ApplyPatchReport) -> String {
    let mut out = String::new();
    let _ = writeln!(out, "Status: {:?}", report.status);
    let _ = writeln!(out, "Source: {}", report.source.path.display());
    if let Some(output) = &report.output {
        let _ = writeln!(out, "Output: {}", output.path.display());
    }
    let _ = writeln!(out, "Operations: {}", report.operations.len());
    out
}

fn render_render_report_text(report: &officemd_agent::RenderReport) -> String {
    let mut out = String::new();
    let _ = writeln!(out, "Backend: {:?}", report.backend);
    let _ = writeln!(out, "Images: {}", report.images.len());
    for diagnostic in &report.diagnostics {
        let _ = writeln!(out, "{}: {}", diagnostic.code, diagnostic.message);
    }
    out
}

fn render_verify_text(report: &officemd_agent::VerificationReport) -> String {
    let mut out = String::new();
    let _ = writeln!(out, "Status: {:?}", report.status);
    for check in &report.checks {
        let _ = writeln!(
            out,
            "{:?}: {:?} - {}",
            check.kind, check.status, check.message
        );
    }
    out
}

fn render_agent_probe_text(report: &officemd_agent::InspectionReport) -> String {
    let mut out = String::new();
    let _ = writeln!(out, "Format: {}", report.artifact.format);
    let _ = writeln!(out, "SHA-256: {}", report.artifact.fingerprint.sha256);
    let _ = writeln!(out, "Bytes: {}", report.artifact.fingerprint.byte_length);
    let _ = writeln!(out, "Readable: {}", report.capability.readable);
    if report.capability.mutable_operations.is_empty() {
        out.push_str("Mutable operations: none\n");
    } else {
        let operations = report
            .capability
            .mutable_operations
            .iter()
            .map(|op| format!("{op:?}"))
            .collect::<Vec<_>>()
            .join(", ");
        let _ = writeln!(out, "Mutable operations: {operations}");
    }
    out
}

fn run_plan_command(input: &Path, common: CommonOptions) -> Result<(), String> {
    let bytes = std::fs::read(input)
        .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?;
    let resolved = resolve_format(&bytes, Some(input), common.format)?;
    let inspect = inspect_input(&bytes, resolved, &common)?;
    let plan = build_agent_plan(input, resolved, &common, inspect);
    let output = serde_json::to_string_pretty(&plan).map_err(|e| e.to_string())?;
    write_stdout(&output)
}

fn inspect_input(
    bytes: &[u8],
    resolved: DocumentFormat,
    common: &CommonOptions,
) -> Result<InspectInfo, String> {
    if resolved == DocumentFormat::Xlsx {
        if common.slides.is_some() {
            return Err("--slides can only be used with PPTX files".to_string());
        }
        build_xlsx_inspect_info(bytes, common.sheets.as_deref())
    } else if resolved == DocumentFormat::Pdf {
        if common.sheets.is_some() {
            eprintln!("Warning: --sheets is ignored for PDF files");
        }
        if common.slides.is_some() {
            eprintln!("Warning: --slides is ignored for PDF files");
        }
        build_pdf_inspect_info(bytes)
    } else {
        let doc = extract_ir_document(bytes, resolved, common)?;
        Ok(build_inspect_info(&doc, resolved))
    }
}

fn run_create_command(output: &Path, input: &Path) -> Result<(), String> {
    let use_stdin = input == Path::new("-");
    let markdown = if use_stdin {
        let bytes = read_all_from_stdin()?;
        String::from_utf8(bytes).map_err(|e| format!("invalid UTF-8 in stdin: {e}"))?
    } else {
        std::fs::read_to_string(input)
            .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?
    };

    let doc = officemd_markdown::parse_document(&markdown)
        .map_err(|e| format!("failed to parse markdown: {e}"))?;
    let bytes = create_document_bytes(output, &doc)?;
    std::fs::write(output, &bytes)
        .map_err(|e| format!("failed to write '{}': {e}", output.display()))?;
    eprintln!("Created {} ({} bytes)", output.display(), bytes.len());
    Ok(())
}

fn create_document_bytes(output: &Path, doc: &OoxmlDocument) -> Result<Vec<u8>, String> {
    let ext = output
        .extension()
        .and_then(|e| e.to_str())
        .map(str::to_ascii_lowercase);

    match ext.as_deref() {
        Some("docx") => {
            officemd_docx::generate_docx(doc).map_err(|e| format!("DOCX generation failed: {e}"))
        }
        Some("xlsx") => {
            officemd_xlsx::generate_xlsx(doc).map_err(|e| format!("XLSX generation failed: {e}"))
        }
        Some("pptx") => {
            officemd_pptx::generate_pptx(doc).map_err(|e| format!("PPTX generation failed: {e}"))
        }
        _ => Err(format!(
            "unsupported output format '{}'. Use .docx, .xlsx, or .pptx",
            output.display()
        )),
    }
}

fn run() -> Result<(), String> {
    let cli = Cli::parse();

    // Handle --help-tree before anything else
    if let Some(depth) = cli.help_tree {
        print_help_tree(depth);
        return Ok(());
    }

    let command = cli.command.ok_or_else(|| {
        // No subcommand and no --help-tree: show usage hint
        "no subcommand provided. Use --help for usage or --help-tree for a command overview."
            .to_string()
    })?;

    match command {
        Command::Probe {
            input,
            output,
            format,
        } => run_probe_command(&input, output, format)?,
        Command::Apply {
            input,
            patch,
            output,
            expected_source_sha256,
            output_options,
        } => run_apply_command(
            &input,
            &patch,
            &output,
            expected_source_sha256,
            output_options,
        )?,
        Command::RenderArtifact {
            input,
            output_dir,
            output,
        } => run_render_artifact_command(&input, &output_dir, output)?,
        Command::Verify {
            input,
            checks,
            render_output_dir,
            output,
        } => run_verify_command(&input, checks, render_output_dir, output)?,
        Command::Markdown { file, common } | Command::Render { file, common } => {
            run_markdown_command(&file, &common)?;
        }
        Command::Diff {
            file_a,
            file_b,
            common,
        } => run_diff_command(&file_a, &file_b, &common)?,
        Command::Convert {
            input,
            output,
            common,
        } => run_convert_command(&input, output, common)?,
        Command::Stream { input, common } => run_stream_command(&input, common)?,
        Command::Inspect {
            input,
            common,
            agent_query,
        } => {
            if let Some(agent_query) = agent_query {
                run_agent_inspect_command(&input, common, &agent_query)?;
            } else {
                run_inspect_command(&input, common)?;
            }
        }
        Command::Plan { input, common } => run_plan_command(&input, common)?,
        Command::Create { output, input } => run_create_command(&output, &input)?,
    }

    Ok(())
}

fn main() {
    if let Err(err) = run() {
        eprintln!("Error: {err}");
        std::process::exit(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    fn build_zip(parts: Vec<(&str, &str)>) -> Vec<u8> {
        let mut buffer = Vec::new();
        let mut writer = ZipWriter::new(std::io::Cursor::new(&mut buffer));
        let options: FileOptions<'_, ()> = FileOptions::default();

        for (path, contents) in parts {
            writer.start_file(path, options).unwrap();
            writer.write_all(contents.as_bytes()).unwrap();
        }

        writer.finish().unwrap();
        buffer
    }

    fn build_test_xlsx_for_inspect() -> Vec<u8> {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#;
        let summary = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>OK</t></is></c></row>
  </sheetData>
</worksheet>"#;
        let data = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
  </sheetData>
</worksheet>"#;

        build_zip(vec![
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", summary),
            ("xl/worksheets/sheet2.xml", data),
        ])
    }

    fn markdown_common_options() -> CommonOptions {
        CommonOptions {
            format: None,
            output: CommonOutputOptions {
                output_format: Some(OutputFormatArg::Markdown),
                pretty: false,
            },
            sheets: None,
            pages: None,
            slides: None,
            pdf: PdfCliOptions { force: false },
            xlsx: XlsxCliOptions {
                style_aware: false,
                streaming: false,
            },
            include: MarkdownIncludeCliOptions {
                document_properties: false,
                no_headers_footers: false,
                no_formulas: false,
                no_frontmatter: false,
            },
            table: MarkdownTableCliOptions {
                no_first_row_header: false,
            },
            markdown_style: MarkdownStyleArg::Compact,
        }
    }

    #[test]
    fn detects_format_from_path_extension() {
        assert_eq!(
            detect_format_from_path(Path::new("report.docx")),
            Some(DocumentFormat::Docx)
        );
        assert_eq!(
            detect_format_from_path(Path::new("sheet.XLSX")),
            Some(DocumentFormat::Xlsx)
        );
        assert_eq!(
            detect_format_from_path(Path::new("table.csv")),
            Some(DocumentFormat::Csv)
        );
        assert_eq!(
            detect_format_from_path(Path::new("scan.pdf")),
            Some(DocumentFormat::Pdf)
        );
        assert_eq!(detect_format_from_path(Path::new("notes.txt")), None);
    }

    #[test]
    fn resolves_format_with_explicit_value() {
        let resolved = resolve_format(b"not a zip", None, Some(FormatArg::Pptx)).unwrap();
        assert_eq!(resolved, DocumentFormat::Pptx);

        let csv_resolved = resolve_format(b"col1,col2\n1,2\n", None, Some(FormatArg::Csv)).unwrap();
        assert_eq!(csv_resolved, DocumentFormat::Csv);
    }

    #[test]
    fn detects_format_from_package_contents() {
        let docx = build_zip(vec![("word/document.xml", "<w:document/>")]);
        let xlsx = build_zip(vec![("xl/workbook.xml", "<workbook/>")]);
        let pptx = build_zip(vec![("ppt/presentation.xml", "<p:presentation/>")]);

        assert_eq!(
            detect_format_from_bytes(&docx).unwrap(),
            DocumentFormat::Docx
        );
        assert_eq!(
            detect_format_from_bytes(&xlsx).unwrap(),
            DocumentFormat::Xlsx
        );
        assert_eq!(
            detect_format_from_bytes(&pptx).unwrap(),
            DocumentFormat::Pptx
        );
        assert_eq!(
            detect_format_from_bytes(b"%PDF-1.7\n").unwrap(),
            DocumentFormat::Pdf
        );
    }

    #[test]
    fn parses_number_ranges() {
        assert_eq!(parse_number_ranges("1,3,5").unwrap(), vec![1, 3, 5]);
        assert_eq!(parse_number_ranges("1-3,5").unwrap(), vec![1, 2, 3, 5]);
        assert_eq!(parse_number_ranges("2-4").unwrap(), vec![2, 3, 4]);
        assert!(parse_number_ranges("abc").is_err());
        assert!(parse_number_ranges("3-1").is_err());
    }

    #[test]
    fn parse_number_ranges_rejects_oversized_expansion() {
        let err = parse_number_ranges("1-100001").expect_err("expected size guard");
        assert!(err.contains("too large"));
    }

    #[test]
    fn parses_sheet_filter_names_and_indices() {
        let filter = parse_sheet_filter("Summary,2,  Data ,0");
        assert!(filter.names.contains("Summary"));
        assert!(filter.names.contains("Data"));
        assert!(filter.indices_1_based.contains(&2));
        assert!(filter.indices_1_based.contains(&0));
    }

    #[test]
    fn builds_sheet_filter_from_pages_as_indices() {
        let filter = build_sheet_filter(Some("Summary"), Some("2-3"))
            .expect("valid sheet filter")
            .expect("filter");
        assert!(filter.names.contains("Summary"));
        assert!(filter.indices_1_based.contains(&2));
        assert!(filter.indices_1_based.contains(&3));
    }

    #[test]
    fn parses_positive_u32_page_ranges() {
        assert_eq!(parse_positive_u32_ranges("1,3-4").unwrap(), vec![1, 3, 4]);
        let err = parse_positive_u32_ranges("0").expect_err("zero page should fail");
        assert!(err.contains(">= 1"));
    }

    #[test]
    fn builds_xlsx_inspect_info_with_sheet_filter() {
        let content = build_test_xlsx_for_inspect();
        let info = build_xlsx_inspect_info(&content, Some("2")).expect("inspect xlsx");
        let sheets = info.sheets.expect("sheets");
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Data");
        assert_eq!(sheets[0].rows, 2);
        assert_eq!(sheets[0].cols, 2);
    }

    #[test]
    fn renders_markdown_output_for_xlsx_document() {
        let content = build_test_xlsx_for_inspect();
        let common = markdown_common_options();
        let doc = extract_ir_document(&content, DocumentFormat::Xlsx, &common).expect("extract");
        let markdown = render_output(&doc, &common).expect("render markdown");
        assert!(markdown.contains("## Sheet: Summary"));
        assert!(markdown.contains("## Sheet: Data"));
    }

    #[test]
    fn xlsx_markdown_fast_path_rejects_slides_option() {
        let content = build_test_xlsx_for_inspect();
        let mut path = std::env::temp_dir();
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system time")
            .as_nanos();
        path.push(format!("officemd-test-{unique}.xlsx"));
        std::fs::write(&path, content).expect("write xlsx fixture");

        let mut common = markdown_common_options();
        common.slides = Some("1".to_string());
        let err = extract_markdown_from_file(&path, &common).expect_err("slides should fail");
        let _ = std::fs::remove_file(&path);

        assert_eq!(err, "--slides can only be used with PPTX files");
    }

    #[test]
    fn renders_markdown_output_for_csv_document() {
        let content = b"name,value\nwidget,42\n";
        let common = markdown_common_options();
        let doc = extract_ir_document(content, DocumentFormat::Csv, &common).expect("extract");
        let markdown = render_output(&doc, &common).expect("render markdown");
        assert!(markdown.contains("## Sheet: Sheet1"));
        assert!(markdown.contains("| name | value |"));
    }

    #[test]
    fn default_output_path_uses_correct_extension() {
        assert_eq!(
            default_output_path(Path::new("doc.docx"), OutputFormatArg::Markdown),
            PathBuf::from("doc.md")
        );
        assert_eq!(
            default_output_path(Path::new("doc.docx"), OutputFormatArg::Json),
            PathBuf::from("doc.json")
        );
    }

    #[test]
    fn builds_pdf_inspect_info_and_text_output() {
        let doc = OoxmlDocument {
            kind: officemd_core::ir::DocumentKind::Pdf,
            pdf: Some(officemd_core::ir::PdfDocument {
                pages: vec![],
                diagnostics: officemd_core::ir::PdfDiagnostics {
                    classification: officemd_core::ir::PdfClassification::Scanned,
                    confidence: 0.75,
                    page_count: 2,
                    pages_needing_ocr: vec![1, 2],
                    has_encoding_issues: false,
                },
            }),
            ..Default::default()
        };

        let info = build_inspect_info(&doc, DocumentFormat::Pdf);
        assert_eq!(info.format, "pdf");
        let text = render_inspect_text(&info);
        assert!(text.contains("PDF Diagnostics"));
        assert!(text.contains("Classification: Scanned"));
        assert!(text.contains("Pages needing OCR: 1, 2"));
    }

    #[test]
    fn agent_plan_includes_sheet_selector_commands() {
        let inspect = InspectInfo {
            format: "xlsx".to_string(),
            sections: None,
            sheets: Some(vec![SheetInfo {
                name: "Summary".to_string(),
                rows: 2,
                cols: 3,
            }]),
            slides: None,
            pdf: None,
        };
        let common = markdown_common_options();

        let plan = build_agent_plan(
            Path::new("book.xlsx"),
            DocumentFormat::Xlsx,
            &common,
            inspect,
        );

        assert_eq!(plan.format, "xlsx");
        assert_eq!(plan.selectors.len(), 1);
        let sheet = &plan.selectors[0].values[0];
        assert_eq!(sheet.value, "1");
        assert_eq!(sheet.label, "Summary");
        assert_eq!(
            sheet.extract_json,
            vec![
                "officemd",
                "stream",
                "book.xlsx",
                "--format",
                "xlsx",
                "--output-format",
                "json",
                "--pretty",
                "--sheets",
                "1",
            ]
        );
    }

    #[test]
    fn agent_plan_includes_pdf_page_selector_and_ocr_warning() {
        let inspect = InspectInfo {
            format: "pdf".to_string(),
            sections: None,
            sheets: None,
            slides: None,
            pdf: Some(PdfInfo {
                classification: "Mixed".to_string(),
                confidence: 0.8,
                page_count: 3,
                pages_needing_ocr: vec![2],
                has_encoding_issues: true,
            }),
        };
        let common = markdown_common_options();

        let plan = build_agent_plan(
            Path::new("report.pdf"),
            DocumentFormat::Pdf,
            &common,
            inspect,
        );

        assert_eq!(plan.selectors[0].kind, "page");
        assert!(
            plan.selectors[0]
                .values
                .iter()
                .any(|value| value.value == "2" && value.needs_ocr == Some(true))
        );
        assert_eq!(plan.warnings.len(), 2);
    }
}
