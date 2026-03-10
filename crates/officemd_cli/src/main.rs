use clap::{Parser, Subcommand, ValueEnum};
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
    command: Command,
}

/// Shared options used by convert, stream, and inspect.
#[derive(clap::Args, Debug, Clone)]
#[allow(clippy::struct_excessive_bools)]
struct CommonOptions {
    /// Explicitly set the document format.
    #[arg(long, value_enum)]
    format: Option<FormatArg>,

    /// Include document properties in markdown output.
    #[arg(long, default_value_t = false)]
    include_document_properties: bool,

    /// Output format: markdown (default) or json.
    #[arg(long, value_enum, default_value_t = OutputFormatArg::Markdown)]
    output_format: OutputFormatArg,

    /// Pretty-print JSON output.
    #[arg(long, default_value_t = false)]
    pretty: bool,

    /// Filter XLSX sheets by name or 1-based index (comma-separated).
    #[arg(long)]
    sheets: Option<String>,

    /// Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5").
    #[arg(long)]
    pages: Option<String>,

    /// Filter PPTX slides by number or range (e.g. "1-3,5").
    #[arg(long)]
    slides: Option<String>,

    /// Force extraction even for scanned/image-based PDFs.
    #[arg(long, default_value_t = false)]
    force: bool,

    /// Use style-aware cell values for XLSX.
    #[arg(long, default_value_t = false)]
    style_aware: bool,

    /// Use streaming row parser for XLSX.
    #[arg(long, default_value_t = false)]
    streaming: bool,

    /// Omit DOCX header/footer sections from markdown output.
    #[arg(long, default_value_t = false)]
    no_headers_footers: bool,

    /// Use synthetic Col1/Col2 headers instead of first data row.
    #[arg(long, default_value_t = false)]
    no_first_row_header: bool,

    /// Markdown style profile.
    #[arg(long, value_enum, default_value_t = MarkdownStyleArg::Compact)]
    markdown_style: MarkdownStyleArg,
}

#[derive(Subcommand, Debug)]
enum Command {
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

// --- Inspect info types ---

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

// --- Format detection ---

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

// --- IR extraction ---

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
            let options = XlsxExtractOptions {
                style_aware_values: common.style_aware,
                streaming_rows: common.streaming,
                sheet_filter: common.sheets.as_deref().map(parse_sheet_filter),
                include_document_properties: common.include_document_properties
                    || common.output_format == OutputFormatArg::Json,
                trim_empty: matches!(common.markdown_style, MarkdownStyleArg::Compact),
            };
            officemd_xlsx::extract_tables_ir_with_options(content, &options)
                .map_err(|e| e.to_string())?
        }
        DocumentFormat::Csv => {
            let options = officemd_csv::table_ir::CsvExtractOptions {
                include_document_properties: common.include_document_properties
                    || common.output_format == OutputFormatArg::Json,
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
            officemd_pdf::extract_ir_force(content, common.force).map_err(|e| e.to_string())?
        }
    };

    // Warn about XLSX-specific flags used with non-XLSX formats
    if format != DocumentFormat::Xlsx {
        if common.style_aware {
            eprintln!("Warning: --style-aware is only effective with XLSX files");
        }
        if common.streaming {
            eprintln!("Warning: --streaming is only effective with XLSX files");
        }
    }

    Ok(doc)
}

// --- Filters ---

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

// --- Output rendering ---

fn render_output(doc: &OoxmlDocument, common: &CommonOptions) -> Result<String, String> {
    match common.output_format {
        OutputFormatArg::Markdown => {
            let markdown_profile = match common.markdown_style {
                MarkdownStyleArg::Compact => officemd_markdown::MarkdownProfile::LlmCompact,
                MarkdownStyleArg::Human => officemd_markdown::MarkdownProfile::Human,
            };
            let options = officemd_markdown::RenderOptions {
                include_document_properties: common.include_document_properties,
                use_first_row_as_header: !common.no_first_row_header,
                include_headers_footers: !common.no_headers_footers,
                markdown_profile,
            };
            Ok(officemd_markdown::render_document_with_options(
                doc, options,
            ))
        }
        OutputFormatArg::Json => {
            if common.pretty {
                serde_json::to_string_pretty(doc).map_err(|e| e.to_string())
            } else {
                serde_json::to_string(doc).map_err(|e| e.to_string())
            }
        }
    }
}

// --- Inspect ---

#[allow(clippy::similar_names)]
fn build_xlsx_inspect_info(
    content: &[u8],
    sheets_filter: Option<&str>,
) -> Result<InspectInfo, String> {
    let sheet_filter = sheets_filter.map(parse_sheet_filter);
    let sheets =
        inspect_sheet_summaries(content, sheet_filter.as_ref()).map_err(|e| e.to_string())?;
    Ok(InspectInfo {
        format: "xlsx".to_string(),
        sections: None,
        sheets: Some(
            sheets
                .into_iter()
                .map(|sheet| SheetInfo {
                    name: sheet.name,
                    rows: sheet.rows,
                    cols: sheet.cols,
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

// --- I/O helpers ---

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

// --- Helpers for markdown/render/diff ---

fn extract_markdown_from_file(path: &Path, common: &CommonOptions) -> Result<String, String> {
    let bytes =
        std::fs::read(path).map_err(|e| format!("failed to read '{}': {e}", path.display()))?;
    let resolved = resolve_format(&bytes, Some(path), common.format)?;
    let doc = extract_ir_document(&bytes, resolved, common)?;

    // Force markdown output regardless of --output-format
    let markdown_profile = match common.markdown_style {
        MarkdownStyleArg::Compact => officemd_markdown::MarkdownProfile::LlmCompact,
        MarkdownStyleArg::Human => officemd_markdown::MarkdownProfile::Human,
    };
    let options = officemd_markdown::RenderOptions {
        include_document_properties: common.include_document_properties,
        use_first_row_as_header: !common.no_first_row_header,
        include_headers_footers: !common.no_headers_footers,
        markdown_profile,
    };
    let md = officemd_markdown::render_document_with_options(&doc, options);

    // --pages for XLSX/CSV acts as sheet index selector: hint users to use --sheets
    if common.pages.is_some()
        && (resolved == DocumentFormat::Xlsx || resolved == DocumentFormat::Csv)
    {
        eprintln!(
            "Hint: use --sheets for sheet selection with {} files",
            resolved
        );
    }

    // Warn about scanned PDFs
    if resolved == DocumentFormat::Pdf && md.trim().len() < 50 {
        warn_scanned_pdf(&bytes, common.force);
    }

    Ok(md)
}

fn warn_scanned_pdf(content: &[u8], force: bool) {
    let diagnostics = match officemd_pdf::inspect_pdf(content) {
        Ok(d) => d,
        Err(_) => return,
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

fn render_colored_diff(text_a: &str, text_b: &str, label_a: &str, label_b: &str) -> String {
    let diff = TextDiff::from_lines(text_a, text_b);
    let mut out = String::new();

    // ANSI colors
    const RED: &str = "\x1b[31m";
    const GREEN: &str = "\x1b[32m";
    const CYAN: &str = "\x1b[36m";
    const RESET: &str = "\x1b[0m";

    for hunk in diff.unified_diff().header(label_a, label_b).iter_hunks() {
        for change in hunk.iter_changes() {
            let (color, sign) = match change.tag() {
                ChangeTag::Delete => (RED, "-"),
                ChangeTag::Insert => (GREEN, "+"),
                ChangeTag::Equal => ("", " "),
            };
            if color.is_empty() {
                let _ = write!(out, "{sign}{change}");
            } else {
                let _ = write!(out, "{color}{sign}{change}{RESET}");
            }
            if change.missing_newline() {
                out.push('\n');
            }
        }
        let _ = writeln!(out, "{CYAN}---{RESET}");
    }

    if out.is_empty() {
        out.push_str("(no differences)\n");
    }

    out
}

// --- Main ---

fn run() -> Result<(), String> {
    let cli = Cli::parse();

    match cli.command {
        Command::Markdown { file, common } => {
            let md = extract_markdown_from_file(&file, &common)?;
            let mut stdout = std::io::stdout().lock();
            stdout
                .write_all(md.as_bytes())
                .map_err(|e| format!("failed to write stdout: {e}"))?;
        }
        Command::Render { file, common } => {
            let md = extract_markdown_from_file(&file, &common)?;
            let mut stdout = std::io::stdout().lock();
            stdout
                .write_all(md.as_bytes())
                .map_err(|e| format!("failed to write stdout: {e}"))?;
        }
        Command::Diff {
            file_a,
            file_b,
            common,
        } => {
            let md_a = extract_markdown_from_file(&file_a, &common)?;
            let md_b = extract_markdown_from_file(&file_b, &common)?;
            let label_a = file_a.display().to_string();
            let label_b = file_b.display().to_string();
            let output = render_colored_diff(&md_a, &md_b, &label_a, &label_b);
            let mut stdout = std::io::stdout().lock();
            stdout
                .write_all(output.as_bytes())
                .map_err(|e| format!("failed to write stdout: {e}"))?;
        }
        Command::Convert {
            input,
            output,
            common,
        } => {
            let bytes = std::fs::read(&input)
                .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?;
            let resolved = resolve_format(&bytes, Some(&input), common.format)?;
            let doc = extract_ir_document(&bytes, resolved, &common)?;
            let rendered = render_output(&doc, &common)?;
            let output_path =
                output.unwrap_or_else(|| default_output_path(&input, common.output_format));
            std::fs::write(&output_path, &rendered)
                .map_err(|e| format!("failed to write output '{}': {e}", output_path.display()))?;
            let format_label = match common.output_format {
                OutputFormatArg::Markdown => "markdown",
                OutputFormatArg::Json => "JSON",
            };
            eprintln!(
                "Wrote {} for {} document to {}",
                format_label,
                resolved,
                output_path.display()
            );
        }
        Command::Stream { input, common } => {
            let use_stdin = input == Path::new("-");
            let bytes = if use_stdin {
                read_all_from_stdin()?
            } else {
                std::fs::read(&input)
                    .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?
            };

            let resolved = resolve_format(
                &bytes,
                if use_stdin { None } else { Some(&input) },
                common.format,
            )?;
            let doc = extract_ir_document(&bytes, resolved, &common)?;
            drop(bytes);
            let rendered = render_output(&doc, &common)?;

            let mut stdout = std::io::stdout().lock();
            stdout
                .write_all(rendered.as_bytes())
                .map_err(|e| format!("failed to write stdout: {e}"))?;
        }
        Command::Inspect { input, common } => {
            let bytes = std::fs::read(&input)
                .map_err(|e| format!("failed to read input '{}': {e}", input.display()))?;
            let resolved = resolve_format(&bytes, Some(&input), common.format)?;
            let info = if resolved == DocumentFormat::Xlsx {
                if common.slides.is_some() {
                    return Err("--slides can only be used with PPTX files".to_string());
                }
                build_xlsx_inspect_info(&bytes, common.sheets.as_deref())?
            } else if resolved == DocumentFormat::Pdf {
                if common.sheets.is_some() {
                    eprintln!("Warning: --sheets is ignored for PDF files");
                }
                if common.slides.is_some() {
                    eprintln!("Warning: --slides is ignored for PDF files");
                }
                build_pdf_inspect_info(&bytes)?
            } else {
                let doc = extract_ir_document(&bytes, resolved, &common)?;
                build_inspect_info(&doc, resolved)
            };

            let output = match common.output_format {
                OutputFormatArg::Json => {
                    if common.pretty {
                        serde_json::to_string_pretty(&info).map_err(|e| e.to_string())?
                    } else {
                        serde_json::to_string(&info).map_err(|e| e.to_string())?
                    }
                }
                OutputFormatArg::Markdown => render_inspect_text(&info),
            };

            let mut stdout = std::io::stdout().lock();
            stdout
                .write_all(output.as_bytes())
                .map_err(|e| format!("failed to write stdout: {e}"))?;
        }
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
            include_document_properties: false,
            output_format: OutputFormatArg::Markdown,
            pretty: false,
            sheets: None,
            pages: None,
            slides: None,
            force: false,
            style_aware: false,
            streaming: false,
            no_headers_footers: false,
            no_first_row_header: false,
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
}
