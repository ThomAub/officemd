use std::process::Command;

use insta_cmd::{assert_cmd_snapshot, get_cargo_bin};

fn cli() -> Command {
    Command::new(get_cargo_bin("officemd"))
}

fn bind_common_filters() -> insta::internals::SettingsBindDropGuard {
    let mut settings = insta::Settings::clone_current();
    settings.add_filter(r"\bofficemd\.exe\b", "officemd");
    settings.add_filter(r"(?m)[ \t]+$", "");
    settings.bind_to_scope()
}

#[test]
fn top_level_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().arg("--help"),
        @r###"
    success: true
    exit_code: 0
    ----- stdout -----
    Fast Office document extraction for LLMs and agents

    Usage: officemd [OPTIONS] [COMMAND]

    Commands:
      probe            Probe artifact identity, capabilities, and operational risks
      apply            Apply a typed agent patch plan to a new output artifact
      render-artifact  Render an artifact into visual evidence images when a backend is available
      verify           Verify semantic, structural, or visual artifact invariants
      markdown         Extract markdown, print to stdout
      render           Extract markdown, render to terminal with ANSI formatting
      diff             Diff markdown output of two documents
      convert          Convert a document file to an output file
      stream           Stream output to stdout from a file path or stdin
      inspect          Inspect document metadata without full content rendering
      plan             Emit an agent-friendly parsing plan with follow-up commands
      create           Create an Office document from markdown input
      help             Print this message or the help of the given subcommand(s)

    Options:
          --help-tree [<DEPTH>]  Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
      -h, --help                 Print help
      -V, --version              Print version

    ----- stderr -----
    "###,
    );
}

#[test]
fn probe_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["probe", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Probe artifact identity, capabilities, and operational risks

    Usage: officemd probe [OPTIONS] <INPUT>

    Arguments:
      <INPUT>  Input document path (.docx/.xlsx/.csv/.pptx/.pdf)

    Options:
          --output-format <OUTPUT_FORMAT>  Output format: markdown (default) or json [possible values: markdown, json]
          --pretty                         Pretty-print JSON output
          --format <FORMAT>                Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --help-tree [<DEPTH>]            Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
      -h, --help                           Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn markdown_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["markdown", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Extract markdown, print to stdout

    Usage: officemd markdown [OPTIONS] <FILE>

    Arguments:
      <FILE>  Path to an input document

    Options:
          --format <FORMAT>
              Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --output-format <OUTPUT_FORMAT>
              Output format: markdown (default) or json [possible values: markdown, json]
          --pretty
              Pretty-print JSON output
          --sheets <SHEETS>
              Filter XLSX sheets by name or 1-based index (comma-separated)
          --pages <PAGES>
              Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5")
          --slides <SLIDES>
              Filter PPTX slides by number or range (e.g. "1-3,5")
          --force
              Force extraction even for scanned/image-based PDFs
          --style-aware
              Use style-aware cell values for XLSX
          --streaming
              Use streaming row parser for XLSX
          --include-document-properties
              Include document properties in markdown output
          --no-headers-footers
              Omit DOCX header/footer sections from markdown output
          --no-formulas
              Omit XLSX formula footnotes from markdown output
          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
          --no-frontmatter
              Omit the leading `<!-- officemd: ... -->` frontmatter comment
          --no-first-row-header
              Use synthetic Col1/Col2 headers instead of first data row
          --markdown-style <MARKDOWN_STYLE>
              Markdown style profile [default: compact] [possible values: compact, human]
      -h, --help
              Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn render_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["render", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Extract markdown, render to terminal with ANSI formatting

    Usage: officemd render [OPTIONS] <FILE>

    Arguments:
      <FILE>  Path to an input document

    Options:
          --format <FORMAT>
              Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --output-format <OUTPUT_FORMAT>
              Output format: markdown (default) or json [possible values: markdown, json]
          --pretty
              Pretty-print JSON output
          --sheets <SHEETS>
              Filter XLSX sheets by name or 1-based index (comma-separated)
          --pages <PAGES>
              Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5")
          --slides <SLIDES>
              Filter PPTX slides by number or range (e.g. "1-3,5")
          --force
              Force extraction even for scanned/image-based PDFs
          --style-aware
              Use style-aware cell values for XLSX
          --streaming
              Use streaming row parser for XLSX
          --include-document-properties
              Include document properties in markdown output
          --no-headers-footers
              Omit DOCX header/footer sections from markdown output
          --no-formulas
              Omit XLSX formula footnotes from markdown output
          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
          --no-frontmatter
              Omit the leading `<!-- officemd: ... -->` frontmatter comment
          --no-first-row-header
              Use synthetic Col1/Col2 headers instead of first data row
          --markdown-style <MARKDOWN_STYLE>
              Markdown style profile [default: compact] [possible values: compact, human]
      -h, --help
              Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn diff_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["diff", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Diff markdown output of two documents

    Usage: officemd diff [OPTIONS] <FILE_A> <FILE_B>

    Arguments:
      <FILE_A>  Path to first input document
      <FILE_B>  Path to second input document

    Options:
          --format <FORMAT>
              Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --output-format <OUTPUT_FORMAT>
              Output format: markdown (default) or json [possible values: markdown, json]
          --pretty
              Pretty-print JSON output
          --sheets <SHEETS>
              Filter XLSX sheets by name or 1-based index (comma-separated)
          --pages <PAGES>
              Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5")
          --slides <SLIDES>
              Filter PPTX slides by number or range (e.g. "1-3,5")
          --force
              Force extraction even for scanned/image-based PDFs
          --style-aware
              Use style-aware cell values for XLSX
          --streaming
              Use streaming row parser for XLSX
          --include-document-properties
              Include document properties in markdown output
          --no-headers-footers
              Omit DOCX header/footer sections from markdown output
          --no-formulas
              Omit XLSX formula footnotes from markdown output
          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
          --no-frontmatter
              Omit the leading `<!-- officemd: ... -->` frontmatter comment
          --no-first-row-header
              Use synthetic Col1/Col2 headers instead of first data row
          --markdown-style <MARKDOWN_STYLE>
              Markdown style profile [default: compact] [possible values: compact, human]
      -h, --help
              Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn convert_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["convert", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Convert a document file to an output file

    Usage: officemd convert [OPTIONS] <INPUT>

    Arguments:
      <INPUT>  Input document path (.docx/.xlsx/.csv/.pptx/.pdf)

    Options:
      -o, --output <OUTPUT>
              Output file path. Defaults to <input>.md or <input>.json
          --format <FORMAT>
              Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --output-format <OUTPUT_FORMAT>
              Output format: markdown (default) or json [possible values: markdown, json]
          --pretty
              Pretty-print JSON output
          --sheets <SHEETS>
              Filter XLSX sheets by name or 1-based index (comma-separated)
          --pages <PAGES>
              Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5")
          --slides <SLIDES>
              Filter PPTX slides by number or range (e.g. "1-3,5")
          --force
              Force extraction even for scanned/image-based PDFs
          --style-aware
              Use style-aware cell values for XLSX
          --streaming
              Use streaming row parser for XLSX
          --include-document-properties
              Include document properties in markdown output
          --no-headers-footers
              Omit DOCX header/footer sections from markdown output
          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
          --no-formulas
              Omit XLSX formula footnotes from markdown output
          --no-frontmatter
              Omit the leading `<!-- officemd: ... -->` frontmatter comment
          --no-first-row-header
              Use synthetic Col1/Col2 headers instead of first data row
          --markdown-style <MARKDOWN_STYLE>
              Markdown style profile [default: compact] [possible values: compact, human]
      -h, --help
              Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn stream_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["stream", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Stream output to stdout from a file path or stdin

    Usage: officemd stream [OPTIONS] [INPUT]

    Arguments:
      [INPUT]  Input path or '-' for stdin [default: -]

    Options:
          --format <FORMAT>
              Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --output-format <OUTPUT_FORMAT>
              Output format: markdown (default) or json [possible values: markdown, json]
          --pretty
              Pretty-print JSON output
          --sheets <SHEETS>
              Filter XLSX sheets by name or 1-based index (comma-separated)
          --pages <PAGES>
              Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5")
          --slides <SLIDES>
              Filter PPTX slides by number or range (e.g. "1-3,5")
          --force
              Force extraction even for scanned/image-based PDFs
          --style-aware
              Use style-aware cell values for XLSX
          --streaming
              Use streaming row parser for XLSX
          --include-document-properties
              Include document properties in markdown output
          --no-headers-footers
              Omit DOCX header/footer sections from markdown output
          --no-formulas
              Omit XLSX formula footnotes from markdown output
          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
          --no-frontmatter
              Omit the leading `<!-- officemd: ... -->` frontmatter comment
          --no-first-row-header
              Use synthetic Col1/Col2 headers instead of first data row
          --markdown-style <MARKDOWN_STYLE>
              Markdown style profile [default: compact] [possible values: compact, human]
      -h, --help
              Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn inspect_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["inspect", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Inspect document metadata without full content rendering

    Usage: officemd inspect [OPTIONS] <INPUT>

    Arguments:
      <INPUT>  Input document path (.docx/.xlsx/.csv/.pptx/.pdf)

    Options:
          --format <FORMAT>
              Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --output-format <OUTPUT_FORMAT>
              Output format: markdown (default) or json [possible values: markdown, json]
          --pretty
              Pretty-print JSON output
          --sheets <SHEETS>
              Filter XLSX sheets by name or 1-based index (comma-separated)
          --pages <PAGES>
              Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5")
          --slides <SLIDES>
              Filter PPTX slides by number or range (e.g. "1-3,5")
          --force
              Force extraction even for scanned/image-based PDFs
          --style-aware
              Use style-aware cell values for XLSX
          --streaming
              Use streaming row parser for XLSX
          --include-document-properties
              Include document properties in markdown output
          --no-headers-footers
              Omit DOCX header/footer sections from markdown output
          --no-formulas
              Omit XLSX formula footnotes from markdown output
          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
          --no-frontmatter
              Omit the leading `<!-- officemd: ... -->` frontmatter comment
          --no-first-row-header
              Use synthetic Col1/Col2 headers instead of first data row
          --markdown-style <MARKDOWN_STYLE>
              Markdown style profile [default: compact] [possible values: compact, human]
          --agent-query <AGENT_QUERY>
              JSON file containing an agent InspectQuery
      -h, --help
              Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn plan_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["plan", "--help"]),
        @r#"
    success: true
    exit_code: 0
    ----- stdout -----
    Emit an agent-friendly parsing plan with follow-up commands

    Usage: officemd plan [OPTIONS] <INPUT>

    Arguments:
      <INPUT>  Input document path (.docx/.xlsx/.csv/.pptx/.pdf)

    Options:
          --format <FORMAT>
              Explicitly set the document format [possible values: docx, xlsx, csv, pptx, pdf]
          --output-format <OUTPUT_FORMAT>
              Output format: markdown (default) or json [possible values: markdown, json]
          --pretty
              Pretty-print JSON output
          --sheets <SHEETS>
              Filter XLSX sheets by name or 1-based index (comma-separated)
          --pages <PAGES>
              Select PDF pages/PPTX slides or XLSX/CSV sheet indices (e.g. "1,3-5")
          --slides <SLIDES>
              Filter PPTX slides by number or range (e.g. "1-3,5")
          --force
              Force extraction even for scanned/image-based PDFs
          --style-aware
              Use style-aware cell values for XLSX
          --streaming
              Use streaming row parser for XLSX
          --include-document-properties
              Include document properties in markdown output
          --no-headers-footers
              Omit DOCX header/footer sections from markdown output
          --no-formulas
              Omit XLSX formula footnotes from markdown output
          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options
          --no-frontmatter
              Omit the leading `<!-- officemd: ... -->` frontmatter comment
          --no-first-row-header
              Use synthetic Col1/Col2 headers instead of first data row
          --markdown-style <MARKDOWN_STYLE>
              Markdown style profile [default: compact] [possible values: compact, human]
      -h, --help
              Print help

    ----- stderr -----
    "#,
    );
}

#[test]
fn create_help_snapshot() {
    let _guard = bind_common_filters();
    assert_cmd_snapshot!(
        cli().args(["create", "--help"]),
        @r###"
    success: true
    exit_code: 0
    ----- stdout -----
    Create an Office document from markdown input.

    Reads officemd-flavored markdown and generates a .docx, .xlsx, or .pptx file. The output format is detected from the file extension.

    Examples: officemd create report.docx < input.md officemd create data.xlsx --input table.md officemd create slides.pptx --input deck.md

    Usage: officemd create [OPTIONS] <OUTPUT>

    Arguments:
      <OUTPUT>
              Output file path (.docx, .xlsx, .pptx)

    Options:
      -i, --input <INPUT>
              Input markdown file, or '-' for stdin (default)

              [default: -]

          --help-tree [<DEPTH>]
              Show commands and options in a tree format. Depth 1 shows commands only, depth 2 includes arguments and options

      -h, --help
              Print help (see a summary with '-h')

    ----- stderr -----
    "###,
    );
}
