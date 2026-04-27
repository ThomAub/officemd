//! Compare XLSX markdown output with and without `trim_empty`.
//!
//! Usage:
//!   `cargo run -p officemd_examples --bin compare_trim_xlsx -- <path/to/file.xlsx>`

use std::{env, fs};

use officemd_markdown::{MarkdownProfile, RenderOptions};
use officemd_xlsx::table_ir::{
    XlsxExtractOptions, XlsxTrimOptions, extract_tables_ir_with_options,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).unwrap_or_else(|| {
        eprintln!(
            "Usage: cargo run -p officemd_examples --bin compare_trim_xlsx -- <path/to/file.xlsx>"
        );
        std::process::exit(2);
    });

    let content = fs::read(&path)?;

    let doc_no_trim = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())?;
    let md_no_trim = officemd_markdown::render_document_with_options(
        &doc_no_trim,
        RenderOptions {
            markdown_profile: MarkdownProfile::LlmCompact,
            ..Default::default()
        },
    );

    let doc_trim = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            trim: XlsxTrimOptions { empty_edges: true },
            ..Default::default()
        },
    )?;
    let md_trim = officemd_markdown::render_document_with_options(
        &doc_trim,
        RenderOptions {
            markdown_profile: MarkdownProfile::LlmCompact,
            ..Default::default()
        },
    );

    println!("=== {path} ===\n");

    println!("--- WITHOUT trim_empty ({} chars) ---", md_no_trim.len());
    for sheet in &doc_no_trim.sheets {
        let t = &sheet.tables[0];
        println!(
            "  Sheet {:?}: {} headers, {} rows",
            sheet.name,
            t.headers.len(),
            t.rows.len()
        );
    }
    println!();
    print!("{md_no_trim}");
    println!();

    println!("--- WITH trim_empty ({} chars) ---", md_trim.len());
    for sheet in &doc_trim.sheets {
        let t = &sheet.tables[0];
        println!(
            "  Sheet {:?}: {} headers, {} rows",
            sheet.name,
            t.headers.len(),
            t.rows.len()
        );
    }
    println!();
    print!("{md_trim}");
    println!();

    let no_trim_len = md_no_trim.len();
    let trim_len = md_trim.len();
    let (saved_prefix, saved) = if no_trim_len >= trim_len {
        ("", no_trim_len - trim_len)
    } else {
        ("-", trim_len - no_trim_len)
    };
    let pct_tenths = saved
        .saturating_mul(1_000)
        .checked_div(no_trim_len)
        .unwrap_or(0);
    let pct_whole = pct_tenths / 10;
    let pct_fraction = pct_tenths % 10;
    println!(
        "--- Token savings: {saved_prefix}{saved} chars ({pct_whole}.{pct_fraction}% reduction) ---"
    );

    Ok(())
}
