//! Compare XLSX markdown output with and without `trim_empty`.
//!
//! Usage:
//!   cargo run -p officemd_examples --bin compare_trim_xlsx -- <path/to/file.xlsx>

use std::{env, fs};

use officemd_markdown::{MarkdownProfile, RenderOptions};
use officemd_xlsx::table_ir::{XlsxExtractOptions, extract_tables_ir_with_options};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).unwrap_or_else(|| {
        eprintln!(
            "Usage: cargo run -p officemd_examples --bin compare_trim_xlsx -- <path/to/file.xlsx>"
        );
        std::process::exit(2);
    });

    let content = fs::read(&path)?;

    // --- Extract without trim ---
    let doc_no_trim = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            trim_empty: false,
            ..Default::default()
        },
    )?;
    let md_no_trim = officemd_markdown::render_document_with_options(
        &doc_no_trim,
        RenderOptions {
            markdown_profile: MarkdownProfile::LlmCompact,
            ..Default::default()
        },
    );

    // --- Extract with trim ---
    let doc_trim = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            trim_empty: true,
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

    // --- Report ---
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

    let saved = md_no_trim.len() as i64 - md_trim.len() as i64;
    let pct = if md_no_trim.is_empty() {
        0.0
    } else {
        (saved as f64 / md_no_trim.len() as f64) * 100.0
    };
    println!("--- Token savings: {saved} chars ({pct:.1}% reduction) ---");

    Ok(())
}
