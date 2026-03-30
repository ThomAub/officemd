use officemd_core::{
    DocxPatch, DocxTextScope, PptxPatch, PptxTextScope, ScopedDocxReplace, ScopedPptxReplace,
    TextReplace, patch_docx_batch_with_report, patch_pptx_batch_with_report,
};
use serde_json::json;
use std::fs;
use std::path::PathBuf;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("examples dir")
        .to_path_buf()
}

fn main() {
    let root = repo_root();
    let data_dir = root.join("data");
    let out_dir = root.join("out").join("batch_rust");
    fs::create_dir_all(&out_dir).expect("create out dir");

    println!("Plan:");
    println!("1. Load showcase DOCX/PPTX fixture bytes");
    println!(
        "2. Patch batches with officemd_core::patch_docx_batch_with_report / patch_pptx_batch_with_report"
    );
    println!("3. Write outputs and print per-file replacement counters");
    println!();

    let docx_bytes = fs::read(data_dir.join("showcase.docx")).expect("read docx");
    let docx_patch = DocxPatch {
        set_core_title: Some("Edited DOCX Showcase Batch From Rust".to_string()),
        replace_body_title: None,
        scoped_replacements: vec![
            ScopedDocxReplace {
                scope: DocxTextScope::Headers,
                replace: TextReplace::all(
                    "OOXML Showcase Header",
                    "OfficeMD Showcase Header — batch edited from Rust",
                ),
            },
            ScopedDocxReplace {
                scope: DocxTextScope::Body,
                replace: TextReplace::first(
                    "Quarterly Operations Summary",
                    "Quarterly Operations Summary — batch edited from Rust",
                ),
            },
            ScopedDocxReplace {
                scope: DocxTextScope::Comments,
                replace: TextReplace::all(
                    "Example DOCX comment captured as markdown footnote.",
                    "Edited DOCX comment from Rust batch patch API.",
                ),
            },
        ],
    };
    let docx_results =
        patch_docx_batch_with_report(vec![docx_bytes.clone(), docx_bytes], &docx_patch, Some(2))
            .expect("patch docx batch");

    let pptx_bytes = fs::read(data_dir.join("showcase.pptx")).expect("read pptx");
    let pptx_patch = PptxPatch {
        set_core_title: Some("Edited PPTX Showcase Batch From Rust".to_string()),
        scoped_replacements: vec![
            ScopedPptxReplace {
                scope: PptxTextScope::AllText,
                replace: TextReplace::first(
                    "Quarterly Review",
                    "Quarterly Review — batch edited from Rust",
                ),
            },
            ScopedPptxReplace {
                scope: PptxTextScope::Comments,
                replace: TextReplace::all(
                    "Add one slide on operating margin.",
                    "Edited PPTX comment from Rust batch patch API.",
                ),
            },
        ],
    };
    let pptx_results =
        patch_pptx_batch_with_report(vec![pptx_bytes.clone(), pptx_bytes], &pptx_patch, Some(2))
            .expect("patch pptx batch");

    for (idx, item) in docx_results.iter().enumerate() {
        let path = out_dir.join(format!("showcase_batch_{idx}.docx"));
        fs::write(&path, &item.content).expect("write docx batch result");
        println!(
            "{}",
            serde_json::to_string_pretty(&json!({
                "path": path.display().to_string(),
                "format": "docx",
                "parts_scanned": item.report.parts_scanned,
                "parts_modified": item.report.parts_modified,
                "replacements_applied": item.report.replacements_applied,
            }))
            .unwrap()
        );
    }

    for (idx, item) in pptx_results.iter().enumerate() {
        let path = out_dir.join(format!("showcase_batch_{idx}.pptx"));
        fs::write(&path, &item.content).expect("write pptx batch result");
        println!(
            "{}",
            serde_json::to_string_pretty(&json!({
                "path": path.display().to_string(),
                "format": "pptx",
                "parts_scanned": item.report.parts_scanned,
                "parts_modified": item.report.parts_modified,
                "replacements_applied": item.report.replacements_applied,
            }))
            .unwrap()
        );
    }
}
