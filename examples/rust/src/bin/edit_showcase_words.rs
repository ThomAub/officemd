use officemd_core::{
    DocxPatch, DocxTextScope, MatchPolicy, PptxPatch, PptxTextScope, ScopedDocxReplace,
    ScopedPptxReplace, ScopedXlsxReplace, TextReplace, XlsxPatch, XlsxSheetRename, XlsxTextScope,
    patch_docx, patch_pptx, patch_xlsx_with_report,
};
use serde_json::json;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("examples dir")
        .to_path_buf()
}

fn check_libreoffice(path: &Path) -> (bool, String) {
    let soffice = ["soffice", "libreoffice"]
        .into_iter()
        .find_map(|name| which::which(name).ok());
    let Some(soffice) = soffice else {
        return (false, "LibreOffice CLI not found".to_string());
    };

    let profile_dir = tempfile::tempdir().expect("profile tempdir");
    let out_dir = tempfile::tempdir().expect("out tempdir");
    let output = Command::new(soffice)
        .arg(format!(
            "-env:UserInstallation=file://{}",
            profile_dir.path().display()
        ))
        .arg("--headless")
        .arg("--convert-to")
        .arg("pdf")
        .arg("--outdir")
        .arg(out_dir.path())
        .arg(path)
        .output()
        .expect("run soffice");
    let produced = fs::read_dir(out_dir.path())
        .expect("read out dir")
        .flatten()
        .any(|entry| entry.path().extension().and_then(|s| s.to_str()) == Some("pdf"));
    let text = format!(
        "{}{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    )
    .trim()
    .to_string();
    (output.status.success() && produced, text)
}

fn main() {
    let root = repo_root();
    let data_dir = root.join("data");
    let out_dir = root.join("out");
    fs::create_dir_all(&out_dir).expect("create out dir");

    println!("Plan:");
    println!("1. Load examples/data/showcase.docx, showcase.pptx, and showcase.xlsx");
    println!(
        "2. Patch OOXML parts directly with officemd_core::patch_docx / patch_pptx / patch_xlsx_with_report"
    );
    println!("3. Showcase metadata/comment-author scopes plus ALL_TEXT semantics");
    println!("4. Verify the edited files with existing officemd extractors");
    println!("5. Open-check via LibreOffice headless PDF conversion");
    println!();

    let docx_bytes = fs::read(data_dir.join("showcase.docx")).expect("read docx");
    let edited_docx = patch_docx(
        &docx_bytes,
        &DocxPatch {
            set_core_title: Some("Edited DOCX Showcase From Rust".to_string()),
            replace_body_title: None,
            scoped_replacements: vec![
                ScopedDocxReplace {
                    scope: DocxTextScope::Headers,
                    replace: TextReplace::all(
                        "OOXML Showcase Header",
                        "OfficeMD Showcase Header — edited from Rust",
                    ),
                },
                ScopedDocxReplace {
                    scope: DocxTextScope::Body,
                    replace: TextReplace::first(
                        "Quarterly Operations Summary",
                        "Quarterly Operations Summary — edited from Rust",
                    ),
                },
                ScopedDocxReplace {
                    scope: DocxTextScope::Comments,
                    replace: TextReplace::all(
                        "Example DOCX comment captured as markdown footnote.",
                        "Edited DOCX comment from Rust patch API.",
                    )
                    .with_match_policy(MatchPolicy::Exact),
                },
                ScopedDocxReplace {
                    scope: DocxTextScope::MetadataApp,
                    replace: TextReplace::all("OfficeMD", "OfficeMD Rust Example"),
                },
                ScopedDocxReplace {
                    scope: DocxTextScope::MetadataCustom,
                    replace: TextReplace::all("showcase", "showcase-rust"),
                },
            ],
        },
    )
    .expect("patch docx");
    let docx_out = out_dir.join("showcase_edited_rust.docx");
    fs::write(&docx_out, &edited_docx).expect("write docx");
    let docx_md = officemd_docx::markdown_from_bytes(&edited_docx).expect("markdown docx");
    let docx_ir: serde_json::Value =
        serde_json::from_str(&officemd_docx::extract_ir_json(&edited_docx).expect("ir docx"))
            .expect("parse ir docx");
    let (docx_ok, docx_lo) = check_libreoffice(&docx_out);

    let pptx_bytes = fs::read(data_dir.join("showcase.pptx")).expect("read pptx");
    let edited_pptx = patch_pptx(
        &pptx_bytes,
        &PptxPatch {
            set_core_title: Some("Edited PPTX Showcase From Rust".to_string()),
            scoped_replacements: vec![
                ScopedPptxReplace {
                    scope: PptxTextScope::AllText,
                    replace: TextReplace::first(
                        "Quarterly Review",
                        "Quarterly Review — edited from Rust",
                    ),
                },
                ScopedPptxReplace {
                    scope: PptxTextScope::Comments,
                    replace: TextReplace::all(
                        "Add one slide on operating margin.",
                        "Edited PPTX comment from Rust patch API.",
                    ),
                },
                ScopedPptxReplace {
                    scope: PptxTextScope::CommentAuthors,
                    replace: TextReplace::all("Alice", "Rust Reviewer"),
                },
                ScopedPptxReplace {
                    scope: PptxTextScope::MetadataApp,
                    replace: TextReplace::all("OfficeMD", "OfficeMD Rust Example"),
                },
            ],
        },
    )
    .expect("patch pptx");
    let pptx_out = out_dir.join("showcase_edited_rust.pptx");
    fs::write(&pptx_out, &edited_pptx).expect("write pptx");
    let pptx_md = officemd_pptx::markdown_from_bytes(&edited_pptx).expect("markdown pptx");
    let pptx_ir: serde_json::Value =
        serde_json::from_str(&officemd_pptx::extract_ir_json(&edited_pptx).expect("ir pptx"))
            .expect("parse ir pptx");
    let (pptx_ok, pptx_lo) = check_libreoffice(&pptx_out);

    let xlsx_bytes = fs::read(data_dir.join("showcase.xlsx")).expect("read xlsx");
    let edited_xlsx = patch_xlsx_with_report(
        &xlsx_bytes,
        &XlsxPatch {
            set_core_title: Some("Edited XLSX Showcase From Rust".to_string()),
            rename_sheets: vec![XlsxSheetRename {
                from: "Sales".to_string(),
                to: "Sales Rust".to_string(),
                update_references: true,
            }],
            scoped_replacements: vec![
                ScopedXlsxReplace {
                    scope: XlsxTextScope::AllText,
                    replace: TextReplace::all("North", "North Rust"),
                },
                ScopedXlsxReplace {
                    scope: XlsxTextScope::Comments,
                    replace: TextReplace::all("Review", "Reviewed from Rust"),
                },
                ScopedXlsxReplace {
                    scope: XlsxTextScope::MetadataApp,
                    replace: TextReplace::all("OfficeMD", "OfficeMD Rust Example"),
                },
                ScopedXlsxReplace {
                    scope: XlsxTextScope::MetadataCustom,
                    replace: TextReplace::all("showcase", "showcase-rust"),
                },
            ],
        },
    )
    .expect("patch xlsx");
    let xlsx_out = out_dir.join("showcase_edited_rust.xlsx");
    fs::write(&xlsx_out, &edited_xlsx.content).expect("write xlsx");
    let xlsx_md = officemd_xlsx::markdown_from_bytes(&edited_xlsx.content).expect("markdown xlsx");
    let xlsx_ir: serde_json::Value = serde_json::from_str(
        &officemd_xlsx::extract_ir::extract_ir_json(&edited_xlsx.content).expect("ir xlsx"),
    )
    .expect("parse ir xlsx");
    let (xlsx_ok, xlsx_lo) = check_libreoffice(&xlsx_out);

    println!("DOCX result: {}", docx_out.display());
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "core_title": docx_ir["properties"]["core"]["title"],
            "body_title": docx_ir["sections"][0]["blocks"][0]["Paragraph"]["inlines"][0]["Text"],
            "header_text": docx_ir["sections"][1]["blocks"][0]["Paragraph"]["inlines"][0]["Text"],
            "markdown_has_comment": docx_md.contains("Edited DOCX comment from Rust patch API."),
            "libreoffice_ok": docx_ok,
            "libreoffice_output": docx_lo,
        }))
        .unwrap()
    );
    println!();

    println!("PPTX result: {}", pptx_out.display());
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "core_title": pptx_ir["properties"]["core"]["title"],
            "slide_1_title": pptx_ir["slides"][0]["title"],
            "slide_1_comment_author": pptx_ir["slides"][0]["comments"][0]["author"],
            "markdown_has_comment": pptx_md.contains("Edited PPTX comment from Rust patch API."),
            "libreoffice_ok": pptx_ok,
            "libreoffice_output": pptx_lo,
        }))
        .unwrap()
    );
    println!();

    println!("XLSX result: {}", xlsx_out.display());
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "core_title": xlsx_ir["properties"]["core"]["title"],
            "sheet_1_name": xlsx_ir["sheets"][0]["name"],
            "markdown_has_edit": xlsx_md.contains("North Rust") || xlsx_md.contains("Reviewed from Rust"),
            "report": {
                "parts_scanned": edited_xlsx.report.parts_scanned,
                "parts_modified": edited_xlsx.report.parts_modified,
                "replacements_applied": edited_xlsx.report.replacements_applied,
            },
            "libreoffice_ok": xlsx_ok,
            "libreoffice_output": xlsx_lo,
        }))
        .unwrap()
    );
}
