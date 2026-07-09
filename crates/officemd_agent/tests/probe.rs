use std::path::PathBuf;

use officemd_agent::{
    AgentDocumentFormat, AgentService, ApplyPatchRequest, ArtifactPatchPlan, InspectQuery,
    InspectRequest, InspectionPayload, MutationKind, PatchOperation, PatchPlanVersion,
    VerificationCheckKind, VerificationStatus, VerifyRequest, XlsxRangeInclude,
    artifact::fingerprint_bytes,
    locator::{ArtifactLocator, DocxPartLocator},
};
use officemd_core::test_helpers::build_zip;

fn temp_file(name: &str, contents: &[u8]) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!("officemd-agent-test-{}-{name}", std::process::id()));
    std::fs::write(&path, contents).expect("write fixture");
    path
}

#[test]
fn probe_csv_reports_identity_and_capabilities() {
    let path = temp_file("probe.csv", b"name,count\napples,3\n");
    let service = AgentService::new();
    let report = service
        .inspect(&InspectRequest {
            input: path.clone(),
            format: Some(AgentDocumentFormat::Csv),
            query: InspectQuery::DocumentSummary,
        })
        .expect("probe csv");

    assert_eq!(report.artifact.format, AgentDocumentFormat::Csv);
    assert_eq!(report.artifact.fingerprint.byte_length, 20);
    assert!(report.capability.readable);
    assert!(report.capability.mutable_operations.is_empty());
    assert_eq!(report.findings.len(), 1);
    assert!(matches!(
        report.findings[0].payload,
        InspectionPayload::Summary {
            format: AgentDocumentFormat::Csv,
            sheets: 1,
            ..
        }
    ));

    std::fs::remove_file(path).ok();
}

#[test]
fn xlsx_sheet_query_rejects_csv_artifact() {
    let path = temp_file("wrong-query.csv", b"name,count\napples,3\n");
    let service = AgentService::new();
    let err = service
        .inspect(&InspectRequest {
            input: path.clone(),
            format: Some(AgentDocumentFormat::Csv),
            query: InspectQuery::XlsxSheets,
        })
        .expect_err("query should reject wrong format");

    assert!(err.to_string().contains("query requires xlsx"));
    std::fs::remove_file(path).ok();
}

#[test]
fn verify_csv_structure_passes_and_visual_is_warning() {
    let path = temp_file("verify.csv", b"name,count\napples,3\n");
    let service = AgentService::new();
    let report = service
        .verify(&VerifyRequest {
            input: path.clone(),
            checks: vec![
                VerificationCheckKind::Structure,
                VerificationCheckKind::VisualRender,
            ],
            render_output_dir: None,
        })
        .expect("verify csv");

    assert_eq!(report.status, VerificationStatus::PassedWithWarnings);
    assert_eq!(report.checks.len(), 2);
    assert_eq!(report.checks[0].status, VerificationStatus::Passed);
    assert_eq!(report.checks[1].status, VerificationStatus::NotRunnable);

    std::fs::remove_file(path).ok();
}

#[test]
fn apply_rejects_unsupported_csv_without_output() {
    let input = temp_file("apply.csv", b"name,count\napples,3\n");
    let mut output = input.clone();
    output.set_file_name("officemd-agent-test-apply-out.csv");
    std::fs::remove_file(&output).ok();
    let bytes = std::fs::read(&input).expect("read source");
    let err = AgentService::new()
        .apply_patch(&ApplyPatchRequest {
            input: input.clone(),
            output: output.clone(),
            expected_source_sha256: fingerprint_bytes(&bytes).sha256,
            patch: ArtifactPatchPlan {
                version: PatchPlanVersion::V1,
                request_id: "test-request".to_string(),
                operations: Vec::new(),
            },
        })
        .expect_err("csv mutation should be unsupported");

    assert!(err.to_string().contains("csv mutation is not supported"));
    assert!(!output.exists());

    std::fs::remove_file(input).ok();
}

#[test]
fn xlsx_capability_reports_supported_apply_operations() {
    let path = temp_file("capability.xlsx", &sparse_xlsx());
    let service = AgentService::new();
    let report = service
        .inspect(&InspectRequest {
            input: path.clone(),
            format: Some(AgentDocumentFormat::Xlsx),
            query: InspectQuery::DocumentSummary,
        })
        .expect("inspect xlsx");

    assert_eq!(
        report.capability.mutable_operations,
        vec![MutationKind::ReplaceText, MutationKind::RenameXlsxSheet]
    );

    std::fs::remove_file(path).ok();
}

#[test]
fn xlsx_range_uses_absolute_sparse_cell_addresses() {
    let path = temp_file("sparse.xlsx", &sparse_xlsx());
    let service = AgentService::new();
    let report = service
        .inspect(&InspectRequest {
            input: path.clone(),
            format: Some(AgentDocumentFormat::Xlsx),
            query: InspectQuery::XlsxRange {
                sheet: "Data".to_string(),
                range: "C5:C5".to_string(),
                include: XlsxRangeInclude {
                    values: true,
                    formulas: true,
                    number_formats: false,
                },
            },
        })
        .expect("inspect sparse cell");

    assert_eq!(report.findings.len(), 1);
    assert!(matches!(
        &report.findings[0].payload,
        InspectionPayload::Cell {
            address,
            value: Some(value),
            formula: Some(formula),
            ..
        } if address == "C5" && value == "42" && formula == "SUM(A1:A1)"
    ));

    std::fs::remove_file(path).ok();
}

#[test]
fn docx_apply_replaces_only_targeted_paragraph() {
    let input = temp_file("targeted.docx", &docx_with_duplicate_text());
    let mut output = input.clone();
    output.set_file_name("officemd-agent-test-targeted-out.docx");
    std::fs::remove_file(&output).ok();
    let bytes = std::fs::read(&input).expect("read source");

    let report = AgentService::new()
        .apply_patch(&ApplyPatchRequest {
            input: input.clone(),
            output: output.clone(),
            expected_source_sha256: fingerprint_bytes(&bytes).sha256,
            patch: ArtifactPatchPlan {
                version: PatchPlanVersion::V1,
                request_id: "targeted-docx".to_string(),
                operations: vec![PatchOperation::ReplaceText {
                    target: ArtifactLocator::DocxParagraph {
                        part: DocxPartLocator {
                            part: "body".to_string(),
                        },
                        paragraph_index: 1,
                    },
                    expected_text: "Status".to_string(),
                    replacement: "Done".to_string(),
                    preserve_formatting: false,
                }],
            },
        })
        .expect("apply targeted docx patch");

    assert!(report.output.is_some());
    let patched = std::fs::read(&output).expect("read output");
    let inspected = AgentService::new()
        .inspect(&InspectRequest {
            input: output.clone(),
            format: Some(AgentDocumentFormat::Docx),
            query: InspectQuery::DocxOutline,
        })
        .expect("inspect patched docx");
    let texts = inspected
        .findings
        .into_iter()
        .map(|finding| match finding.payload {
            InspectionPayload::Text { text } => text,
            _ => String::new(),
        })
        .collect::<Vec<_>>();

    assert_eq!(texts, vec!["Status".to_string(), "Done".to_string()]);
    assert_ne!(
        fingerprint_bytes(&bytes).sha256,
        fingerprint_bytes(&patched).sha256
    );

    std::fs::remove_file(input).ok();
    std::fs::remove_file(output).ok();
}

fn sparse_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Data" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="5">
            <c r="C5"><f>SUM(A1:A1)</f><v>42</v></c>
        </row>
    </sheetData>
</worksheet>"#;
    build_zip(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet),
    ])
}

fn docx_with_duplicate_text() -> Vec<u8> {
    let document = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Status</w:t></w:r></w:p>
    <w:p><w:r><w:t>Status</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
    build_zip(vec![("word/document.xml", document)])
}
