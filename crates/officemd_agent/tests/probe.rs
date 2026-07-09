use std::path::PathBuf;

use officemd_agent::{
    AgentDocumentFormat, AgentService, ApplyPatchRequest, ArtifactPatchPlan, InspectQuery,
    InspectRequest, InspectionPayload, PatchPlanVersion, VerificationCheckKind, VerificationStatus,
    VerifyRequest, artifact::fingerprint_bytes,
};

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
