use std::path::PathBuf;

use officemd_agent::{
    AgentDocumentFormat, AgentService, InspectQuery, InspectRequest, InspectionPayload,
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
