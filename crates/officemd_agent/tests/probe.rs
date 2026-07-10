use std::path::PathBuf;

use officemd_agent::{
    AgentDocumentFormat, AgentService, ApplyPatchRequest, ApplyPatchStatus, ArtifactPatchPlan,
    ArtifactRef, ArtifactRenderer, CellValue, DiffArtifactRequest, InspectQuery, InspectRequest,
    InspectionPayload, MutationKind, PatchOperation, PatchPlanVersion, RenderBackendKind,
    RenderCapability, RenderReport, RenderRequest, VerificationCheckKind, VerificationStatus,
    VerifyRequest, XlsxRangeInclude,
    artifact::fingerprint_bytes,
    locator::{ArtifactLocator, DocxPartLocator, XlsxCellLocator, XlsxSheetLocator},
};
use officemd_core::test_helpers::build_zip;

fn temp_file(name: &str, contents: &[u8]) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!("officemd-agent-test-{}-{name}", std::process::id()));
    std::fs::write(&path, contents).expect("write fixture");
    path
}

fn remove_artifact(path: &PathBuf) {
    std::fs::remove_file(path).ok();
    let mut sidecar = path.as_os_str().to_os_string();
    sidecar.push(".officemd.json");
    std::fs::remove_file(PathBuf::from(sidecar)).ok();
}

#[derive(Debug, Clone, Copy)]
struct FakeRenderer;

impl ArtifactRenderer for FakeRenderer {
    fn capability(&self) -> RenderCapability {
        RenderCapability::Available {
            backend: RenderBackendKind::Poppler,
            formats: vec![AgentDocumentFormat::Csv],
        }
    }

    fn render(&self, request: &RenderRequest) -> officemd_agent::AgentResult<RenderReport> {
        let bytes = std::fs::read(&request.input).expect("read fake render input");
        Ok(RenderReport {
            schema_version: officemd_agent::AGENT_SCHEMA_VERSION,
            artifact: ArtifactRef {
                path: request.input.clone(),
                format: AgentDocumentFormat::Csv,
                fingerprint: fingerprint_bytes(&bytes),
            },
            backend: RenderBackendKind::Poppler,
            images: Vec::new(),
            diagnostics: Vec::new(),
        })
    }
}

#[test]
fn probe_csv_reports_identity_and_capabilities() {
    let path = temp_file("probe.csv", b"name,count\napples,3\n");
    let service = AgentService::new();
    let report = service
        .probe_path_as(&path, Some(AgentDocumentFormat::Csv))
        .expect("probe csv");

    assert_eq!(report.artifact.format, AgentDocumentFormat::Csv);
    assert_eq!(report.artifact.fingerprint.byte_length, 20);
    assert!(report.capability.readable);
    assert!(report.capability.mutable_operations.is_empty());
    assert!(report.findings.is_empty());

    std::fs::remove_file(path).ok();
}

#[test]
fn probe_rejects_malformed_ooxml_package() {
    let path = temp_file("malformed.docx", b"not a zip package");
    let error = AgentService::new()
        .probe_path(&path)
        .expect_err("malformed OOXML should fail probe");

    assert!(matches!(error, officemd_agent::AgentError::Extraction(_)));
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
fn verification_uses_injected_renderer_through_agent_service() {
    let path = temp_file("verify-render.csv", b"name,count\napples,3\n");
    let mut render_dir = path.clone();
    render_dir.set_extension("rendered");
    let report = AgentService::with_renderer(FakeRenderer)
        .verify(&VerifyRequest {
            input: path.clone(),
            checks: vec![VerificationCheckKind::VisualRender],
            render_output_dir: Some(render_dir),
        })
        .expect("verify through renderer seam");

    assert_eq!(report.status, VerificationStatus::Passed);
    assert!(report.rendered_evidence.is_some());
    std::fs::remove_file(path).ok();
}

#[test]
fn artifact_diff_compares_canonical_semantics() {
    let left = temp_file("diff-left.csv", b"name,count\napples,3\n");
    let right = temp_file("diff-right.csv", b"name,count\napples,4\n");
    let report = AgentService::new()
        .diff_artifacts(&DiffArtifactRequest {
            left: left.clone(),
            right: right.clone(),
            semantic: true,
            rendered: false,
            render_output_dir: None,
        })
        .expect("diff CSV semantics");

    assert!(!report.equal);
    assert!(report.semantic.is_some_and(|semantic| !semantic.equal));
    assert!(report.visual.is_none());
    std::fs::remove_file(left).ok();
    std::fs::remove_file(right).ok();
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
                operations: vec![PatchOperation::ReplaceText {
                    target: ArtifactLocator::XlsxCell {
                        sheet: "Sheet1".to_string(),
                        address: "A1".to_string(),
                    },
                    expected_text: "apples".to_string(),
                    replacement: "oranges".to_string(),
                    preserve_formatting: false,
                }],
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
        vec![
            MutationKind::SetXlsxCellValue,
            MutationKind::SetXlsxCellFormula,
            MutationKind::RenameXlsxSheet,
        ]
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
                    number_formats: true,
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
            number_format: Some(number_format),
        } if address == "C5" && value == "42" && formula == "SUM(A1:A1)" && number_format == "0.00"
    ));

    std::fs::remove_file(path).ok();
}

#[test]
fn xlsx_verification_detects_stored_formula_error() {
    let input = temp_file("formula-error.xlsx", &formula_error_xlsx());
    let report = AgentService::new()
        .verify(&VerifyRequest {
            input: input.clone(),
            checks: vec![VerificationCheckKind::FormulaReferences],
            render_output_dir: None,
        })
        .expect("verify stored formula error");

    assert_eq!(report.status, VerificationStatus::Failed);
    assert!(matches!(
        report.checks[0].locator,
        Some(ArtifactLocator::XlsxCell { ref sheet, ref address })
            if sheet == "Data" && address == "B2"
    ));
    std::fs::remove_file(input).ok();
}

#[test]
fn docx_apply_replaces_only_targeted_paragraph() {
    let input = temp_file("targeted.docx", &docx_with_duplicate_text());
    let mut output = input.clone();
    output.set_file_name("officemd-agent-test-targeted-out.docx");
    remove_artifact(&output);
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
                    replacement: "Done & Verified".to_string(),
                    preserve_formatting: true,
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

    assert_eq!(
        texts,
        vec!["Status".to_string(), "Done & Verified".to_string()]
    );
    assert_ne!(
        fingerprint_bytes(&bytes).sha256,
        fingerprint_bytes(&patched).sha256
    );

    std::fs::remove_file(input).ok();
    remove_artifact(&output);
}

#[test]
fn rejected_patch_does_not_report_uncommitted_operations_as_applied() {
    let input = temp_file("rejected.docx", &docx_with_duplicate_text());
    let mut output = input.clone();
    output.set_file_name("officemd-agent-test-rejected-out.docx");
    remove_artifact(&output);
    let bytes = std::fs::read(&input).expect("read source");
    let report = AgentService::new()
        .apply_patch(&ApplyPatchRequest {
            input: input.clone(),
            output: output.clone(),
            expected_source_sha256: fingerprint_bytes(&bytes).sha256,
            patch: ArtifactPatchPlan {
                version: PatchPlanVersion::V1,
                request_id: "rejected-plan".to_string(),
                operations: vec![
                    PatchOperation::ReplaceText {
                        target: ArtifactLocator::DocxParagraph {
                            part: DocxPartLocator {
                                part: "body".to_string(),
                            },
                            paragraph_index: 0,
                        },
                        expected_text: "Status".to_string(),
                        replacement: "Changed".to_string(),
                        preserve_formatting: false,
                    },
                    PatchOperation::SetXlsxCellValue {
                        target: XlsxCellLocator {
                            sheet: "Data".to_string(),
                            address: "A1".to_string(),
                        },
                        value: CellValue::Text("unsupported".to_string()),
                    },
                ],
            },
        })
        .expect("return rejected report");

    assert_eq!(report.status, ApplyPatchStatus::Rejected);
    assert_eq!(
        report.operations[0].status,
        officemd_agent::OperationStatus::Skipped
    );
    assert_eq!(
        report.operations[1].status,
        officemd_agent::OperationStatus::Unsupported
    );
    assert!(!output.exists());
    std::fs::remove_file(input).ok();
    remove_artifact(&output);
}

#[test]
fn xlsx_apply_sets_exact_cells_and_retries_idempotently() {
    let input = temp_file("set-cells.xlsx", &sparse_xlsx());
    let mut output = input.clone();
    output.set_file_name("officemd-agent-test-set-cells-out.xlsx");
    remove_artifact(&output);
    let bytes = std::fs::read(&input).expect("read source");
    let request = ApplyPatchRequest {
        input: input.clone(),
        output: output.clone(),
        expected_source_sha256: fingerprint_bytes(&bytes).sha256,
        patch: ArtifactPatchPlan {
            version: PatchPlanVersion::V1,
            request_id: "set-cells".to_string(),
            operations: vec![
                PatchOperation::SetXlsxCellValue {
                    target: XlsxCellLocator {
                        sheet: "Data".to_string(),
                        address: "D5".to_string(),
                    },
                    value: CellValue::Text("ready & reviewed".to_string()),
                },
                PatchOperation::SetXlsxCellFormula {
                    target: XlsxCellLocator {
                        sheet: "Data".to_string(),
                        address: "E5".to_string(),
                    },
                    formula: "=C5*2".to_string(),
                },
            ],
        },
    };

    let first = AgentService::new()
        .apply_patch(&request)
        .expect("apply exact XLSX cells");
    assert_eq!(first.status, ApplyPatchStatus::Applied);
    let retry = AgentService::new()
        .apply_patch(&request)
        .expect("retry exact XLSX cells");
    assert_eq!(retry.status, ApplyPatchStatus::AlreadyApplied);

    let inspected = AgentService::new()
        .inspect(&InspectRequest {
            input: output.clone(),
            format: Some(AgentDocumentFormat::Xlsx),
            query: InspectQuery::XlsxRange {
                sheet: "Data".to_string(),
                range: "D5:E5".to_string(),
                include: XlsxRangeInclude {
                    values: true,
                    formulas: true,
                    number_formats: true,
                },
            },
        })
        .expect("inspect exact XLSX cells");
    assert!(matches!(
        &inspected.findings[0].payload,
        InspectionPayload::Cell { value: Some(value), .. } if value == "ready & reviewed"
    ));
    assert!(matches!(
        &inspected.findings[1].payload,
        InspectionPayload::Cell { formula: Some(formula), .. } if formula == "C5*2"
    ));

    std::fs::remove_file(input).ok();
    remove_artifact(&output);
}

#[test]
fn xlsx_sheet_rename_rejects_duplicate_name() {
    let input = temp_file("duplicate-sheet.xlsx", &two_sheet_xlsx());
    let mut output = input.clone();
    output.set_file_name("officemd-agent-test-duplicate-sheet-out.xlsx");
    remove_artifact(&output);
    let bytes = std::fs::read(&input).expect("read source");
    let error = AgentService::new()
        .apply_patch(&ApplyPatchRequest {
            input: input.clone(),
            output: output.clone(),
            expected_source_sha256: fingerprint_bytes(&bytes).sha256,
            patch: ArtifactPatchPlan {
                version: PatchPlanVersion::V1,
                request_id: "duplicate-sheet".to_string(),
                operations: vec![PatchOperation::RenameXlsxSheet {
                    target: XlsxSheetLocator {
                        name: "Data".to_string(),
                    },
                    new_name: "Summary".to_string(),
                }],
            },
        })
        .expect_err("duplicate sheet should fail");
    assert!(error.to_string().contains("already exists"));
    assert!(!output.exists());

    std::fs::remove_file(input).ok();
    remove_artifact(&output);
}

#[test]
fn pptx_inspection_and_apply_use_stable_shape_id() {
    let input = temp_file("shape.pptx", &single_shape_pptx());
    let mut output = input.clone();
    output.set_file_name("officemd-agent-test-shape-out.pptx");
    remove_artifact(&output);
    let inspection = AgentService::new()
        .inspect(&InspectRequest {
            input: input.clone(),
            format: Some(AgentDocumentFormat::Pptx),
            query: InspectQuery::PptxSlides { start: 1, end: 1 },
        })
        .expect("inspect PPTX shapes");
    assert!(inspection.findings.iter().any(|finding| matches!(
        &finding.payload,
        InspectionPayload::Shape {
            slide_number: 1,
            shape_id: 42,
            text,
        } if text == "Original"
    )));

    let bytes = std::fs::read(&input).expect("read source");
    AgentService::new()
        .apply_patch(&ApplyPatchRequest {
            input: input.clone(),
            output: output.clone(),
            expected_source_sha256: fingerprint_bytes(&bytes).sha256,
            patch: ArtifactPatchPlan {
                version: PatchPlanVersion::V1,
                request_id: "shape-replace".to_string(),
                operations: vec![PatchOperation::ReplacePptxShapeText {
                    target: ArtifactLocator::PptxShape {
                        slide_number: 1,
                        shape_id: 42,
                    },
                    expected_text: "Original".to_string(),
                    replacement: "Updated & safe".to_string(),
                }],
            },
        })
        .expect("replace exact PPTX shape");
    let updated = AgentService::new()
        .inspect(&InspectRequest {
            input: output.clone(),
            format: Some(AgentDocumentFormat::Pptx),
            query: InspectQuery::PptxSlides { start: 1, end: 1 },
        })
        .expect("inspect updated PPTX shape");
    assert!(updated.findings.iter().any(|finding| matches!(
        &finding.payload,
        InspectionPayload::Shape { shape_id: 42, text, .. } if text == "Updated & safe"
    )));

    std::fs::remove_file(input).ok();
    remove_artifact(&output);
}

#[test]
fn pptx_canvas_overflow_reports_shape_locator() {
    let input = temp_file("overflow.pptx", &single_shape_pptx());
    let report = AgentService::new()
        .verify(&VerifyRequest {
            input: input.clone(),
            checks: vec![VerificationCheckKind::PptxCanvasOverflow],
            render_output_dir: None,
        })
        .expect("verify PPTX canvas bounds");

    assert_eq!(report.status, VerificationStatus::Failed);
    assert!(matches!(
        report.checks[0].locator,
        Some(ArtifactLocator::PptxShape {
            slide_number: 1,
            shape_id: 42,
        })
    ));
    std::fs::remove_file(input).ok();
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
            <c r="C5" s="1"><f>SUM(A1:A1)</f><v>42</v></c>
        </row>
    </sheetData>
</worksheet>"#;
    let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>
  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs>
</styleSheet>"#;
    build_zip(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet),
        ("xl/styles.xml", styles),
    ])
}

fn two_sheet_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Data" sheetId="1" r:id="rId1"/>
        <sheet name="Summary" sheetId="2" r:id="rId2"/>
    </sheets>
</workbook>"#;
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#;
    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
    build_zip(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet),
        ("xl/worksheets/sheet2.xml", sheet),
    ])
}

fn formula_error_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
    let relationships = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="2"><c r="B2" t="e"><f>Missing!A1</f><v>#REF!</v></c></row></sheetData>
</worksheet>"#;
    build_zip(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", relationships),
        ("xl/worksheets/sheet1.xml", sheet),
    ])
}

fn docx_with_duplicate_text() -> Vec<u8> {
    let document = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Status</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><w:t>Sta</w:t></w:r><w:r><w:t>tus</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
    build_zip(vec![("word/document.xml", document)])
}

fn single_shape_pptx() -> Vec<u8> {
    let presentation = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="1000" cy="1000"/>
</p:presentation>"#;
    let relationships = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;
    let slide = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="42" name="Text 42"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="900" y="100"/><a:ext cx="200" cy="200"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
    build_zip(vec![
        ("ppt/presentation.xml", presentation),
        ("ppt/_rels/presentation.xml.rels", relationships),
        ("ppt/slides/slide1.xml", slide),
    ])
}
