use officemd_core::ir::DocumentKind;
use officemd_csv::extract_ir::extract_sheet_names;
use officemd_csv::render::markdown_from_bytes;
use officemd_csv::table_ir::{
    CsvExtractOptions, extract_tables_ir, extract_tables_ir_json_with_options,
    extract_tables_ir_with_options,
};

#[test]
fn extracts_single_sheet_name() {
    let names = extract_sheet_names(b"a,b\n1,2\n").expect("sheet names");
    assert_eq!(names, vec!["Sheet1".to_string()]);
}

#[test]
fn extracts_tables_ir_from_comma_csv() {
    let content = b"name,value\nwidget,42\n";
    let doc = extract_tables_ir(content).expect("extract tables ir");

    assert_eq!(doc.kind, DocumentKind::Xlsx);
    assert_eq!(doc.sheets.len(), 1);
    assert_eq!(doc.sheets[0].name, "Sheet1");
    assert_eq!(doc.sheets[0].tables.len(), 1);
    assert_eq!(doc.sheets[0].tables[0].headers, vec!["Col1", "Col2"]);
}

#[test]
fn delimiter_option_is_applied() {
    let content = b"name;value\nwidget;42\n";
    let doc = extract_tables_ir_with_options(
        content,
        CsvExtractOptions {
            delimiter: b';',
            ..Default::default()
        },
    )
    .expect("extract tables ir");

    let first_row = &doc.sheets[0].tables[0].rows[0];
    let a = &first_row[0].content[0].inlines[0];
    let b = &first_row[1].content[0].inlines[0];
    assert_eq!(format!("{a:?}"), "Text(\"name\")");
    assert_eq!(format!("{b:?}"), "Text(\"value\")");
}

#[test]
fn formulas_are_collected_from_equals_cells() {
    let content = b"item,formula\nwidget,\"=SUM(1,2,3)\"\n";
    let doc = extract_tables_ir(content).expect("extract tables ir");

    assert_eq!(doc.sheets[0].formulas.len(), 1);
    assert_eq!(doc.sheets[0].formulas[0].cell_ref, "B2");
    assert_eq!(doc.sheets[0].formulas[0].formula, "SUM(1,2,3)");
}

#[test]
fn markdown_renders_sheet_table_and_formulas() {
    let content = b"item,formula\nwidget,\"=SUM(1,2,3)\"\n";
    let markdown = markdown_from_bytes(content).expect("markdown");

    assert!(markdown.contains("## Sheet: Sheet1"));
    assert!(markdown.contains("| item | formula |"));
    assert!(markdown.contains("B2=`=SUM(1,2,3)`"));
}

#[test]
fn extract_tables_ir_json_with_options_omits_properties_by_default() {
    let content = b"name,value\nwidget,42\n";
    let payload = extract_tables_ir_json_with_options(content, b',', false).expect("extract json");
    let value: serde_json::Value = serde_json::from_str(&payload).expect("valid json");

    assert!(value["properties"].is_null());
}

#[test]
fn extract_tables_ir_json_with_options_includes_properties_when_requested() {
    let content = b"name,value\nwidget,42\n";
    let payload = extract_tables_ir_json_with_options(content, b',', true).expect("extract json");
    let value: serde_json::Value = serde_json::from_str(&payload).expect("valid json");

    assert!(value["properties"].is_object());
    assert_eq!(value["properties"]["custom"]["source_format"], "csv");
}
