use std::fs;
use std::io::Write;

use officemd_core::ir::{DocumentKind, Inline};
use officemd_xlsx::extract_ir::{extract_ir, extract_sheet_names};
use officemd_xlsx::table_ir::{
    SheetFilter, XlsxExtractOptions, extract_tables_ir, extract_tables_ir_json_with_options,
    extract_tables_ir_with_options,
};
use zip::ZipWriter;
use zip::write::FileOptions;

fn build_xlsx(parts: Vec<(&str, &str)>) -> Vec<u8> {
    let mut buffer = Vec::new();
    let mut writer = ZipWriter::new(std::io::Cursor::new(&mut buffer));
    let options: FileOptions<'_, ()> = FileOptions::default();

    for (path, contents) in parts {
        writer.start_file(path, options).unwrap();
        writer.write_all(contents.as_bytes()).unwrap();
    }

    writer.finish().unwrap();
    buffer
}

fn minimal_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
        <sheet name="Data" sheetId="2" r:id="rId2"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#;

    let sheet1 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
            <c r="B1" t="inlineStr"><is><t>World</t></is></c>
        </row>
    </sheetData>
</worksheet>"#;

    let sheet2 = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
            <c r="B1" t="inlineStr"><is><t>Value</t></is></c>
        </row>
        <row r="2">
            <c r="A2" t="inlineStr"><is><t>Item</t></is></c>
            <c r="B2"><f>SUM(C1:C10)</f><v>100</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1),
        ("xl/worksheets/sheet2.xml", sheet2),
    ])
}

fn styled_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Styled" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" applyNumberFormat="0"/>
    <xf numFmtId="14" applyNumberFormat="1"/>
    <xf numFmtId="10" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

    let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Hello Shared</t></si>
</sst>"#;

    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" s="1"><v>45292</v></c>
            <c r="B1" s="2"><v>0.125</v></c>
            <c r="C1" t="s"><v>0</v></c>
            <c r="D1" t="inlineStr"><is><t>Hello Inline</t></is></c>
        </row>
    </sheetData>
</worksheet>"#;

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/styles.xml", styles),
        ("xl/sharedStrings.xml", shared_strings),
        ("xl/worksheets/sheet1.xml", sheet),
    ])
}

fn styled_xlsx_1904() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr date1904="1"/>
    <sheets>
        <sheet name="Styled1904" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" applyNumberFormat="0"/>
    <xf numFmtId="14" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" s="1"><v>43830</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/styles.xml", styles),
        ("xl/worksheets/sheet1.xml", sheet),
    ])
}

fn col_to_name(mut n: usize) -> String {
    let mut name = String::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        let offset = u8::try_from(rem).expect("column remainder is less than 26");
        name.insert(0, char::from(b'A' + offset));
        n = (n - 1) / 26;
    }
    name
}

fn large_sheet_xlsx(row_count: usize, col_count: usize) -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Large" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let mut sheet = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
"#,
    );
    for row in 1..=row_count {
        std::fmt::Write::write_fmt(&mut sheet, format_args!("<row r=\"{row}\">"))
            .expect("write row");
        for col in 1..=col_count {
            let col_name = col_to_name(col);
            let cell_ref = format!("{col_name}{row}");
            let value = row * 1000 + col;
            std::fmt::Write::write_fmt(
                &mut sheet,
                format_args!("<c r=\"{cell_ref}\"><v>{value}</v></c>"),
            )
            .expect("write cell");
        }
        sheet.push_str("</row>");
    }
    sheet.push_str("</sheetData></worksheet>");

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", &sheet),
    ])
}

fn sparse_high_row_xlsx(row_1_based: usize) -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sparse" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="{row_1_based}">
            <c r="A{row_1_based}"><v>42</v></c>
        </row>
    </sheetData>
</worksheet>"#
    );

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", &sheet),
    ])
}

fn minimal_xlsx_with_doc_props() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1"><c r="A1"><v>1</v></c></row>
    </sheetData>
</worksheet>"#;

    let core = r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Quarterly Results</dc:title>
</cp:coreProperties>"#;

    let app = r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Excel</Application>
</Properties>"#;

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet),
        ("docProps/core.xml", core),
        ("docProps/app.xml", app),
    ])
}

fn first_inline_text(doc: &officemd_core::ir::OoxmlDocument, row: usize, col: usize) -> String {
    let table = &doc.sheets[0].tables[0];
    let cell = &table.rows[row][col];
    for para in &cell.content {
        for inline in &para.inlines {
            if let Inline::Text(text) = inline {
                return text.clone();
            }
        }
    }
    String::new()
}

#[test]
fn extracts_sheet_names() {
    let content = minimal_xlsx();
    let names = extract_sheet_names(&content).expect("extract sheet names");

    assert_eq!(names.len(), 2);
    assert_eq!(names[0], "Sheet1");
    assert_eq!(names[1], "Data");
}

#[test]
fn extracts_sheet_names_with_entities() {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="G&amp;A" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let empty_sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData/>
</worksheet>"#;

    let content = build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", empty_sheet),
    ]);

    let names = extract_sheet_names(&content).expect("extract sheet names");
    assert_eq!(names.len(), 1);
    assert_eq!(names[0], "G&A");
}

#[test]
fn extracts_basic_ir() {
    let content = minimal_xlsx();
    let doc = extract_ir(&content).expect("extract IR");

    assert_eq!(doc.kind, DocumentKind::Xlsx);
    assert_eq!(doc.sheets.len(), 2);
    assert_eq!(doc.sheets[0].name, "Sheet1");
    assert_eq!(doc.sheets[1].name, "Data");
}

#[test]
fn extracts_ir_from_fixture() {
    let fixture_path = "tests/fixtures/sample.xlsx";
    if !std::path::Path::new(fixture_path).exists() {
        eprintln!("Skipping fixture test: {fixture_path} not found");
        return;
    }

    let content = fs::read(fixture_path).expect("read fixture");
    let names = extract_sheet_names(&content).expect("extract sheet names");
    assert!(!names.is_empty(), "fixture should have at least one sheet");

    let doc = extract_ir(&content).expect("extract IR");
    assert_eq!(doc.kind, DocumentKind::Xlsx);
    assert_eq!(doc.sheets.len(), names.len());
}

#[test]
fn extracts_tables_ir_from_fixture() {
    let fixture_path = "tests/fixtures/sample.xlsx";
    if !std::path::Path::new(fixture_path).exists() {
        eprintln!("Skipping fixture test: {fixture_path} not found");
        return;
    }

    let content = fs::read(fixture_path).expect("read fixture");
    let doc = extract_tables_ir(&content).expect("extract tables IR");

    assert_eq!(doc.kind, DocumentKind::Xlsx);
    assert!(!doc.sheets.is_empty(), "should have at least one sheet");

    // Each sheet should have at least one table
    for sheet in &doc.sheets {
        assert!(
            !sheet.tables.is_empty(),
            "sheet {} should have at least one table",
            sheet.name
        );
        for table in &sheet.tables {
            assert!(!table.headers.is_empty(), "table should have headers");
            assert!(!table.rows.is_empty(), "table should have at least one row");
        }
    }
}

#[test]
fn table_ir_contains_cell_text() {
    let fixture_path = "tests/fixtures/sample.xlsx";
    if !std::path::Path::new(fixture_path).exists() {
        eprintln!("Skipping fixture test: {fixture_path} not found");
        return;
    }

    let content = fs::read(fixture_path).expect("read fixture");
    let doc = extract_tables_ir(&content).expect("extract tables IR");

    // Collect all text from the document
    let mut all_text = String::new();
    for sheet in &doc.sheets {
        all_text.push_str(&sheet.name);
        for table in &sheet.tables {
            for row in &table.rows {
                for cell in row {
                    for para in &cell.content {
                        for inline in &para.inlines {
                            if let Inline::Text(t) = inline {
                                all_text.push_str(t);
                            }
                        }
                    }
                }
            }
        }
    }

    // The fixture should have some content
    assert!(!all_text.is_empty(), "document should contain text");
}

#[test]
fn empty_sheet_gets_synthetic_table() {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Empty" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let empty_sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData/>
</worksheet>"#;

    let content = build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", empty_sheet),
    ]);

    let doc = extract_tables_ir(&content).expect("extract tables IR");
    assert_eq!(doc.sheets.len(), 1);
    assert_eq!(doc.sheets[0].name, "Empty");

    // Empty sheets should still have at least one table with one row
    assert_eq!(doc.sheets[0].tables.len(), 1);
    assert!(!doc.sheets[0].tables[0].rows.is_empty());
}

#[test]
fn options_default_matches_current_behavior() {
    let content = minimal_xlsx();
    let baseline = extract_tables_ir(&content).expect("baseline extract");
    let with_options = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())
        .expect("extract with options");

    let baseline_json = serde_json::to_string(&baseline).expect("serialize baseline");
    let options_json = serde_json::to_string(&with_options).expect("serialize options");
    assert_eq!(baseline_json, options_json);
}

#[test]
fn sheet_filter_is_applied_during_extraction() {
    let content = minimal_xlsx();

    let mut by_index = SheetFilter::default();
    by_index.indices_1_based.insert(2);
    let doc_by_index = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            sheet_filter: Some(by_index),
            ..Default::default()
        },
    )
    .expect("extract with index filter");
    assert_eq!(doc_by_index.sheets.len(), 1);
    assert_eq!(doc_by_index.sheets[0].name, "Data");

    let mut by_name = SheetFilter::default();
    by_name.names.insert("Sheet1".to_string());
    let doc_by_name = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            sheet_filter: Some(by_name),
            ..Default::default()
        },
    )
    .expect("extract with name filter");
    assert_eq!(doc_by_name.sheets.len(), 1);
    assert_eq!(doc_by_name.sheets[0].name, "Sheet1");
}

#[test]
fn style_aware_values_are_opt_in() {
    let content = styled_xlsx();

    let default_doc = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())
        .expect("extract default");
    let styled_doc = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: true,
                streaming_rows: false,
            },
            ..Default::default()
        },
    )
    .expect("extract style-aware");

    // The legacy default mode keeps date-styled numerics empty for backward
    // compatibility.
    assert_eq!(first_inline_text(&default_doc, 0, 0), "");
    assert_eq!(first_inline_text(&default_doc, 0, 1), "0.125");
    assert_eq!(first_inline_text(&default_doc, 0, 2), "Hello Shared");
    assert_eq!(first_inline_text(&default_doc, 0, 3), "Hello Inline");

    assert_eq!(first_inline_text(&styled_doc, 0, 0), "2024-01-01");
    assert_eq!(first_inline_text(&styled_doc, 0, 1), "12.50%");
    assert_eq!(first_inline_text(&styled_doc, 0, 2), "Hello Shared");
    assert_eq!(first_inline_text(&styled_doc, 0, 3), "Hello Inline");
}

#[test]
fn style_aware_values_respect_1904_date_system() {
    let content = styled_xlsx_1904();

    let dense = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: true,
                streaming_rows: false,
            },
            ..Default::default()
        },
    )
    .expect("extract dense");
    let streaming = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: true,
                streaming_rows: true,
            },
            ..Default::default()
        },
    )
    .expect("extract streaming");

    assert_eq!(first_inline_text(&dense, 0, 0), "2024-01-01");
    assert_eq!(first_inline_text(&streaming, 0, 0), "2024-01-01");
}

#[test]
fn streaming_rows_matches_dense_default_mode() {
    let content = minimal_xlsx();

    let dense = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())
        .expect("extract dense");
    let streaming = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: false,
                streaming_rows: true,
            },
            ..Default::default()
        },
    )
    .expect("extract streaming");

    let dense_json = serde_json::to_string(&dense).expect("serialize dense");
    let streaming_json = serde_json::to_string(&streaming).expect("serialize streaming");
    assert_eq!(dense_json, streaming_json);
}

#[test]
fn streaming_rows_matches_dense_style_aware_mode() {
    let content = styled_xlsx();

    let dense = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: true,
                streaming_rows: false,
            },
            ..Default::default()
        },
    )
    .expect("extract dense");
    let streaming = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: true,
                streaming_rows: true,
            },
            ..Default::default()
        },
    )
    .expect("extract streaming");

    let dense_json = serde_json::to_string(&dense).expect("serialize dense");
    let streaming_json = serde_json::to_string(&streaming).expect("serialize streaming");
    assert_eq!(dense_json, streaming_json);
}

#[test]
fn streaming_rows_large_sheet_parity_smoke() {
    let content = large_sheet_xlsx(300, 12);

    let dense = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())
        .expect("extract dense");
    let streaming = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: false,
                streaming_rows: true,
            },
            ..Default::default()
        },
    )
    .expect("extract streaming");

    let dense_json = serde_json::to_string(&dense).expect("serialize dense");
    let streaming_json = serde_json::to_string(&streaming).expect("serialize streaming");
    assert_eq!(dense_json, streaming_json);
}

#[test]
fn streaming_rows_handles_sparse_high_row_indices() {
    let content = sparse_high_row_xlsx(1_000_000);
    let doc = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            text: officemd_xlsx::table_ir::XlsxTextOptions {
                style_aware_values: false,
                streaming_rows: true,
            },
            ..Default::default()
        },
    )
    .expect("extract sparse streaming");

    assert_eq!(doc.sheets.len(), 1);
    assert_eq!(doc.sheets[0].tables.len(), 1);
    assert_eq!(doc.sheets[0].tables[0].rows.len(), 1);
    assert_eq!(first_inline_text(&doc, 0, 0), "42");
}

#[test]
fn document_properties_are_opt_in() {
    let content = minimal_xlsx_with_doc_props();

    let without_props = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())
        .expect("extract without properties");
    assert!(without_props.properties.is_none());

    let with_props = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            include: officemd_xlsx::table_ir::XlsxIncludeOptions {
                document_properties: true,
            },
            ..Default::default()
        },
    )
    .expect("extract with properties");
    let props = with_props
        .properties
        .as_ref()
        .expect("expected document properties");
    assert!(!props.core.is_empty());
    assert!(!props.app.is_empty());
}

#[test]
fn extract_tables_ir_json_with_options_omits_properties_by_default() {
    let content = minimal_xlsx_with_doc_props();
    let payload =
        extract_tables_ir_json_with_options(&content, false, false, false).expect("extract json");
    let value: serde_json::Value = serde_json::from_str(&payload).expect("valid json");

    assert!(value["properties"].is_null());
}

#[test]
fn extract_tables_ir_json_with_options_includes_properties_when_requested() {
    let content = minimal_xlsx_with_doc_props();
    let payload =
        extract_tables_ir_json_with_options(&content, false, false, true).expect("extract json");
    let value: serde_json::Value = serde_json::from_str(&payload).expect("valid json");

    assert!(value["properties"].is_object());
    let core = value["properties"]["core"]
        .as_object()
        .expect("core properties map");
    assert!(!core.is_empty());
}

// ---------------------------------------------------------------------------
// Fixtures for trim_empty edge cases
// ---------------------------------------------------------------------------

/// Data in A1:C3, row 4 empty, A5 has isolated value, rows 6-10 empty,
/// cols D-E empty across all rows.
fn sparse_trailing_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets><sheet name="Sparse" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    // Data in A1:C3, nothing in row 4, A5 has value, then a cell in E10 that
    // is empty (just an empty <c> tag) to push the grid wide.
    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>H1</t></is></c>
            <c r="B1" t="inlineStr"><is><t>H2</t></is></c>
            <c r="C1" t="inlineStr"><is><t>H3</t></is></c>
        </row>
        <row r="2">
            <c r="A2"><v>10</v></c>
            <c r="B2"><v>20</v></c>
            <c r="C2"><v>30</v></c>
        </row>
        <row r="3">
            <c r="A3"><v>40</v></c>
            <c r="B3"><v>50</v></c>
            <c r="C3"><v>60</v></c>
        </row>
        <row r="5">
            <c r="A5" t="inlineStr"><is><t>Footer</t></is></c>
        </row>
        <row r="10">
            <c r="E10"><v></v></c>
        </row>
    </sheetData>
</worksheet>"#;

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet),
    ])
}

/// Header row in A1:B1, data in A2:B2, isolated value in Z1.
fn wide_sparse_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets><sheet name="Wide" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
            <c r="B1" t="inlineStr"><is><t>Value</t></is></c>
            <c r="Z1" t="inlineStr"><is><t>Note</t></is></c>
        </row>
        <row r="2">
            <c r="A2" t="inlineStr"><is><t>Item</t></is></c>
            <c r="B2"><v>100</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet),
    ])
}

/// Sheet where no cell has actual text content.
fn all_empty_content_xlsx() -> Vec<u8> {
    let workbook = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets><sheet name="Blank" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    // Cells exist in XML but have no text value.
    let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1"><c r="A1"/><c r="B1"/></row>
        <row r="2"><c r="A2"/></row>
    </sheetData>
</worksheet>"#;

    build_xlsx(vec![
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet),
    ])
}

// ---------------------------------------------------------------------------
// trim_empty tests
// ---------------------------------------------------------------------------

#[test]
fn trim_empty_strips_trailing_empty_rows_and_cols() {
    let content = sparse_trailing_xlsx();

    let untrimmed = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())
        .expect("untrimmed");

    let trimmed = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            trim: officemd_xlsx::table_ir::XlsxTrimOptions { empty_edges: true },
            ..Default::default()
        },
    )
    .expect("trimmed");

    let ut = &untrimmed.sheets[0].tables[0];
    let tt = &trimmed.sheets[0].tables[0];

    // Untrimmed: 10 rows (1..10), 5 cols (A..E)
    assert_eq!(ut.rows.len(), 10);
    assert_eq!(ut.headers.len(), 5);

    // Trimmed: 5 rows (1..5), 3 cols (A..C)
    // Row 5 has "Footer" in col A so it stays; rows 6-10 are trailing empties.
    assert_eq!(tt.rows.len(), 5);
    assert_eq!(tt.headers.len(), 3);

    // All non-empty data preserved
    assert_eq!(first_inline_text(&trimmed, 0, 0), "H1");
    assert_eq!(first_inline_text(&trimmed, 0, 2), "H3");
    assert_eq!(first_inline_text(&trimmed, 4, 0), "Footer");
}

#[test]
fn trim_empty_wide_sparse_preserves_isolated_column() {
    let content = wide_sparse_xlsx();

    let untrimmed = extract_tables_ir_with_options(&content, &XlsxExtractOptions::default())
        .expect("untrimmed");

    let trimmed = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            trim: officemd_xlsx::table_ir::XlsxTrimOptions { empty_edges: true },
            ..Default::default()
        },
    )
    .expect("trimmed");

    let ut = &untrimmed.sheets[0].tables[0];
    let tt = &trimmed.sheets[0].tables[0];

    // Untrimmed: 26 cols (A..Z), 2 rows
    assert_eq!(ut.headers.len(), 26);
    assert_eq!(ut.rows.len(), 2);

    // Trimmed: Z1 has "Note" so col Z (26) is the rightmost non-empty col.
    // All 26 cols are still needed since Z has data.
    assert_eq!(tt.headers.len(), 26);
    // But row 2 only extends to B — still, trimming is about trailing empties.
    // Row count stays 2 because neither row is fully empty.
    assert_eq!(tt.rows.len(), 2);

    // Verify data integrity
    assert_eq!(first_inline_text(&trimmed, 0, 0), "Name");
    assert_eq!(first_inline_text(&trimmed, 0, 25), "Note");
    assert_eq!(first_inline_text(&trimmed, 1, 0), "Item");
}

#[test]
fn trim_empty_all_blank_preserves_min_table() {
    let content = all_empty_content_xlsx();

    let trimmed = extract_tables_ir_with_options(
        &content,
        &XlsxExtractOptions {
            trim: officemd_xlsx::table_ir::XlsxTrimOptions { empty_edges: true },
            ..Default::default()
        },
    )
    .expect("trimmed");

    let tt = &trimmed.sheets[0].tables[0];

    // Even after trimming, the spec guarantees at least 1x1 table.
    assert!(!tt.rows.is_empty());
    assert!(!tt.headers.is_empty());
}

#[test]
fn trim_empty_token_savings_on_sparse_trailing() {
    let content = sparse_trailing_xlsx();

    let untrimmed_md = officemd_xlsx::markdown_from_bytes_with_options(
        &content,
        officemd_markdown::RenderOptions {
            markdown_profile: officemd_markdown::MarkdownProfile::Human,
            ..Default::default()
        },
    )
    .expect("untrimmed markdown");

    let trimmed_md = officemd_xlsx::markdown_from_bytes_with_options(
        &content,
        officemd_markdown::RenderOptions {
            markdown_profile: officemd_markdown::MarkdownProfile::LlmCompact,
            ..Default::default()
        },
    )
    .expect("trimmed markdown");

    // LlmCompact (with trim) should produce shorter output
    assert!(
        trimmed_md.len() < untrimmed_md.len(),
        "compact/trimmed ({} chars) should be shorter than human/untrimmed ({} chars)",
        trimmed_md.len(),
        untrimmed_md.len()
    );

    // Verify no data loss: all non-empty values present in trimmed output
    for value in &[
        "H1", "H2", "H3", "10", "20", "30", "40", "50", "60", "Footer",
    ] {
        assert!(
            trimmed_md.contains(value),
            "trimmed output missing value: {value}"
        );
    }
}
