//! Generate XLSX files from the officemd IR.
//!
//! Converts an [`OoxmlDocument`] with `kind: Xlsx` into a valid `.xlsx` ZIP
//! archive that opens in Microsoft Excel and `LibreOffice` Calc.

use std::collections::HashMap;
use std::fmt::Write as _;

use officemd_core::ir::{Inline, OoxmlDocument, Paragraph, Sheet, TableCell};
use officemd_core::opc::writer::{OpcWriter, RelEntry, xml_escape_attr, xml_escape_text};

use crate::error::XlsxError;

const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const REL_TYPE_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_TYPE_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

const NS_SS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Generate an `.xlsx` file from an officemd IR document.
///
/// Each `Sheet` in the IR becomes a worksheet. Tables are written as rows
/// with headers as the first row. Formula notes are injected as `<f>` elements
/// when the cell reference matches.
///
/// # Errors
///
/// Returns an error if ZIP assembly fails.
pub fn generate_xlsx(doc: &OoxmlDocument) -> Result<Vec<u8>, XlsxError> {
    let mut w = OpcWriter::new();
    w.register_content_type_default(
        "rels",
        "application/vnd.openxmlformats-package.relationships+xml",
    );
    w.register_content_type_default("xml", "application/xml");
    w.register_content_type_override("/xl/workbook.xml", CT_WORKBOOK);

    let mut sst = SharedStringTable::new();
    let mut workbook_rels: Vec<RelEntry> = Vec::new();
    let mut sheet_entries: Vec<(String, usize)> = Vec::new(); // (name, sheet_id)

    // Build worksheets
    for (i, sheet) in doc.sheets.iter().enumerate() {
        let sheet_num = i + 1;
        let ws_path = format!("xl/worksheets/sheet{sheet_num}.xml");
        let ws_xml = build_worksheet_xml(sheet, &mut sst);

        w.register_content_type_override(&format!("/{ws_path}"), CT_WORKSHEET);
        w.add_xml_part(&ws_path, &ws_xml)?;

        let rid = format!("rId{sheet_num}");
        workbook_rels.push(RelEntry {
            id: rid,
            rel_type: REL_TYPE_WORKSHEET.to_string(),
            target: format!("worksheets/sheet{sheet_num}.xml"),
            target_mode: None,
        });

        sheet_entries.push((sheet.name.clone(), sheet_num));
    }

    // Add shared strings
    let sst_rid = format!("rId{}", doc.sheets.len() + 1);
    workbook_rels.push(RelEntry {
        id: sst_rid,
        rel_type: REL_TYPE_SHARED_STRINGS.to_string(),
        target: "sharedStrings.xml".to_string(),
        target_mode: None,
    });
    w.register_content_type_override("/xl/sharedStrings.xml", CT_SHARED_STRINGS);
    w.add_xml_part("xl/sharedStrings.xml", &sst.to_xml())?;

    // Build workbook
    let workbook_xml = build_workbook_xml(&sheet_entries);
    w.add_xml_part("xl/workbook.xml", &workbook_xml)?;
    w.add_part_rels("xl/workbook.xml", &workbook_rels)?;

    // Root relationship
    w.add_root_relationship(RelEntry {
        id: "rId1".to_string(),
        rel_type: REL_TYPE_OFFICE_DOC.to_string(),
        target: "xl/workbook.xml".to_string(),
        target_mode: None,
    });

    Ok(w.finish()?)
}

struct SharedStringTable {
    strings: Vec<String>,
    index_map: HashMap<String, usize>,
}

impl SharedStringTable {
    fn new() -> Self {
        Self {
            strings: Vec::new(),
            index_map: HashMap::new(),
        }
    }

    fn get_or_insert(&mut self, value: &str) -> usize {
        if let Some(&idx) = self.index_map.get(value) {
            return idx;
        }
        let idx = self.strings.len();
        self.strings.push(value.to_string());
        self.index_map.insert(value.to_string(), idx);
        idx
    }

    fn to_xml(&self) -> String {
        let count = self.strings.len();
        let mut xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <sst xmlns=\"{NS_SS}\" count=\"{count}\" uniqueCount=\"{count}\">"
        );
        for s in &self.strings {
            let escaped = xml_escape_text(s);
            let _ = write!(xml, "<si><t>{escaped}</t></si>");
        }
        xml.push_str("</sst>");
        xml
    }
}

/// Convert a 0-based column index to an Excel column name (A, B, ..., Z, AA, AB, ...).
fn col_to_name(mut col: usize) -> String {
    let mut name = String::new();
    loop {
        let offset = u8::try_from(col % 26).expect("column remainder is less than 26");
        name.insert(0, char::from(b'A' + offset));
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    name
}

/// Format a cell reference from 0-based row and column indices.
fn cell_ref(row: usize, col: usize) -> String {
    format!("{}{}", col_to_name(col), row + 1)
}

fn build_workbook_xml(sheets: &[(String, usize)]) -> String {
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <workbook xmlns=\"{NS_SS}\" xmlns:r=\"{NS_R}\"><sheets>"
    );
    for (name, sheet_id) in sheets {
        let escaped = xml_escape_attr(name);
        let _ = write!(
            xml,
            "<sheet name=\"{escaped}\" sheetId=\"{sheet_id}\" r:id=\"rId{sheet_id}\"/>"
        );
    }
    xml.push_str("</sheets></workbook>");
    xml
}

fn build_worksheet_xml(sheet: &Sheet, sst: &mut SharedStringTable) -> String {
    // Build formula lookup: cell_ref -> formula
    let formula_map: HashMap<&str, &str> = sheet
        .formulas
        .iter()
        .map(|f| (f.cell_ref.as_str(), f.formula.as_str()))
        .collect();

    let mut rows_xml = String::new();
    let mut current_row: usize = 0;

    let table_count = sheet.tables.len();
    for (table_idx, table) in sheet.tables.iter().enumerate() {
        // Headers row
        if !table.headers.is_empty() && !table.synthetic_headers {
            current_row += 1;
            let _ = write!(rows_xml, "<row r=\"{current_row}\">");
            for (col, header) in table.headers.iter().enumerate() {
                write_cell(&mut rows_xml, current_row, col, header, sst, &formula_map);
            }
            rows_xml.push_str("</row>");
        }

        // Data rows
        for row in &table.rows {
            current_row += 1;
            let _ = write!(rows_xml, "<row r=\"{current_row}\">");
            for (col, cell) in row.iter().enumerate() {
                let text = cell_to_text(cell);
                write_cell(&mut rows_xml, current_row, col, &text, sst, &formula_map);
            }
            rows_xml.push_str("</row>");
        }

        // Blank row separator between tables (not after the last)
        if table_idx + 1 < table_count {
            current_row += 1;
        }
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <worksheet xmlns=\"{NS_SS}\">\
         <sheetData>{rows_xml}</sheetData>\
         </worksheet>"
    )
}

/// Check if a string value should be written as a numeric cell.
///
/// Returns false for values with leading zeros (like ZIP codes "00123" or IDs
/// "007") that would lose information if coerced to numbers. Allows "0" and
/// "0.xxx" forms.
fn is_numeric_value(value: &str) -> bool {
    if value.parse::<f64>().is_err() {
        return false;
    }
    // Reject leading zeros that indicate text (ZIP codes, IDs).
    // Allow "0" and "0.xxx" but reject "00", "01", "007", etc.
    let digits = value.strip_prefix('-').unwrap_or(value);
    if digits.len() > 1 && digits.starts_with('0') && !digits.starts_with("0.") {
        return false;
    }
    true
}

fn write_cell(
    out: &mut String,
    row_1based: usize,
    col: usize,
    value: &str,
    sst: &mut SharedStringTable,
    formula_map: &HashMap<&str, &str>,
) {
    let cr = cell_ref(row_1based - 1, col);

    // Check if this cell has a formula
    if let Some(formula) = formula_map.get(cr.as_str()) {
        let escaped_formula = xml_escape_text(formula);
        // Formula cells: emit <f> with cached value in <v>
        // If the cached value is numeric, emit raw; otherwise use t="str" with <v>
        if is_numeric_value(value) {
            let _ = write!(
                out,
                "<c r=\"{cr}\"><f>{escaped_formula}</f><v>{value}</v></c>"
            );
        } else {
            let escaped_value = xml_escape_text(value);
            let _ = write!(
                out,
                "<c r=\"{cr}\" t=\"str\"><f>{escaped_formula}</f><v>{escaped_value}</v></c>"
            );
        }
    } else if value.is_empty() {
        // Empty cell — omit entirely or write minimal
        let _ = write!(out, "<c r=\"{cr}\"/>");
    } else if is_numeric_value(value) {
        // Numeric value — no type attribute needed (default is number)
        let _ = write!(out, "<c r=\"{cr}\"><v>{value}</v></c>");
    } else {
        // String value — use shared string table
        let sst_idx = sst.get_or_insert(value);
        let _ = write!(out, "<c r=\"{cr}\" t=\"s\"><v>{sst_idx}</v></c>");
    }
}

fn cell_to_text(cell: &TableCell) -> String {
    cell.content
        .iter()
        .map(paragraph_to_text)
        .collect::<Vec<_>>()
        .join("\n")
}

fn paragraph_to_text(para: &Paragraph) -> String {
    para.inlines
        .iter()
        .map(|i| match i {
            Inline::Text(t) => t.as_str(),
            Inline::Link(l) => l.display.as_str(),
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use officemd_core::ir::{
        DocumentKind, FormulaNote, Inline, OoxmlDocument, Paragraph, Sheet, Table, TableCell,
    };

    fn simple_cell(text: &str) -> TableCell {
        TableCell {
            content: vec![Paragraph {
                inlines: vec![Inline::Text(text.to_string())],
            }],
        }
    }

    fn simple_xlsx(sheets: Vec<Sheet>) -> OoxmlDocument {
        OoxmlDocument {
            kind: DocumentKind::Xlsx,
            properties: None,
            sheets,
            slides: vec![],
            sections: vec![],
            pdf: None,
        }
    }

    #[test]
    fn col_to_name_converts_correctly() {
        assert_eq!(col_to_name(0), "A");
        assert_eq!(col_to_name(1), "B");
        assert_eq!(col_to_name(25), "Z");
        assert_eq!(col_to_name(26), "AA");
        assert_eq!(col_to_name(27), "AB");
        assert_eq!(col_to_name(701), "ZZ");
        assert_eq!(col_to_name(702), "AAA");
    }

    #[test]
    fn cell_ref_formats_correctly() {
        assert_eq!(cell_ref(0, 0), "A1");
        assert_eq!(cell_ref(0, 1), "B1");
        assert_eq!(cell_ref(1, 0), "A2");
        assert_eq!(cell_ref(9, 25), "Z10");
        assert_eq!(cell_ref(0, 26), "AA1");
    }

    #[test]
    fn shared_string_table_deduplicates() {
        let mut sst = SharedStringTable::new();
        let idx0 = sst.get_or_insert("hello");
        let idx1 = sst.get_or_insert("world");
        let idx2 = sst.get_or_insert("hello");
        assert_eq!(idx0, 0);
        assert_eq!(idx1, 1);
        assert_eq!(idx2, 0); // deduplicated
        assert_eq!(sst.strings.len(), 2);
    }

    #[test]
    fn generates_valid_xlsx_single_sheet() {
        let doc = simple_xlsx(vec![Sheet {
            name: "Data".to_string(),
            tables: vec![Table {
                caption: None,
                headers: vec!["Name".to_string(), "Score".to_string()],
                rows: vec![
                    vec![simple_cell("Alice"), simple_cell("95")],
                    vec![simple_cell("Bob"), simple_cell("87")],
                ],
                synthetic_headers: false,
            }],
            formulas: vec![],
            hyperlinks: vec![],
        }]);

        let bytes = generate_xlsx(&doc).expect("generate");
        assert!(!bytes.is_empty());

        // Verify round-trip
        let ir = crate::extract_tables_ir(&bytes).expect("extract");
        assert_eq!(ir.sheets.len(), 1);
        assert_eq!(ir.sheets[0].name, "Data");
    }

    #[test]
    fn generates_xlsx_with_multiple_sheets() {
        let doc = simple_xlsx(vec![
            Sheet {
                name: "Sheet1".to_string(),
                tables: vec![Table {
                    caption: None,
                    headers: vec!["A".to_string()],
                    rows: vec![vec![simple_cell("1")]],
                    synthetic_headers: false,
                }],
                formulas: vec![],
                hyperlinks: vec![],
            },
            Sheet {
                name: "Sheet2".to_string(),
                tables: vec![Table {
                    caption: None,
                    headers: vec!["B".to_string()],
                    rows: vec![vec![simple_cell("2")]],
                    synthetic_headers: false,
                }],
                formulas: vec![],
                hyperlinks: vec![],
            },
        ]);

        let bytes = generate_xlsx(&doc).expect("generate");
        let ir = crate::extract_tables_ir(&bytes).expect("extract");
        assert_eq!(ir.sheets.len(), 2);
        assert_eq!(ir.sheets[0].name, "Sheet1");
        assert_eq!(ir.sheets[1].name, "Sheet2");
    }

    #[test]
    fn generates_xlsx_with_formulas() {
        let doc = simple_xlsx(vec![Sheet {
            name: "Calc".to_string(),
            tables: vec![Table {
                caption: None,
                headers: vec!["A".to_string(), "B".to_string(), "Sum".to_string()],
                rows: vec![vec![
                    simple_cell("10"),
                    simple_cell("20"),
                    simple_cell("30"),
                ]],
                synthetic_headers: false,
            }],
            formulas: vec![FormulaNote {
                cell_ref: "C2".to_string(),
                formula: "A2+B2".to_string(),
            }],
            hyperlinks: vec![],
        }]);

        let bytes = generate_xlsx(&doc).expect("generate");
        // Verify the formula is in the file
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let ws = pkg
            .read_part_string("xl/worksheets/sheet1.xml")
            .unwrap()
            .unwrap();
        assert!(ws.contains("<f>A2+B2</f>"), "formula should be in XML");
    }

    #[test]
    fn empty_xlsx_produces_valid_file() {
        let doc = simple_xlsx(vec![]);
        let bytes = generate_xlsx(&doc).expect("generate");
        // Should be a valid ZIP even with no sheets
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        assert!(pkg.has_part("xl/workbook.xml"));
    }

    #[test]
    fn xml_special_chars_in_cell_values() {
        let doc = simple_xlsx(vec![Sheet {
            name: "Special".to_string(),
            tables: vec![Table {
                caption: None,
                headers: vec!["Data".to_string()],
                rows: vec![vec![simple_cell("A & B < C")]],
                synthetic_headers: false,
            }],
            formulas: vec![],
            hyperlinks: vec![],
        }]);
        let bytes = generate_xlsx(&doc).expect("generate");
        let ir = crate::extract_tables_ir(&bytes).expect("extract");
        assert_eq!(ir.sheets.len(), 1);
    }

    #[test]
    fn leading_zeros_preserved_as_text() {
        let doc = simple_xlsx(vec![Sheet {
            name: "IDs".to_string(),
            tables: vec![Table {
                caption: None,
                headers: vec!["ZIP".to_string()],
                rows: vec![
                    vec![simple_cell("00123")],
                    vec![simple_cell("007")],
                    vec![simple_cell("0")],
                ],
                synthetic_headers: false,
            }],
            formulas: vec![],
            hyperlinks: vec![],
        }]);
        let bytes = generate_xlsx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let ws = pkg
            .read_part_string("xl/worksheets/sheet1.xml")
            .unwrap()
            .unwrap();
        // "00123" and "007" should be string cells (t="s"), not numeric
        assert!(
            !ws.contains("<v>00123</v>"),
            "00123 should not be a raw numeric value"
        );
        assert!(
            !ws.contains("<v>007</v>"),
            "007 should not be a raw numeric value"
        );
        // "0" is fine as numeric
        assert!(ws.contains("<v>0</v>"), "0 should be a raw numeric value");

        // Verify round-trip preserves leading zeros as text
        let ir = crate::extract_tables_ir(&bytes).expect("extract");
        let all_text: String = ir.sheets[0].tables[0]
            .rows
            .iter()
            .flat_map(|r| r.iter())
            .flat_map(|c| c.content.iter())
            .flat_map(|p| p.inlines.iter())
            .map(|i| match i {
                Inline::Text(t) => t.as_str(),
                Inline::Link(l) => l.display.as_str(),
            })
            .collect::<Vec<_>>()
            .join(" ");
        assert!(
            all_text.contains("00123"),
            "00123 should survive round-trip with leading zeros, got: {all_text}"
        );
        assert!(
            all_text.contains("007"),
            "007 should survive round-trip with leading zeros, got: {all_text}"
        );
    }

    #[test]
    fn is_numeric_value_checks() {
        assert!(is_numeric_value("42"));
        assert!(is_numeric_value("3.14"));
        assert!(is_numeric_value("-5"));
        assert!(is_numeric_value("0"));
        assert!(is_numeric_value("0.5"));
        assert!(!is_numeric_value("00123"));
        assert!(!is_numeric_value("007"));
        assert!(!is_numeric_value("01"));
        assert!(!is_numeric_value("text"));
        assert!(!is_numeric_value("-007"));
    }

    #[test]
    fn numeric_values_written_as_numbers() {
        let doc = simple_xlsx(vec![Sheet {
            name: "Nums".to_string(),
            tables: vec![Table {
                caption: None,
                headers: vec!["Value".to_string()],
                rows: vec![
                    vec![simple_cell("42")],
                    vec![simple_cell("3.14")],
                    vec![simple_cell("text")],
                ],
                synthetic_headers: false,
            }],
            formulas: vec![],
            hyperlinks: vec![],
        }]);
        let bytes = generate_xlsx(&doc).expect("generate");
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        let ws = pkg
            .read_part_string("xl/worksheets/sheet1.xml")
            .unwrap()
            .unwrap();
        // Numeric cells should NOT have t="s"
        assert!(ws.contains("<v>42</v>"), "42 should be a raw numeric value");
        assert!(
            ws.contains("<v>3.14</v>"),
            "3.14 should be a raw numeric value"
        );
        // String cells should use t="s"
        assert!(ws.contains("t=\"s\""), "text should use shared string");
    }

    #[test]
    fn round_trip_preserves_cell_content() {
        let doc = simple_xlsx(vec![Sheet {
            name: "Data".to_string(),
            tables: vec![Table {
                caption: None,
                headers: vec!["Name".to_string(), "Score".to_string()],
                rows: vec![
                    vec![simple_cell("Alice"), simple_cell("95")],
                    vec![simple_cell("Bob"), simple_cell("87")],
                ],
                synthetic_headers: false,
            }],
            formulas: vec![],
            hyperlinks: vec![],
        }]);
        let bytes = generate_xlsx(&doc).expect("generate");
        let ir = crate::extract_tables_ir(&bytes).expect("extract");
        let sheet = &ir.sheets[0];
        let table = &sheet.tables[0];
        // Check header row (becomes first data row or header depending on extraction)
        // The important thing: the data should be present somewhere
        let all_text: String = table
            .rows
            .iter()
            .flat_map(|r| r.iter())
            .flat_map(|c| c.content.iter())
            .flat_map(|p| p.inlines.iter())
            .map(|i| match i {
                Inline::Text(t) => t.as_str(),
                Inline::Link(l) => l.display.as_str(),
            })
            .collect::<Vec<_>>()
            .join(" ");
        assert!(
            all_text.contains("Alice"),
            "Alice should survive round-trip"
        );
        assert!(all_text.contains("Bob"), "Bob should survive round-trip");
    }

    #[test]
    fn sheet_name_with_special_chars() {
        let doc = simple_xlsx(vec![Sheet {
            name: "Q1 \"Results\" & Data".to_string(),
            tables: vec![Table {
                caption: None,
                headers: vec!["A".to_string()],
                rows: vec![vec![simple_cell("1")]],
                synthetic_headers: false,
            }],
            formulas: vec![],
            hyperlinks: vec![],
        }]);
        let bytes = generate_xlsx(&doc).expect("generate");
        // Should produce valid XML despite special chars in sheet name
        let mut pkg = officemd_core::opc::OpcPackage::from_bytes(&bytes).expect("valid");
        assert!(pkg.has_part("xl/workbook.xml"));
    }
}
