//! Lightweight IR extraction entrypoints for XLSX.

use officemd_core::ir::{DocumentKind, OoxmlDocument, Sheet};
use officemd_core::opc::OpcPackage;

use crate::error::XlsxError;

/// Extract sheet names (workbook order) from XLSX bytes.
/// Core sheet name extractor (Rust).
///
/// # Errors
///
/// Returns an error if the XLSX archive cannot be read or the workbook XML is malformed.
pub fn extract_sheet_names(content: &[u8]) -> Result<Vec<String>, XlsxError> {
    let mut package = OpcPackage::from_bytes(content).map_err(XlsxError::from)?;
    let mut names = Vec::new();
    if let Some(xml) = package
        .read_part_string("xl/workbook.xml")
        .map_err(XlsxError::from)?
    {
        names = collect_sheet_names(&xml);
    }
    Ok(names)
}

/// Extract minimal IR (sheets only for now) as JSON string.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed or serialized to JSON.
pub fn extract_ir_json(content: &[u8]) -> Result<String, XlsxError> {
    let doc = extract_ir(content)?;
    serde_json::to_string(&doc).map_err(|e| XlsxError::Xml(e.to_string()))
}

/// Build a minimal `OoxmlDocument` for XLSX.
///
/// # Errors
///
/// Returns an error if the XLSX content cannot be parsed.
pub fn extract_ir(content: &[u8]) -> Result<OoxmlDocument, XlsxError> {
    let sheet_names = extract_sheet_names(content)?;
    let sheets: Vec<Sheet> = sheet_names
        .into_iter()
        .map(|name| Sheet {
            name,
            ..Default::default()
        })
        .collect();

    Ok(OoxmlDocument {
        kind: DocumentKind::Xlsx,
        sheets,
        ..Default::default()
    })
}

fn collect_sheet_names(xml: &str) -> Vec<String> {
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut names = Vec::new();

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Empty(ref e) | quick_xml::events::Event::Start(ref e)) => {
                if e.name().as_ref().ends_with(b"sheet") {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"name" {
                            let name = attr.unescape_value().map_or_else(
                                |_| String::from_utf8_lossy(&attr.value).to_string(),
                                std::borrow::Cow::into_owned,
                            );
                            names.push(name);
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            Ok(_) => {}
        }
    }
    names
}
