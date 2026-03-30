use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::package::OpcError;

/// Parsed `[Content_Types].xml` data.
#[derive(Debug, Clone, Default)]
pub struct ContentTypes {
    default_types: HashMap<String, String>,
    override_types: HashMap<String, String>,
}

impl ContentTypes {
    /// Parse content types XML.
    ///
    /// # Errors
    ///
    /// Returns `OpcError::Xml` if the XML is malformed.
    pub fn parse(xml: &str) -> Result<Self, OpcError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut parsed = ContentTypes::default();
        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                    match local_name(e.name().as_ref()) {
                        b"Default" => {
                            if let (Some(ext), Some(content_type)) =
                                (attr_value(e, b"Extension"), attr_value(e, b"ContentType"))
                            {
                                parsed
                                    .default_types
                                    .insert(ext.to_ascii_lowercase(), content_type);
                            }
                        }
                        b"Override" => {
                            if let (Some(part_name), Some(content_type)) =
                                (attr_value(e, b"PartName"), attr_value(e, b"ContentType"))
                            {
                                parsed.override_types.insert(part_name, content_type);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => return Err(OpcError::Xml(e.to_string())),
            }
        }

        Ok(parsed)
    }

    /// Resolve content type for a part path.
    pub fn content_type_for_part(&self, path: &str) -> Option<&str> {
        let normalized = path.trim_start_matches('/');
        let override_key = format!("/{normalized}");
        if let Some(content_type) = self.override_types.get(&override_key) {
            return Some(content_type.as_str());
        }

        let ext = normalized
            .rsplit_once('.')
            .map(|(_, ext)| ext.to_ascii_lowercase())?;
        self.default_types.get(&ext).map(String::as_str)
    }
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let attr_key = local_name(attr.key.as_ref());
        if attr_key == key {
            if let Ok(value) = attr.unescape_value() {
                return Some(value.into_owned());
            }
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

pub(crate) fn local_name(name: &[u8]) -> &[u8] {
    if let Some(idx) = name.iter().rposition(|b| *b == b':') {
        &name[idx + 1..]
    } else if let Some(idx) = name.iter().rposition(|b| *b == b'}') {
        &name[idx + 1..]
    } else {
        name
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_defaults_and_overrides() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

        let types = ContentTypes::parse(xml).expect("parse");
        assert_eq!(
            types.content_type_for_part("xl/workbook.xml"),
            Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        );
        assert_eq!(
            types.content_type_for_part("xl/_rels/workbook.xml.rels"),
            Some("application/vnd.openxmlformats-package.relationships+xml")
        );
        assert_eq!(
            types.content_type_for_part("docProps/core.xml"),
            Some("application/xml")
        );
    }
}
