//! Shared extraction of OPC document properties (`docProps/`).

use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::Event;

use super::content_types::local_name;
use super::package::{OpcError, OpcPackage};
use crate::ir::DocumentProperties;

/// Extract core, app, and custom document properties from a package.
///
/// Returns `None` when all three maps are empty (no properties present).
///
/// # Errors
///
/// Returns an error if a package part cannot be read or parsed as XML.
pub fn extract_properties(
    package: &mut OpcPackage<'_>,
) -> Result<Option<DocumentProperties>, OpcError> {
    let props = DocumentProperties {
        core: extract_props_map(package, "docProps/core.xml")?,
        app: extract_props_map(package, "docProps/app.xml")?,
        custom: extract_custom_props_map(package, "docProps/custom.xml")?,
    };

    if props.core.is_empty() && props.app.is_empty() && props.custom.is_empty() {
        Ok(None)
    } else {
        Ok(Some(props))
    }
}

/// Parse a flat XML properties file into a tag-name to text-content map.
///
/// Namespace prefixes are stripped from tag names so the resulting keys are
/// always local names such as `"title"` or `"creator"`.
///
/// # Errors
///
/// Returns an error if the package part cannot be read or if XML text cannot be decoded.
pub fn extract_props_map(
    package: &mut OpcPackage<'_>,
    path: &str,
) -> Result<HashMap<String, String>, OpcError> {
    let Some(xml) = package.read_part_string(path)? else {
        return Ok(HashMap::new());
    };

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut current_tag: Option<String> = None;
    let mut map = HashMap::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                current_tag =
                    Some(String::from_utf8_lossy(local_name(e.name().as_ref())).into_owned());
            }
            Ok(Event::Text(ref t)) => {
                if let Some(tag) = &current_tag {
                    let val = t
                        .unescape()
                        .map_err(|e| OpcError::Xml(e.to_string()))?
                        .to_string();
                    map.insert(tag.clone(), val);
                }
            }
            Ok(Event::End(_)) => {
                current_tag = None;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(OpcError::Xml(e.to_string())),
        }
    }

    Ok(map)
}

/// Parse a custom properties file (`docProps/custom.xml`) into a name-value map.
///
/// # Errors
///
/// Returns an error if the package part cannot be read or if XML text cannot be decoded.
pub fn extract_custom_props_map(
    package: &mut OpcPackage<'_>,
    path: &str,
) -> Result<HashMap<String, String>, OpcError> {
    let Some(xml) = package.read_part_string(path)? else {
        return Ok(HashMap::new());
    };

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut map = HashMap::new();
    let mut current_name: Option<String> = None;
    let mut current_value = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if local_name(e.name().as_ref()) == b"property" {
                    current_name = property_name_attr(e);
                    current_value.clear();
                }
            }
            Ok(Event::Text(ref t)) => {
                if current_name.is_some() {
                    let text = t
                        .unescape()
                        .map_err(|e| OpcError::Xml(e.to_string()))?
                        .to_string();
                    current_value.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                if local_name(e.name().as_ref()) == b"property"
                    && let Some(name) = current_name.take()
                {
                    let value = current_value.trim().to_string();
                    map.insert(name, value);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(OpcError::Xml(e.to_string())),
        }
    }

    Ok(map)
}

fn property_name_attr(e: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    for attr in e.attributes().flatten() {
        if local_name(attr.key.as_ref()) == b"name" {
            if let Ok(value) = attr.unescape_value() {
                return Some(value.to_string());
            }
            return Some(String::from_utf8_lossy(&attr.value).into_owned());
        }
    }
    None
}
