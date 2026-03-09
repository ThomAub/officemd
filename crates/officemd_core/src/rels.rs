//! Relationship (.rels) parsing helpers.

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;
use thiserror::Error;

/// Relationship entry.
#[derive(Debug, Clone, Default)]
pub struct Relationship {
    pub id: String,
    pub target: String,
    pub rel_type: String,
    pub target_mode: Option<String>,
}

/// Relationship parsing error.
#[derive(Debug, Error)]
pub enum RelsError {
    #[error("XML parse error: {0}")]
    Xml(String),
}

/// Parse a .rels XML string into a vector of relationships.
///
/// # Errors
///
/// Returns `RelsError::Xml` if the XML is malformed.
pub fn parse_relationships(xml: &str) -> Result<Vec<Relationship>, RelsError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut rels = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.name().as_ref().ends_with(b"Relationship") {
                    rels.push(parse_relationship(e));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(RelsError::Xml(e.to_string())),
        }
    }

    Ok(rels)
}

fn parse_relationship(e: &BytesStart<'_>) -> Relationship {
    let mut rel = Relationship::default();
    for attr in e.attributes().flatten() {
        let value = attr_value(&attr);
        match attr.key.as_ref() {
            b"Id" => rel.id = value,
            b"Target" => rel.target = value,
            b"Type" => rel.rel_type = value,
            b"TargetMode" => rel.target_mode = Some(value),
            _ => {}
        }
    }
    rel
}

fn attr_value(attr: &quick_xml::events::attributes::Attribute<'_>) -> String {
    attr.unescape_value().map_or_else(
        |_| String::from_utf8_lossy(&attr.value).to_string(),
        std::borrow::Cow::into_owned,
    )
}

/// Build a map of `rel_id` -> target for a specific relationship type.
#[must_use]
pub fn rel_target_map(
    rels: &[Relationship],
    rel_type_filter: Option<&str>,
) -> HashMap<String, String> {
    let mut map = HashMap::new();
    for rel in rels {
        if rel_type_filter.is_none_or(|t| t == rel.rel_type) {
            map.insert(rel.id.clone(), rel.target.clone());
        }
    }
    map
}
