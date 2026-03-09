use std::collections::HashMap;

use crate::rels::{Relationship, parse_relationships};

use super::OpcPackage;
use super::package::OpcError;

/// Return `_rels/*.rels` path for a part.
#[must_use]
pub fn rels_path_for_part(part_path: &str) -> String {
    let part_path = part_path.trim_start_matches('/');
    if let Some((dir, file)) = part_path.rsplit_once('/') {
        format!("{dir}/_rels/{file}.rels")
    } else {
        format!("_rels/{part_path}.rels")
    }
}

/// Return the base directory for a part.
#[must_use]
pub fn part_base_dir(part_path: &str) -> &str {
    let part_path = part_path.trim_start_matches('/');
    if let Some((dir, _)) = part_path.rsplit_once('/') {
        dir
    } else {
        ""
    }
}

/// Resolve a relationship target relative to a part.
#[must_use]
pub fn resolve_relationship_target(part_path: &str, rel: &Relationship) -> String {
    if rel
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
        || is_external_target(&rel.target)
    {
        return rel.target.clone();
    }

    normalize_target_path(part_base_dir(part_path), &rel.target)
}

/// Build `rel_id` -> resolved target map, optionally filtered by relationship type.
#[must_use]
pub fn relationship_target_map(
    rels: &[Relationship],
    part_path: &str,
    rel_type_filter: Option<&str>,
) -> HashMap<String, String> {
    let mut map = HashMap::new();
    for rel in rels {
        if rel_type_filter.is_none_or(|wanted| wanted == rel.rel_type) {
            map.insert(rel.id.clone(), resolve_relationship_target(part_path, rel));
        }
    }
    map
}

/// Load relationships for a specific part. Returns empty when no `.rels` exists.
///
/// # Errors
///
/// Returns an error if the `.rels` part cannot be read or parsed.
pub fn load_relationships_for_part(
    package: &mut OpcPackage<'_>,
    part_path: &str,
) -> Result<Vec<Relationship>, OpcError> {
    let rels_path = rels_path_for_part(part_path);
    let Some(xml) = package.read_part_string(&rels_path)? else {
        return Ok(Vec::new());
    };
    parse_relationships(&xml).map_err(|e| OpcError::Xml(e.to_string()))
}

fn is_external_target(target: &str) -> bool {
    target.starts_with("http://") || target.starts_with("https://") || target.starts_with("mailto:")
}

fn normalize_target_path(base_dir: &str, target: &str) -> String {
    if is_external_target(target) {
        return target.to_string();
    }

    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }

    let base_dir = base_dir.trim_end_matches('/');
    let combined = if base_dir.is_empty() || target.starts_with(&format!("{base_dir}/")) {
        target.to_string()
    } else {
        format!("{base_dir}/{target}")
    };

    if !combined.contains("..") && !combined.contains("./") && !combined.starts_with('.') {
        return combined;
    }

    let mut parts = Vec::new();
    for segment in combined.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(segment),
        }
    }
    parts.join("/")
}

#[cfg(test)]
mod tests {
    use std::io::Write;

    use zip::ZipWriter;
    use zip::write::FileOptions;

    use super::*;

    fn build_zip(parts: Vec<(&str, &str)>) -> Vec<u8> {
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

    #[test]
    fn resolves_relative_targets() {
        let rel = Relationship {
            id: "rId1".to_string(),
            target: "../slides/slide1.xml".to_string(),
            rel_type: "slide".to_string(),
            target_mode: None,
        };
        assert_eq!(
            resolve_relationship_target("ppt/slides/slide2.xml", &rel),
            "ppt/slides/slide1.xml"
        );
    }

    #[test]
    fn loads_rels_for_part() {
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="typeA" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
        let bytes = build_zip(vec![
            ("xl/workbook.xml", "<workbook/>"),
            ("xl/_rels/workbook.xml.rels", rels),
        ]);
        let mut package = OpcPackage::from_bytes(&bytes).expect("open package");
        let loaded = load_relationships_for_part(&mut package, "xl/workbook.xml").expect("rels");
        assert_eq!(loaded.len(), 1);
        let target_map = relationship_target_map(&loaded, "xl/workbook.xml", None);
        assert_eq!(
            target_map.get("rId1"),
            Some(&"xl/worksheets/sheet1.xml".to_string())
        );
    }
}
