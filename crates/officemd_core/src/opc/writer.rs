//! Write-side OPC package assembly.
//!
//! [`OpcWriter`] builds an OOXML ZIP archive from individual XML parts,
//! automatically generating `[Content_Types].xml` and `_rels/.rels`.

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use zip::ZipWriter;
use zip::write::FileOptions;

use super::package::OpcError;

/// Entry in a `.rels` relationship file.
#[derive(Debug, Clone)]
pub struct RelEntry {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

/// Builds an OOXML ZIP archive with automatic content-type and root relationship
/// generation.
///
/// # Examples
///
/// ```
/// use officemd_core::opc::writer::{OpcWriter, RelEntry};
///
/// let mut w = OpcWriter::new();
/// w.register_content_type_default("rels", "application/vnd.openxmlformats-package.relationships+xml");
/// w.register_content_type_default("xml", "application/xml");
/// w.register_content_type_override(
///     "/word/document.xml",
///     "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
/// );
/// w.add_xml_part("word/document.xml", r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#).unwrap();
/// w.add_root_relationship(RelEntry {
///     id: "rId1".into(),
///     rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
///     target: "word/document.xml".into(),
///     target_mode: None,
/// });
/// let bytes = w.finish().unwrap();
/// assert!(!bytes.is_empty());
/// ```
pub struct OpcWriter {
    zip: ZipWriter<Cursor<Vec<u8>>>,
    content_type_defaults: Vec<(String, String)>,
    content_type_overrides: Vec<(String, String)>,
    root_rels: Vec<RelEntry>,
}

impl OpcWriter {
    /// Create a new empty OPC package writer.
    #[must_use]
    pub fn new() -> Self {
        Self {
            zip: ZipWriter::new(Cursor::new(Vec::new())),
            content_type_defaults: Vec::new(),
            content_type_overrides: Vec::new(),
            root_rels: Vec::new(),
        }
    }

    /// Register a default content type for a file extension.
    pub fn register_content_type_default(&mut self, extension: &str, content_type: &str) {
        self.content_type_defaults
            .push((extension.to_string(), content_type.to_string()));
    }

    /// Register a content type override for a specific part path.
    /// The `part_name` should start with `/` (e.g., `/word/document.xml`).
    /// A leading `/` is prepended automatically if missing.
    pub fn register_content_type_override(&mut self, part_name: &str, content_type: &str) {
        let normalized = if part_name.starts_with('/') {
            part_name.to_string()
        } else {
            format!("/{part_name}")
        };
        self.content_type_overrides
            .push((normalized, content_type.to_string()));
    }

    /// Add a root-level relationship (written to `_rels/.rels`).
    pub fn add_root_relationship(&mut self, entry: RelEntry) {
        self.root_rels.push(entry);
    }

    /// Add an XML part to the archive.
    ///
    /// # Errors
    ///
    /// Returns an error if the ZIP entry cannot be written.
    pub fn add_xml_part(&mut self, path: &str, xml: &str) -> Result<(), OpcError> {
        self.add_part(path, xml.as_bytes())
    }

    /// Add a raw binary part to the archive.
    ///
    /// # Errors
    ///
    /// Returns an error if the ZIP entry cannot be written.
    pub fn add_part(&mut self, path: &str, data: &[u8]) -> Result<(), OpcError> {
        let options: FileOptions<'_, ()> = FileOptions::default();
        self.zip.start_file(path, options)?;
        self.zip.write_all(data)?;
        Ok(())
    }

    /// Add a part-level `.rels` file for a given part path.
    ///
    /// For example, calling `add_part_rels("word/document.xml", &rels)` writes
    /// to `word/_rels/document.xml.rels`.
    ///
    /// # Errors
    ///
    /// Returns an error if the ZIP entry cannot be written.
    pub fn add_part_rels(&mut self, part_path: &str, rels: &[RelEntry]) -> Result<(), OpcError> {
        if rels.is_empty() {
            return Ok(());
        }
        let rels_path = super::relationships::rels_path_for_part(part_path);
        let xml = serialize_relationships(rels);
        self.add_xml_part(&rels_path, &xml)
    }

    /// Finalize the archive: auto-generate `[Content_Types].xml` and
    /// `_rels/.rels`, then return the ZIP bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if final entries cannot be written or the ZIP
    /// cannot be finalized.
    pub fn finish(mut self) -> Result<Vec<u8>, OpcError> {
        // Write [Content_Types].xml
        let ct_xml = self.build_content_types_xml();
        self.add_xml_part("[Content_Types].xml", &ct_xml)?;

        // Write _rels/.rels
        let rels_xml = serialize_relationships(&self.root_rels);
        self.add_xml_part("_rels/.rels", &rels_xml)?;

        let cursor = self.zip.finish()?;
        Ok(cursor.into_inner())
    }

    fn build_content_types_xml(&self) -> String {
        let mut xml = String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
        );
        for (ext, ct) in &self.content_type_defaults {
            let _ = write!(
                xml,
                "<Default Extension=\"{}\" ContentType=\"{}\"/>",
                xml_escape_attr(ext),
                xml_escape_attr(ct),
            );
        }
        for (part, ct) in &self.content_type_overrides {
            let _ = write!(
                xml,
                "<Override PartName=\"{}\" ContentType=\"{}\"/>",
                xml_escape_attr(part),
                xml_escape_attr(ct),
            );
        }
        xml.push_str("</Types>");
        xml
    }
}

impl Default for OpcWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// Serialize a list of relationship entries to an OPC `.rels` XML string.
#[must_use]
pub fn serialize_relationships(rels: &[RelEntry]) -> String {
    let mut xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for rel in rels {
        let _ = write!(
            xml,
            "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"",
            xml_escape_attr(&rel.id),
            xml_escape_attr(&rel.rel_type),
            xml_escape_attr(&rel.target),
        );
        if let Some(mode) = &rel.target_mode {
            let _ = write!(xml, " TargetMode=\"{}\"", xml_escape_attr(mode));
        }
        xml.push_str("/>");
    }
    xml.push_str("</Relationships>");
    xml
}

/// Escape text for inclusion in XML text nodes.
#[must_use]
pub fn xml_escape_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(c),
        }
    }
    out
}

/// Escape a value for inclusion in an XML attribute.
/// Handles both double-quoted and single-quoted attribute contexts.
#[must_use]
pub fn xml_escape_attr(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::opc::OpcPackage;
    use crate::rels::parse_relationships;

    #[test]
    fn opc_writer_creates_valid_package() {
        let mut w = OpcWriter::new();
        w.register_content_type_default(
            "rels",
            "application/vnd.openxmlformats-package.relationships+xml",
        );
        w.register_content_type_default("xml", "application/xml");
        w.register_content_type_override(
            "/word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        );
        w.add_xml_part("word/document.xml", "<doc/>").unwrap();
        w.add_root_relationship(RelEntry {
            id: "rId1".into(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    .into(),
            target: "word/document.xml".into(),
            target_mode: None,
        });

        let bytes = w.finish().unwrap();
        let mut pkg = OpcPackage::from_bytes(&bytes).expect("valid OPC package");
        assert!(pkg.has_part("word/document.xml"));
        assert_eq!(
            pkg.read_part_string("word/document.xml")
                .unwrap()
                .as_deref(),
            Some("<doc/>"),
        );
    }

    #[test]
    fn serialize_relationships_produces_parseable_xml() {
        let rels = vec![
            RelEntry {
                id: "rId1".into(),
                rel_type: "http://example.com/type".into(),
                target: "target.xml".into(),
                target_mode: None,
            },
            RelEntry {
                id: "rId2".into(),
                rel_type: "http://example.com/external".into(),
                target: "https://example.com".into(),
                target_mode: Some("External".into()),
            },
        ];
        let xml = serialize_relationships(&rels);
        let parsed = parse_relationships(&xml).expect("valid rels XML");
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].id, "rId1");
        assert_eq!(parsed[0].target, "target.xml");
        assert_eq!(parsed[1].id, "rId2");
        assert_eq!(parsed[1].target_mode.as_deref(), Some("External"),);
    }

    #[test]
    fn part_rels_written_at_correct_path() {
        let mut w = OpcWriter::new();
        w.register_content_type_default("xml", "application/xml");
        w.register_content_type_default(
            "rels",
            "application/vnd.openxmlformats-package.relationships+xml",
        );
        w.add_xml_part("word/document.xml", "<doc/>").unwrap();
        w.add_part_rels(
            "word/document.xml",
            &[RelEntry {
                id: "rId1".into(),
                rel_type: "http://example.com/hyperlink".into(),
                target: "https://example.com".into(),
                target_mode: Some("External".into()),
            }],
        )
        .unwrap();
        w.add_root_relationship(RelEntry {
            id: "rId1".into(),
            rel_type: "officeDocument".into(),
            target: "word/document.xml".into(),
            target_mode: None,
        });

        let bytes = w.finish().unwrap();
        let mut pkg = OpcPackage::from_bytes(&bytes).expect("valid package");
        assert!(pkg.has_part("word/_rels/document.xml.rels"));
        let rels_xml = pkg
            .read_part_string("word/_rels/document.xml.rels")
            .unwrap()
            .unwrap();
        let parsed = parse_relationships(&rels_xml).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].target, "https://example.com");
    }

    #[test]
    fn xml_escape_text_handles_special_chars() {
        assert_eq!(xml_escape_text("a & b < c > d"), "a &amp; b &lt; c &gt; d");
        assert_eq!(xml_escape_text("normal"), "normal");
    }

    #[test]
    fn xml_escape_attr_handles_quotes() {
        assert_eq!(xml_escape_attr(r#"a "b" & c"#), "a &quot;b&quot; &amp; c");
        assert_eq!(xml_escape_attr("it's"), "it&apos;s");
    }

    #[test]
    fn content_type_override_normalizes_leading_slash() {
        let mut w = OpcWriter::new();
        w.register_content_type_default("xml", "application/xml");
        w.register_content_type_default(
            "rels",
            "application/vnd.openxmlformats-package.relationships+xml",
        );
        // Pass without leading slash — should be normalized
        w.register_content_type_override("word/document.xml", "application/test");
        w.add_xml_part("word/document.xml", "<doc/>").unwrap();
        w.add_root_relationship(RelEntry {
            id: "rId1".into(),
            rel_type: "officeDocument".into(),
            target: "word/document.xml".into(),
            target_mode: None,
        });
        let bytes = w.finish().unwrap();
        let mut pkg = OpcPackage::from_bytes(&bytes).expect("valid");
        // The content type should be findable for the part
        assert!(pkg.has_part("word/document.xml"));
    }

    #[test]
    fn relationship_target_with_ampersand_is_escaped() {
        let rels = vec![RelEntry {
            id: "rId1".into(),
            rel_type: "http://example.com/type".into(),
            target: "https://example.com?a=1&b=2".into(),
            target_mode: Some("External".into()),
        }];
        let xml = serialize_relationships(&rels);
        assert!(
            xml.contains("Target=\"https://example.com?a=1&amp;b=2\""),
            "ampersand in URL should be escaped: {xml}",
        );
        // Must still parse correctly
        let parsed = parse_relationships(&xml).expect("valid rels XML");
        assert_eq!(parsed[0].target, "https://example.com?a=1&b=2");
    }

    #[test]
    fn empty_rels_not_written() {
        let mut w = OpcWriter::new();
        w.register_content_type_default("xml", "application/xml");
        w.register_content_type_default(
            "rels",
            "application/vnd.openxmlformats-package.relationships+xml",
        );
        w.add_xml_part("word/document.xml", "<doc/>").unwrap();
        // Empty rels — should not create a file
        w.add_part_rels("word/document.xml", &[]).unwrap();
        w.add_root_relationship(RelEntry {
            id: "rId1".into(),
            rel_type: "officeDocument".into(),
            target: "word/document.xml".into(),
            target_mode: None,
        });

        let bytes = w.finish().unwrap();
        let mut pkg = OpcPackage::from_bytes(&bytes).expect("valid package");
        assert!(!pkg.has_part("word/_rels/document.xml.rels"));
    }
}
