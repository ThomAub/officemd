use std::borrow::Cow;
use std::collections::HashMap;
use std::io::{Cursor, Read};
use std::sync::Arc;

use thiserror::Error;
use zip::ZipArchive;
use zip::result::ZipError;

use super::content_types::ContentTypes;
use super::part::OpcPart;

const MAX_PART_UNCOMPRESSED_BYTES: u64 = 64 * 1024 * 1024;
const MAX_PACKAGE_UNCOMPRESSED_BYTES: u64 = 256 * 1024 * 1024;

/// Errors from the shared OPC package layer.
#[derive(Debug, Error)]
pub enum OpcError {
    #[error("ZIP error: {0}")]
    Zip(#[from] ZipError),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("XML parse error: {0}")]
    Xml(String),
    #[error("Missing required part: {0}")]
    MissingPart(String),
}

/// Read-only OOXML package accessor.
pub struct OpcPackage<'a> {
    archive: ZipArchive<Cursor<&'a [u8]>>,
    content_types: ContentTypes,
    part_cache: HashMap<String, Arc<[u8]>>,
    total_part_bytes_read: u64,
}

impl<'a> OpcPackage<'a> {
    /// Open a package from bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the ZIP archive is invalid or content types cannot be parsed.
    pub fn from_bytes(content: &'a [u8]) -> Result<Self, OpcError> {
        let mut archive = ZipArchive::new(Cursor::new(content))?;
        let content_types = match read_zip_string(&mut archive, "[Content_Types].xml")? {
            Some(xml) => ContentTypes::parse(&xml)?,
            None => ContentTypes::default(),
        };

        Ok(Self {
            archive,
            content_types,
            part_cache: HashMap::new(),
            total_part_bytes_read: 0,
        })
    }

    /// Return `true` if the package contains a part.
    pub fn has_part(&mut self, path: &str) -> bool {
        let normalized = normalize_part_path(path);
        self.archive.by_name(normalized.as_ref()).is_ok()
    }

    /// Read a part as bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the part exceeds size limits or cannot be read.
    pub fn read_part_bytes(&mut self, path: &str) -> Result<Option<Arc<[u8]>>, OpcError> {
        let normalized = normalize_part_path(path);
        if let Some(cached) = self.part_cache.get(normalized.as_ref()) {
            return Ok(Some(Arc::clone(cached)));
        }
        match self.archive.by_name(normalized.as_ref()) {
            Ok(mut file) => {
                let part_size = file.size();
                ensure_part_size_within_limit(normalized.as_ref(), part_size)?;
                ensure_total_size_within_limit(self.total_part_bytes_read, part_size)?;

                let mut out = Vec::with_capacity(usize::try_from(part_size).unwrap_or(0));
                file.read_to_end(&mut out)?;

                self.total_part_bytes_read =
                    self.total_part_bytes_read.saturating_add(out.len() as u64);
                let bytes: Arc<[u8]> = Arc::from(out);
                self.part_cache
                    .insert(normalized.into_owned(), Arc::clone(&bytes));
                Ok(Some(bytes))
            }
            Err(ZipError::FileNotFound) => Ok(None),
            Err(e) => Err(OpcError::Zip(e)),
        }
    }

    /// Read a part as UTF-8 string (lossy conversion for non UTF-8 bytes).
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying byte read fails.
    pub fn read_part_string(&mut self, path: &str) -> Result<Option<String>, OpcError> {
        let normalized = normalize_part_path(path);
        match self.archive.by_name(normalized.as_ref()) {
            Ok(mut file) => {
                let part_size = file.size();
                ensure_part_size_within_limit(normalized.as_ref(), part_size)?;
                ensure_total_size_within_limit(self.total_part_bytes_read, part_size)?;

                let mut out = Vec::with_capacity(usize::try_from(part_size).unwrap_or(0));
                file.read_to_end(&mut out)?;
                self.total_part_bytes_read =
                    self.total_part_bytes_read.saturating_add(out.len() as u64);
                match String::from_utf8(out) {
                    Ok(s) => Ok(Some(s)),
                    Err(e) => Ok(Some(String::from_utf8_lossy(e.as_bytes()).into_owned())),
                }
            }
            Err(ZipError::FileNotFound) => Ok(None),
            Err(e) => Err(OpcError::Zip(e)),
        }
    }

    /// Read a required part as string.
    ///
    /// # Errors
    ///
    /// Returns `OpcError::MissingPart` if the part does not exist.
    pub fn read_required_part_string(&mut self, path: &str) -> Result<String, OpcError> {
        self.read_part_string(path)?
            .ok_or_else(|| OpcError::MissingPart(normalize_part_path(path).into_owned()))
    }

    /// Build lightweight metadata for a part path.
    pub fn part(&self, path: &str) -> OpcPart {
        let normalized = normalize_part_path(path);
        let content_type = self
            .content_types
            .content_type_for_part(normalized.as_ref())
            .map(ToString::to_string);
        OpcPart {
            path: normalized.into_owned(),
            content_type,
        }
    }

    /// Return content type for a part path.
    #[must_use]
    pub fn part_content_type(&self, path: &str) -> Option<&str> {
        let normalized = normalize_part_path(path);
        self.content_types
            .content_type_for_part(normalized.as_ref())
    }

    /// List package parts filtered by prefix/suffix.
    pub fn list_parts(&mut self, prefix: &str, suffix: &str) -> Vec<String> {
        let prefix = normalize_part_path(prefix);
        let mut parts = Vec::new();
        for idx in 0..self.archive.len() {
            if let Ok(file) = self.archive.by_index(idx) {
                let name = file.name();
                if name.starts_with(prefix.as_ref()) && name.ends_with(suffix) {
                    parts.push(name.to_string());
                }
            }
        }
        parts.sort();
        parts
    }
}

pub(crate) fn normalize_part_path(path: &str) -> Cow<'_, str> {
    if let Some(stripped) = path.strip_prefix('/') {
        Cow::Owned(stripped.to_string())
    } else {
        Cow::Borrowed(path)
    }
}

fn read_zip_string(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    path: &str,
) -> Result<Option<String>, OpcError> {
    match archive.by_name(path) {
        Ok(mut file) => {
            ensure_part_size_within_limit(path, file.size())?;
            let mut out = String::new();
            file.read_to_string(&mut out)?;
            Ok(Some(out))
        }
        Err(ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(OpcError::Zip(e)),
    }
}

fn ensure_part_size_within_limit(path: &str, bytes: u64) -> Result<(), OpcError> {
    if bytes > MAX_PART_UNCOMPRESSED_BYTES {
        return Err(OpcError::Xml(format!(
            "Part '{path}' exceeds maximum uncompressed size ({bytes} bytes > {MAX_PART_UNCOMPRESSED_BYTES} bytes)",
        )));
    }
    Ok(())
}

fn ensure_total_size_within_limit(current_total: u64, next_part: u64) -> Result<(), OpcError> {
    let projected = current_total.saturating_add(next_part);
    if projected > MAX_PACKAGE_UNCOMPRESSED_BYTES {
        return Err(OpcError::Xml(format!(
            "Package exceeds maximum uncompressed size ({projected} bytes > {MAX_PACKAGE_UNCOMPRESSED_BYTES} bytes)",
        )));
    }
    Ok(())
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
    fn reads_parts_and_lists() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;
        let bytes = build_zip(vec![
            ("[Content_Types].xml", content_types),
            ("xl/workbook.xml", "<workbook/>"),
            ("docProps/core.xml", "<coreProperties/>"),
        ]);

        let mut package = OpcPackage::from_bytes(&bytes).expect("open package");
        assert!(package.has_part("xl/workbook.xml"));
        assert_eq!(
            package
                .read_required_part_string("/xl/workbook.xml")
                .expect("read required"),
            "<workbook/>"
        );

        let parts = package.list_parts("docProps/", ".xml");
        assert_eq!(parts, vec!["docProps/core.xml".to_string()]);

        assert_eq!(
            package.part_content_type("xl/workbook.xml"),
            Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        );
    }

    #[test]
    fn missing_required_part_returns_error() {
        let bytes = build_zip(vec![("xl/workbook.xml", "<workbook/>")]);
        let mut package = OpcPackage::from_bytes(&bytes).expect("open package");
        let err = package
            .read_required_part_string("xl/missing.xml")
            .unwrap_err();
        assert!(matches!(err, OpcError::MissingPart(_)));
    }

    #[test]
    fn read_part_bytes_reuses_cached_arc() {
        let bytes = build_zip(vec![("xl/workbook.xml", "<workbook/>")]);
        let mut package = OpcPackage::from_bytes(&bytes).expect("open package");

        let first = package
            .read_part_bytes("xl/workbook.xml")
            .expect("first read")
            .expect("part exists");
        let second = package
            .read_part_bytes("/xl/workbook.xml")
            .expect("second read")
            .expect("part exists");

        assert!(Arc::ptr_eq(&first, &second));
        assert_eq!(first.as_ref(), b"<workbook/>");
    }

    #[test]
    fn rejects_oversized_part_by_metadata() {
        let err = ensure_part_size_within_limit("xl/huge.xml", MAX_PART_UNCOMPRESSED_BYTES + 1)
            .expect_err("size check should fail");
        assert!(matches!(err, OpcError::Xml(_)));
    }

    #[test]
    fn rejects_oversized_total_budget() {
        let err = ensure_total_size_within_limit(MAX_PACKAGE_UNCOMPRESSED_BYTES, 1)
            .expect_err("budget check should fail");
        assert!(matches!(err, OpcError::Xml(_)));
    }
}
