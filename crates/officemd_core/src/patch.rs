use rayon::ThreadPoolBuilder;
use rayon::prelude::*;
use regex::{Regex, RegexBuilder};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
use thiserror::Error;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

#[derive(Debug, Error)]
pub enum PatchError {
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("json error: {0}")]
    Json(#[from] serde_json::Error),
    #[error("regex error: {0}")]
    Regex(#[from] regex::Error),
    #[error("thread pool error: {0}")]
    ThreadPool(#[from] rayon::ThreadPoolBuildError),
    #[error("missing part: {0}")]
    MissingPart(String),
    #[error("text not found in {part}: {needle}")]
    TextNotFound { part: String, needle: String },
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
#[serde(rename_all = "snake_case")]
pub enum ReplaceMode {
    First,
    #[default]
    All,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
#[serde(rename_all = "snake_case")]
pub enum MatchPolicy {
    #[default]
    Exact,
    CaseInsensitive,
    WholeWord,
    WholeWordCaseInsensitive,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum DocxTextScope {
    Body,
    Headers,
    Footers,
    Comments,
    Footnotes,
    Endnotes,
    MetadataCoreTitle,
    AllText,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum PptxTextScope {
    SlideTitles,
    SlideBody,
    Notes,
    Comments,
    MetadataCoreTitle,
    AllText,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct TextReplace {
    pub from: String,
    pub to: String,
    #[serde(default)]
    pub mode: ReplaceMode,
    #[serde(default)]
    pub match_policy: MatchPolicy,
}

impl TextReplace {
    pub fn all(from: impl Into<String>, to: impl Into<String>) -> Self {
        Self {
            from: from.into(),
            to: to.into(),
            mode: ReplaceMode::All,
            match_policy: MatchPolicy::Exact,
        }
    }

    pub fn first(from: impl Into<String>, to: impl Into<String>) -> Self {
        Self {
            from: from.into(),
            to: to.into(),
            mode: ReplaceMode::First,
            match_policy: MatchPolicy::Exact,
        }
    }

    pub fn with_match_policy(mut self, match_policy: MatchPolicy) -> Self {
        self.match_policy = match_policy;
        self
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ScopedDocxReplace {
    pub scope: DocxTextScope,
    pub replace: TextReplace,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ScopedPptxReplace {
    pub scope: PptxTextScope,
    pub replace: TextReplace,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq, Eq)]
pub struct DocxPatch {
    #[serde(default)]
    pub set_core_title: Option<String>,
    #[serde(default)]
    pub replace_body_title: Option<TextReplace>,
    #[serde(default)]
    pub scoped_replacements: Vec<ScopedDocxReplace>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq, Eq)]
pub struct PptxPatch {
    #[serde(default)]
    pub set_core_title: Option<String>,
    #[serde(default)]
    pub scoped_replacements: Vec<ScopedPptxReplace>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct OoxmlTextEdit {
    pub part: String,
    pub from: String,
    pub to: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq, Eq)]
pub struct OoxmlPatchRequest {
    #[serde(default)]
    pub edits: Vec<OoxmlTextEdit>,
    #[serde(default)]
    pub core_title: Option<String>,
}

pub fn patch_docx(content: &[u8], patch: &DocxPatch) -> Result<Vec<u8>, PatchError> {
    let mut parts = read_parts(content)?;
    apply_docx_patch_to_parts(&mut parts, patch)?;
    write_parts(&parts)
}

pub fn patch_pptx(content: &[u8], patch: &PptxPatch) -> Result<Vec<u8>, PatchError> {
    let mut parts = read_parts(content)?;
    apply_pptx_patch_to_parts(&mut parts, patch)?;
    write_parts(&parts)
}

pub fn patch_docx_batch(
    contents: Vec<Vec<u8>>,
    patch: &DocxPatch,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    run_batch(contents, workers, |content| patch_docx(&content, patch))
}

pub fn patch_pptx_batch(
    contents: Vec<Vec<u8>>,
    patch: &PptxPatch,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    run_batch(contents, workers, |content| patch_pptx(&content, patch))
}

pub fn apply_ooxml_patch(
    content: &[u8],
    request: &OoxmlPatchRequest,
) -> Result<Vec<u8>, PatchError> {
    let mut parts = read_parts(content)?;
    apply_low_level_patch_to_parts(&mut parts, request)?;
    write_parts(&parts)
}

pub fn apply_ooxml_patch_json(content: &[u8], request_json: &str) -> Result<Vec<u8>, PatchError> {
    let request: OoxmlPatchRequest = serde_json::from_str(request_json)?;
    apply_ooxml_patch(content, &request)
}

pub fn patch_docx_json(content: &[u8], patch_json: &str) -> Result<Vec<u8>, PatchError> {
    let patch: DocxPatch = serde_json::from_str(patch_json)?;
    patch_docx(content, &patch)
}

pub fn patch_pptx_json(content: &[u8], patch_json: &str) -> Result<Vec<u8>, PatchError> {
    let patch: PptxPatch = serde_json::from_str(patch_json)?;
    patch_pptx(content, &patch)
}

pub fn patch_docx_batch_json(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    let patch: DocxPatch = serde_json::from_str(patch_json)?;
    patch_docx_batch(contents, &patch, workers)
}

pub fn patch_pptx_batch_json(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    let patch: PptxPatch = serde_json::from_str(patch_json)?;
    patch_pptx_batch(contents, &patch, workers)
}

fn run_batch<F>(
    contents: Vec<Vec<u8>>,
    workers: Option<usize>,
    job: F,
) -> Result<Vec<Vec<u8>>, PatchError>
where
    F: Fn(Vec<u8>) -> Result<Vec<u8>, PatchError> + Sync + Send,
{
    let worker_count =
        workers.unwrap_or_else(|| std::thread::available_parallelism().map_or(1, usize::from));
    if worker_count <= 1 || contents.len() <= 1 {
        return contents.into_iter().map(job).collect();
    }

    let pool = ThreadPoolBuilder::new().num_threads(worker_count).build()?;
    pool.install(|| contents.into_par_iter().map(job).collect())
}

fn apply_docx_patch_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    patch: &DocxPatch,
) -> Result<(), PatchError> {
    let part_names: Vec<String> = parts.keys().cloned().collect();

    if let Some(replace) = &patch.replace_body_title {
        apply_replace_to_named_part(parts, "word/document.xml", replace)?;
    }

    let mut metadata_requested = false;
    for scoped in &patch.scoped_replacements {
        if scoped.scope == DocxTextScope::MetadataCoreTitle {
            metadata_requested = true;
            apply_core_title_replace(parts, &scoped.replace)?;
            continue;
        }

        let targets = docx_scope_targets(&part_names, scoped.scope);
        apply_replace_to_parts(parts, &targets, &scoped.replace)?;
    }

    if patch.set_core_title.is_some() || metadata_requested {
        apply_core_title(parts, patch.set_core_title.as_deref())?;
    }

    Ok(())
}

fn apply_pptx_patch_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    patch: &PptxPatch,
) -> Result<(), PatchError> {
    let part_names: Vec<String> = parts.keys().cloned().collect();

    let mut metadata_requested = false;
    for scoped in &patch.scoped_replacements {
        if scoped.scope == PptxTextScope::MetadataCoreTitle {
            metadata_requested = true;
            apply_core_title_replace(parts, &scoped.replace)?;
            continue;
        }

        let targets = pptx_scope_targets(&part_names, scoped.scope);
        apply_replace_to_parts(parts, &targets, &scoped.replace)?;
    }

    if patch.set_core_title.is_some() || metadata_requested {
        apply_core_title(parts, patch.set_core_title.as_deref())?;
    }

    Ok(())
}

fn apply_low_level_patch_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    request: &OoxmlPatchRequest,
) -> Result<(), PatchError> {
    for edit in &request.edits {
        let replace = TextReplace::first(&edit.from, &edit.to);
        apply_replace_to_named_part(parts, &edit.part, &replace)?;
    }
    apply_core_title(parts, request.core_title.as_deref())?;
    Ok(())
}

fn apply_replace_to_named_part(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part: &str,
    replace: &TextReplace,
) -> Result<(), PatchError> {
    let data = parts
        .get_mut(part)
        .ok_or_else(|| PatchError::MissingPart(part.to_string()))?;
    let updated = apply_replace_to_text(&String::from_utf8_lossy(data), replace)?;
    if updated == String::from_utf8_lossy(data) {
        return Err(PatchError::TextNotFound {
            part: part.to_string(),
            needle: replace.from.clone(),
        });
    }
    *data = updated.into_bytes();
    Ok(())
}

fn apply_replace_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_parts: &[String],
    replace: &TextReplace,
) -> Result<(), PatchError> {
    let mut matched = false;
    for part in target_parts {
        let Some(data) = parts.get_mut(part) else {
            continue;
        };
        let original = String::from_utf8_lossy(data).into_owned();
        let updated = apply_replace_to_text(&original, replace)?;
        if updated != original {
            *data = updated.into_bytes();
            matched = true;
        }
    }

    if matched {
        Ok(())
    } else {
        Err(PatchError::TextNotFound {
            part: target_parts.join(","),
            needle: replace.from.clone(),
        })
    }
}

fn apply_replace_to_text(text: &str, replace: &TextReplace) -> Result<String, PatchError> {
    if replace.from.is_empty() {
        return Ok(text.to_string());
    }

    let updated = match (replace.match_policy, replace.mode) {
        (MatchPolicy::Exact, ReplaceMode::All) => text.replace(&replace.from, &replace.to),
        (MatchPolicy::Exact, ReplaceMode::First) => text.replacen(&replace.from, &replace.to, 1),
        _ => {
            let regex = build_replace_regex(replace)?;
            let replaced = match replace.mode {
                ReplaceMode::All => regex.replace_all(text, replace.to.as_str()),
                ReplaceMode::First => regex.replace(text, replace.to.as_str()),
            };
            replaced.into_owned()
        }
    };
    Ok(updated)
}

fn build_replace_regex(replace: &TextReplace) -> Result<Regex, PatchError> {
    let escaped = regex::escape(&replace.from);
    let pattern = match replace.match_policy {
        MatchPolicy::Exact | MatchPolicy::CaseInsensitive => escaped,
        MatchPolicy::WholeWord | MatchPolicy::WholeWordCaseInsensitive => {
            format!(r"\b{}\b", escaped)
        }
    };
    let mut builder = RegexBuilder::new(&pattern);
    builder.case_insensitive(matches!(
        replace.match_policy,
        MatchPolicy::CaseInsensitive | MatchPolicy::WholeWordCaseInsensitive
    ));
    Ok(builder.build()?)
}

fn docx_scope_targets(part_names: &[String], scope: DocxTextScope) -> Vec<String> {
    let mut targets = BTreeSet::new();
    for name in part_names {
        let matches = match scope {
            DocxTextScope::Body => name == "word/document.xml",
            DocxTextScope::Headers => name.starts_with("word/header") && name.ends_with(".xml"),
            DocxTextScope::Footers => name.starts_with("word/footer") && name.ends_with(".xml"),
            DocxTextScope::Comments => name == "word/comments.xml",
            DocxTextScope::Footnotes => name == "word/footnotes.xml",
            DocxTextScope::Endnotes => name == "word/endnotes.xml",
            DocxTextScope::MetadataCoreTitle => name == "docProps/core.xml",
            DocxTextScope::AllText => {
                name == "word/document.xml"
                    || (name.starts_with("word/header") && name.ends_with(".xml"))
                    || (name.starts_with("word/footer") && name.ends_with(".xml"))
                    || name == "word/comments.xml"
                    || name == "word/footnotes.xml"
                    || name == "word/endnotes.xml"
            }
        };
        if matches {
            targets.insert(name.clone());
        }
    }
    targets.into_iter().collect()
}

fn pptx_scope_targets(part_names: &[String], scope: PptxTextScope) -> Vec<String> {
    let mut targets = BTreeSet::new();
    for name in part_names {
        let matches = match scope {
            PptxTextScope::SlideTitles | PptxTextScope::SlideBody => {
                name.starts_with("ppt/slides/slide") && name.ends_with(".xml")
            }
            PptxTextScope::Notes => {
                name.starts_with("ppt/notesSlides/notesSlide") && name.ends_with(".xml")
            }
            PptxTextScope::Comments => {
                name.starts_with("ppt/comments/comment") && name.ends_with(".xml")
            }
            PptxTextScope::MetadataCoreTitle => name == "docProps/core.xml",
            PptxTextScope::AllText => {
                (name.starts_with("ppt/slides/slide") && name.ends_with(".xml"))
                    || (name.starts_with("ppt/notesSlides/notesSlide") && name.ends_with(".xml"))
                    || (name.starts_with("ppt/comments/comment") && name.ends_with(".xml"))
            }
        };
        if matches {
            targets.insert(name.clone());
        }
    }
    targets.into_iter().collect()
}

fn apply_core_title(
    parts: &mut BTreeMap<String, Vec<u8>>,
    title: Option<&str>,
) -> Result<(), PatchError> {
    let Some(title) = title else {
        return Ok(());
    };

    let core_xml = parts
        .get_mut("docProps/core.xml")
        .ok_or_else(|| PatchError::MissingPart("docProps/core.xml".to_string()))?;
    let text = String::from_utf8_lossy(core_xml).into_owned();
    let updated = if text.contains("<dc:title/>") {
        text.replacen(
            "<dc:title/>",
            &format!("<dc:title>{}</dc:title>", xml_escape(title)),
            1,
        )
    } else if text.contains("<dc:title></dc:title>") {
        text.replacen(
            "<dc:title></dc:title>",
            &format!("<dc:title>{}</dc:title>", xml_escape(title)),
            1,
        )
    } else if let Some(start) = text.find("<dc:title>") {
        let end = text[start..]
            .find("</dc:title>")
            .map(|idx| start + idx)
            .ok_or_else(|| PatchError::TextNotFound {
                part: "docProps/core.xml".to_string(),
                needle: "</dc:title>".to_string(),
            })?;
        format!(
            "{}<dc:title>{}</dc:title>{}",
            &text[..start],
            xml_escape(title),
            &text[end + "</dc:title>".len()..]
        )
    } else if let Some(insert_at) = text.find("</cp:coreProperties>") {
        format!(
            "{}<dc:title>{}</dc:title>{}",
            &text[..insert_at],
            xml_escape(title),
            &text[insert_at..]
        )
    } else {
        return Err(PatchError::TextNotFound {
            part: "docProps/core.xml".to_string(),
            needle: "</cp:coreProperties>".to_string(),
        });
    };
    *core_xml = updated.into_bytes();
    Ok(())
}

fn apply_core_title_replace(
    parts: &mut BTreeMap<String, Vec<u8>>,
    replace: &TextReplace,
) -> Result<(), PatchError> {
    apply_replace_to_named_part(parts, "docProps/core.xml", replace)
}

fn read_parts(content: &[u8]) -> Result<BTreeMap<String, Vec<u8>>, PatchError> {
    let mut archive = ZipArchive::new(Cursor::new(content))?;
    let mut parts = BTreeMap::new();
    for idx in 0..archive.len() {
        let mut file = archive.by_index(idx)?;
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)?;
        parts.insert(file.name().to_string(), bytes);
    }
    Ok(parts)
}

fn write_parts(parts: &BTreeMap<String, Vec<u8>>) -> Result<Vec<u8>, PatchError> {
    let mut cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut cursor);
    let options: FileOptions<'_, ()> = FileOptions::default();
    for (name, bytes) in parts {
        writer.start_file(name, options)?;
        writer.write_all(bytes)?;
    }
    writer.finish()?;
    Ok(cursor.into_inner())
}

fn xml_escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_helpers::build_zip;

    #[test]
    fn applies_text_edits_and_core_title() {
        let bytes = build_zip(vec![
            ("word/document.xml", "<w:t>Old Title</w:t>"),
            ("word/comments.xml", "<w:t>Old Comment</w:t>"),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title/></cp:coreProperties>",
            ),
        ]);

        let patched = apply_ooxml_patch(
            &bytes,
            &OoxmlPatchRequest {
                edits: vec![
                    OoxmlTextEdit {
                        part: "word/document.xml".to_string(),
                        from: "Old Title".to_string(),
                        to: "New Title".to_string(),
                    },
                    OoxmlTextEdit {
                        part: "word/comments.xml".to_string(),
                        from: "Old Comment".to_string(),
                        to: "New Comment".to_string(),
                    },
                ],
                core_title: Some("Core Title".to_string()),
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(String::from_utf8_lossy(&parts["word/document.xml"]).contains("New Title"));
        assert!(String::from_utf8_lossy(&parts["word/comments.xml"]).contains("New Comment"));
        assert!(String::from_utf8_lossy(&parts["docProps/core.xml"]).contains("Core Title"));
    }

    #[test]
    fn typed_docx_patch_replaces_all_text_scopes() {
        let bytes = build_zip(vec![
            ("word/document.xml", "<w:t>word body word</w:t>"),
            ("word/header1.xml", "<w:t>word header</w:t>"),
            ("word/comments.xml", "<w:t>word comment</w:t>"),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title>old</dc:title></cp:coreProperties>",
            ),
        ]);

        let patched = patch_docx(
            &bytes,
            &DocxPatch {
                set_core_title: Some("new core".to_string()),
                replace_body_title: None,
                scoped_replacements: vec![ScopedDocxReplace {
                    scope: DocxTextScope::AllText,
                    replace: TextReplace::all("word", "term")
                        .with_match_policy(MatchPolicy::WholeWord),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(String::from_utf8_lossy(&parts["word/document.xml"]).contains("term body term"));
        assert!(String::from_utf8_lossy(&parts["word/header1.xml"]).contains("term header"));
        assert!(String::from_utf8_lossy(&parts["word/comments.xml"]).contains("term comment"));
        assert!(String::from_utf8_lossy(&parts["docProps/core.xml"]).contains("new core"));
    }

    #[test]
    fn typed_pptx_patch_replaces_comments_and_notes() {
        let bytes = build_zip(vec![
            ("ppt/slides/slide1.xml", "<a:t>word slide</a:t>"),
            ("ppt/notesSlides/notesSlide1.xml", "<a:t>word notes</a:t>"),
            ("ppt/comments/comment1.xml", "<a:t>word comment</a:t>"),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title/></cp:coreProperties>",
            ),
        ]);

        let patched = patch_pptx(
            &bytes,
            &PptxPatch {
                set_core_title: Some("deck title".to_string()),
                scoped_replacements: vec![ScopedPptxReplace {
                    scope: PptxTextScope::AllText,
                    replace: TextReplace::all("word", "term"),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(String::from_utf8_lossy(&parts["ppt/slides/slide1.xml"]).contains("term slide"));
        assert!(
            String::from_utf8_lossy(&parts["ppt/notesSlides/notesSlide1.xml"])
                .contains("term notes")
        );
        assert!(
            String::from_utf8_lossy(&parts["ppt/comments/comment1.xml"]).contains("term comment")
        );
        assert!(String::from_utf8_lossy(&parts["docProps/core.xml"]).contains("deck title"));
    }

    #[test]
    fn case_insensitive_first_replace_works() {
        let updated = apply_replace_to_text(
            "Word word WORD",
            &TextReplace {
                from: "word".to_string(),
                to: "term".to_string(),
                mode: ReplaceMode::First,
                match_policy: MatchPolicy::CaseInsensitive,
            },
        )
        .unwrap();
        assert_eq!(updated, "term word WORD");
    }

    #[test]
    fn patch_docx_batch_works() {
        let bytes = build_zip(vec![("word/document.xml", "<w:t>word</w:t>")]);
        let patched = patch_docx_batch(
            vec![bytes.clone(), bytes],
            &DocxPatch {
                set_core_title: None,
                replace_body_title: None,
                scoped_replacements: vec![ScopedDocxReplace {
                    scope: DocxTextScope::AllText,
                    replace: TextReplace::all("word", "term"),
                }],
            },
            Some(2),
        )
        .unwrap();
        assert_eq!(patched.len(), 2);
        for item in patched {
            let parts = read_parts(&item).unwrap();
            assert!(String::from_utf8_lossy(&parts["word/document.xml"]).contains("term"));
        }
    }

    #[test]
    fn patch_json_round_trip_works() {
        let patch = DocxPatch {
            set_core_title: Some("Title".to_string()),
            replace_body_title: Some(TextReplace::first("Old", "New")),
            scoped_replacements: vec![ScopedDocxReplace {
                scope: DocxTextScope::AllText,
                replace: TextReplace::all("word", "term"),
            }],
        };
        let json = serde_json::to_string(&patch).unwrap();
        let decoded: DocxPatch = serde_json::from_str(&json).unwrap();
        assert_eq!(decoded, patch);
    }

    #[test]
    fn errors_when_edit_text_missing() {
        let bytes = build_zip(vec![("word/document.xml", "<w:t>Hello</w:t>")]);
        let err = apply_ooxml_patch(
            &bytes,
            &OoxmlPatchRequest {
                edits: vec![OoxmlTextEdit {
                    part: "word/document.xml".to_string(),
                    from: "Missing".to_string(),
                    to: "New".to_string(),
                }],
                core_title: None,
            },
        )
        .unwrap_err();
        assert!(matches!(err, PatchError::TextNotFound { .. }));
    }
}
