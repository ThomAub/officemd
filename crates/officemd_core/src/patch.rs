use rayon::ThreadPoolBuilder;
use rayon::prelude::*;
use regex::{Regex, RegexBuilder};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
use std::sync::LazyLock;
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
    MetadataCore,
    MetadataApp,
    MetadataCustom,
    MetadataAll,
    AllText,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum PptxTextScope {
    SlideTitles,
    SlideBody,
    Notes,
    Comments,
    CommentAuthors,
    MetadataCoreTitle,
    MetadataCore,
    MetadataApp,
    MetadataCustom,
    MetadataAll,
    AllText,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum XlsxTextScope {
    SheetNames,
    Headers,
    CellText,
    SharedStrings,
    InlineStrings,
    Comments,
    CommentAuthors,
    MetadataCoreTitle,
    MetadataCore,
    MetadataApp,
    MetadataCustom,
    MetadataAll,
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
    #[serde(default)]
    pub preserve_formatting: bool,
}

impl TextReplace {
    pub fn all(from: impl Into<String>, to: impl Into<String>) -> Self {
        Self {
            from: from.into(),
            to: to.into(),
            mode: ReplaceMode::All,
            match_policy: MatchPolicy::Exact,
            preserve_formatting: false,
        }
    }

    pub fn first(from: impl Into<String>, to: impl Into<String>) -> Self {
        Self {
            from: from.into(),
            to: to.into(),
            mode: ReplaceMode::First,
            match_policy: MatchPolicy::Exact,
            preserve_formatting: false,
        }
    }

    pub fn with_match_policy(mut self, match_policy: MatchPolicy) -> Self {
        self.match_policy = match_policy;
        self
    }

    pub fn with_preserve_formatting(mut self, preserve_formatting: bool) -> Self {
        self.preserve_formatting = preserve_formatting;
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

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ScopedXlsxReplace {
    pub scope: XlsxTextScope,
    pub replace: TextReplace,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct XlsxSheetRename {
    pub from: String,
    pub to: String,
    #[serde(default = "default_true")]
    pub update_references: bool,
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

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq, Eq)]
pub struct XlsxPatch {
    #[serde(default)]
    pub set_core_title: Option<String>,
    #[serde(default)]
    pub rename_sheets: Vec<XlsxSheetRename>,
    #[serde(default)]
    pub scoped_replacements: Vec<ScopedXlsxReplace>,
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

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq, Eq)]
pub struct PatchReport {
    pub parts_scanned: usize,
    pub parts_modified: usize,
    pub replacements_applied: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct PatchedDocument {
    pub content: Vec<u8>,
    pub report: PatchReport,
}

fn default_true() -> bool {
    true
}

pub fn patch_docx(content: &[u8], patch: &DocxPatch) -> Result<Vec<u8>, PatchError> {
    Ok(patch_docx_with_report(content, patch)?.content)
}

pub fn patch_pptx(content: &[u8], patch: &PptxPatch) -> Result<Vec<u8>, PatchError> {
    Ok(patch_pptx_with_report(content, patch)?.content)
}

pub fn patch_xlsx(content: &[u8], patch: &XlsxPatch) -> Result<Vec<u8>, PatchError> {
    Ok(patch_xlsx_with_report(content, patch)?.content)
}

pub fn patch_docx_with_report(
    content: &[u8],
    patch: &DocxPatch,
) -> Result<PatchedDocument, PatchError> {
    let mut parts = read_parts(content)?;
    let mut report = PatchReport::default();
    apply_docx_patch_to_parts(&mut parts, patch, &mut report)?;
    Ok(PatchedDocument {
        content: write_parts(&parts)?,
        report,
    })
}

pub fn patch_pptx_with_report(
    content: &[u8],
    patch: &PptxPatch,
) -> Result<PatchedDocument, PatchError> {
    let mut parts = read_parts(content)?;
    let mut report = PatchReport::default();
    apply_pptx_patch_to_parts(&mut parts, patch, &mut report)?;
    Ok(PatchedDocument {
        content: write_parts(&parts)?,
        report,
    })
}

pub fn patch_xlsx_with_report(
    content: &[u8],
    patch: &XlsxPatch,
) -> Result<PatchedDocument, PatchError> {
    let mut parts = read_parts(content)?;
    let mut report = PatchReport::default();
    apply_xlsx_patch_to_parts(&mut parts, patch, &mut report)?;
    Ok(PatchedDocument {
        content: write_parts(&parts)?,
        report,
    })
}

pub fn patch_docx_batch(
    contents: Vec<Vec<u8>>,
    patch: &DocxPatch,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    Ok(patch_docx_batch_with_report(contents, patch, workers)?
        .into_iter()
        .map(|item| item.content)
        .collect())
}

pub fn patch_pptx_batch(
    contents: Vec<Vec<u8>>,
    patch: &PptxPatch,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    Ok(patch_pptx_batch_with_report(contents, patch, workers)?
        .into_iter()
        .map(|item| item.content)
        .collect())
}

pub fn patch_xlsx_batch(
    contents: Vec<Vec<u8>>,
    patch: &XlsxPatch,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    Ok(patch_xlsx_batch_with_report(contents, patch, workers)?
        .into_iter()
        .map(|item| item.content)
        .collect())
}

pub fn patch_docx_batch_with_report(
    contents: Vec<Vec<u8>>,
    patch: &DocxPatch,
    workers: Option<usize>,
) -> Result<Vec<PatchedDocument>, PatchError> {
    run_batch(contents, workers, |content| {
        patch_docx_with_report(&content, patch)
    })
}

pub fn patch_pptx_batch_with_report(
    contents: Vec<Vec<u8>>,
    patch: &PptxPatch,
    workers: Option<usize>,
) -> Result<Vec<PatchedDocument>, PatchError> {
    run_batch(contents, workers, |content| {
        patch_pptx_with_report(&content, patch)
    })
}

pub fn patch_xlsx_batch_with_report(
    contents: Vec<Vec<u8>>,
    patch: &XlsxPatch,
    workers: Option<usize>,
) -> Result<Vec<PatchedDocument>, PatchError> {
    run_batch(contents, workers, |content| {
        patch_xlsx_with_report(&content, patch)
    })
}

pub fn apply_ooxml_patch(
    content: &[u8],
    request: &OoxmlPatchRequest,
) -> Result<Vec<u8>, PatchError> {
    let mut parts = read_parts(content)?;
    let mut report = PatchReport::default();
    apply_low_level_patch_to_parts(&mut parts, request, &mut report)?;
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

pub fn patch_xlsx_json(content: &[u8], patch_json: &str) -> Result<Vec<u8>, PatchError> {
    let patch: XlsxPatch = serde_json::from_str(patch_json)?;
    patch_xlsx(content, &patch)
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

pub fn patch_xlsx_batch_json(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<Vec<u8>>, PatchError> {
    let patch: XlsxPatch = serde_json::from_str(patch_json)?;
    patch_xlsx_batch(contents, &patch, workers)
}

pub fn patch_docx_batch_json_with_report(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<PatchedDocument>, PatchError> {
    let patch: DocxPatch = serde_json::from_str(patch_json)?;
    patch_docx_batch_with_report(contents, &patch, workers)
}

pub fn patch_pptx_batch_json_with_report(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<PatchedDocument>, PatchError> {
    let patch: PptxPatch = serde_json::from_str(patch_json)?;
    patch_pptx_batch_with_report(contents, &patch, workers)
}

pub fn patch_xlsx_batch_json_with_report(
    contents: Vec<Vec<u8>>,
    patch_json: &str,
    workers: Option<usize>,
) -> Result<Vec<PatchedDocument>, PatchError> {
    let patch: XlsxPatch = serde_json::from_str(patch_json)?;
    patch_xlsx_batch_with_report(contents, &patch, workers)
}

fn run_batch<T, F>(
    contents: Vec<Vec<u8>>,
    workers: Option<usize>,
    job: F,
) -> Result<Vec<T>, PatchError>
where
    F: Fn(Vec<u8>) -> Result<T, PatchError> + Sync + Send,
    T: Send,
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
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let part_names: Vec<String> = parts.keys().cloned().collect();

    if let Some(replace) = &patch.replace_body_title {
        apply_replace_to_named_part(parts, "word/document.xml", replace, report)?;
    }

    let mut metadata_requested = false;
    for scoped in &patch.scoped_replacements {
        if scoped.scope == DocxTextScope::MetadataCoreTitle {
            metadata_requested = true;
            apply_core_title_replace(parts, &scoped.replace, report)?;
            continue;
        }

        if scoped.replace.preserve_formatting {
            apply_docx_scoped_replace_preserving_formatting(parts, &part_names, scoped, report)?;
        } else {
            let targets = docx_scope_targets(&part_names, scoped.scope);
            apply_replace_to_parts(parts, &targets, &scoped.replace, report)?;
        }
    }

    if patch.set_core_title.is_some() || metadata_requested {
        apply_core_title(parts, patch.set_core_title.as_deref(), report)?;
    }

    Ok(())
}

fn apply_pptx_patch_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    patch: &PptxPatch,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let part_names: Vec<String> = parts.keys().cloned().collect();

    let mut metadata_requested = false;
    for scoped in &patch.scoped_replacements {
        if scoped.scope == PptxTextScope::MetadataCoreTitle {
            metadata_requested = true;
            apply_core_title_replace(parts, &scoped.replace, report)?;
            continue;
        }

        if scoped.replace.preserve_formatting {
            apply_pptx_scoped_replace_preserving_formatting(parts, &part_names, scoped, report)?;
        } else {
            let targets = pptx_scope_targets(&part_names, scoped.scope);
            apply_replace_to_parts(parts, &targets, &scoped.replace, report)?;
        }
    }

    if patch.set_core_title.is_some() || metadata_requested {
        apply_core_title(parts, patch.set_core_title.as_deref(), report)?;
    }

    Ok(())
}

fn apply_xlsx_patch_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    patch: &XlsxPatch,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let part_names: Vec<String> = parts.keys().cloned().collect();

    for rename in &patch.rename_sheets {
        apply_xlsx_sheet_rename(parts, rename, report)?;
    }

    let mut metadata_requested = false;
    for scoped in &patch.scoped_replacements {
        if scoped.replace.preserve_formatting {
            apply_xlsx_scoped_replace_preserving_formatting(parts, &part_names, scoped, report)?;
            continue;
        }

        match scoped.scope {
            XlsxTextScope::MetadataCoreTitle => {
                metadata_requested = true;
                apply_core_title_replace(parts, &scoped.replace, report)?;
            }
            XlsxTextScope::SheetNames => {
                apply_xlsx_sheet_name_replace(parts, &scoped.replace, report)?;
            }
            XlsxTextScope::SharedStrings => {
                apply_replace_to_xml_text_nodes_named_part(
                    parts,
                    "xl/sharedStrings.xml",
                    &XML_TEXT_NODE_RE,
                    &scoped.replace,
                    report,
                )?;
            }
            XlsxTextScope::InlineStrings => {
                let targets = xlsx_inline_string_targets(&part_names);
                apply_replace_to_xml_text_nodes_in_parts(
                    parts,
                    &targets,
                    &XML_TEXT_NODE_RE,
                    &scoped.replace,
                    report,
                )?;
            }
            XlsxTextScope::Headers | XlsxTextScope::CellText => {
                apply_xlsx_workbook_text_replace(parts, &part_names, &scoped.replace, report)?;
            }
            XlsxTextScope::Comments
            | XlsxTextScope::CommentAuthors
            | XlsxTextScope::MetadataCore
            | XlsxTextScope::MetadataApp
            | XlsxTextScope::MetadataCustom
            | XlsxTextScope::MetadataAll => {
                let targets = xlsx_scope_targets(&part_names, scoped.scope);
                apply_replace_to_parts(parts, &targets, &scoped.replace, report)?;
            }
            XlsxTextScope::AllText => {
                apply_xlsx_all_text_replace(parts, &part_names, &scoped.replace, report)?;
            }
        }
    }

    if patch.set_core_title.is_some() || metadata_requested {
        apply_core_title(parts, patch.set_core_title.as_deref(), report)?;
    }

    Ok(())
}

fn apply_low_level_patch_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    request: &OoxmlPatchRequest,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    for edit in &request.edits {
        let replace = TextReplace::first(&edit.from, &edit.to);
        apply_replace_to_named_part(parts, &edit.part, &replace, report)?;
    }
    apply_core_title(parts, request.core_title.as_deref(), report)?;
    Ok(())
}

fn apply_replace_to_named_part(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part: &str,
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    report.parts_scanned += 1;
    let data = parts
        .get_mut(part)
        .ok_or_else(|| PatchError::MissingPart(part.to_string()))?;
    let original = String::from_utf8_lossy(data).into_owned();
    let (updated, replacements_applied) = apply_replace_to_text(&original, replace)?;
    if replacements_applied == 0 {
        return Err(PatchError::TextNotFound {
            part: part.to_string(),
            needle: replace.from.clone(),
        });
    }
    *data = updated.into_bytes();
    report.parts_modified += 1;
    report.replacements_applied += replacements_applied;
    Ok(())
}

fn apply_replace_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_parts: &[String],
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;
    for part in target_parts {
        report.parts_scanned += 1;
        let Some(data) = parts.get_mut(part) else {
            continue;
        };
        let original = String::from_utf8_lossy(data).into_owned();
        let (updated, replacements_applied) = apply_replace_to_text(&original, replace)?;
        if replacements_applied > 0 {
            *data = updated.into_bytes();
            report.parts_modified += 1;
            report.replacements_applied += replacements_applied;
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

fn apply_preserve_formatting_replace_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_parts: &[String],
    container_re: &Regex,
    node_re: &Regex,
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;
    for part in target_parts {
        report.parts_scanned += 1;
        let Some(data) = parts.get_mut(part) else {
            continue;
        };
        let original = String::from_utf8_lossy(data).into_owned();
        let (updated, replacements_applied) =
            rewrite_xml_text_containers_preserving_runs(&original, container_re, node_re, replace)?;
        if replacements_applied > 0 {
            *data = updated.into_bytes();
            report.parts_modified += 1;
            report.replacements_applied += replacements_applied;
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

fn try_apply_preserve_formatting_replace_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_parts: &[String],
    container_re: &Regex,
    node_re: &Regex,
    replace: &TextReplace,
    report: &mut PatchReport,
) -> bool {
    !target_parts.is_empty()
        && apply_preserve_formatting_replace_to_parts(
            parts,
            target_parts,
            container_re,
            node_re,
            replace,
            report,
        )
        .is_ok()
}

fn try_apply_simple_replace_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_parts: &[String],
    replace: &TextReplace,
    report: &mut PatchReport,
) -> bool {
    !target_parts.is_empty() && apply_replace_to_parts(parts, target_parts, replace, report).is_ok()
}

fn apply_replace_to_text(text: &str, replace: &TextReplace) -> Result<(String, usize), PatchError> {
    if replace.from.is_empty() {
        return Ok((text.to_string(), 0));
    }

    let (updated, replacements_applied) = match (replace.match_policy, replace.mode) {
        (MatchPolicy::Exact, ReplaceMode::All) => {
            let count = text.match_indices(&replace.from).count();
            (text.replace(&replace.from, &replace.to), count)
        }
        (MatchPolicy::Exact, ReplaceMode::First) => {
            let count = usize::from(text.contains(&replace.from));
            (text.replacen(&replace.from, &replace.to, 1), count)
        }
        _ => {
            let regex = build_replace_regex(replace)?;
            let count = match replace.mode {
                ReplaceMode::All => regex.find_iter(text).count(),
                ReplaceMode::First => usize::from(regex.find(text).is_some()),
            };
            let replaced = match replace.mode {
                ReplaceMode::All => regex.replace_all(text, replace.to.as_str()),
                ReplaceMode::First => regex.replace(text, replace.to.as_str()),
            };
            (replaced.into_owned(), count)
        }
    };
    Ok((updated, replacements_applied))
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

fn find_match_ranges(text: &str, replace: &TextReplace) -> Result<Vec<(usize, usize)>, PatchError> {
    if replace.from.is_empty() {
        return Ok(Vec::new());
    }

    match (replace.match_policy, replace.mode) {
        (MatchPolicy::Exact, ReplaceMode::All) => Ok(text
            .match_indices(&replace.from)
            .map(|(start, matched)| (start, start + matched.len()))
            .collect()),
        (MatchPolicy::Exact, ReplaceMode::First) => Ok(text
            .find(&replace.from)
            .map(|start| (start, start + replace.from.len()))
            .into_iter()
            .collect()),
        _ => {
            let regex = build_replace_regex(replace)?;
            let iter = regex.find_iter(text).map(|m| (m.start(), m.end()));
            Ok(match replace.mode {
                ReplaceMode::All => iter.collect(),
                ReplaceMode::First => iter.take(1).collect(),
            })
        }
    }
}

static XML_TEXT_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<t(?:\s+xml:space="preserve")?>(.*?)</t>"#).unwrap());
static WORD_TEXT_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<w:t(?:\s+[^>]*)?>(.*?)</w:t>"#).unwrap());
static WORD_PARAGRAPH_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<w:p(?:\s+[^>]*)?>.*?</w:p>"#).unwrap());
static DRAWING_TEXT_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<a:t(?:\s+[^>]*)?>(.*?)</a:t>"#).unwrap());
static DRAWING_PARAGRAPH_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<a:p(?:\s+[^>]*)?>.*?</a:p>"#).unwrap());
static PPT_COMMENT_TEXT_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<p:text(?:\s+[^>]*)?>(.*?)</p:text>"#).unwrap());
static PPT_COMMENT_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<p:cm(?:\s+[^>]*)?>.*?</p:cm>"#).unwrap());
static XLSX_SHARED_STRING_ITEM_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<si(?:\s+[^>]*)?>.*?</si>"#).unwrap());
static XLSX_INLINE_STRING_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<is(?:\s+[^>]*)?>.*?</is>"#).unwrap());
static XLSX_COMMENT_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"(?s)<comment(?:\s+[^>]*)?>.*?</comment>"#).unwrap());
static XLSX_THREADED_COMMENT_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r#"(?s)<threadedComment(?:\s+[^>]*)?>.*?</threadedComment>"#).unwrap()
});
static FORMULA_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<f(?:\s[^>]*)?>(.*?)</f>").unwrap());
static CHART_FORMULA_NODE_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<c:f>(.*?)</c:f>").unwrap());
static DEFINED_NAME_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)<definedName\b[^>]*>(.*?)</definedName>").unwrap());
static SHEET_NAME_ATTR_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"<sheet\b([^>]*)\bname=\"([^\"]*)\"([^>]*)/>"#).unwrap());
static QUOTED_SHEET_REF_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"'((?:[^']|'')+)'!").unwrap());
static UNQUOTED_SHEET_REF_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\b([A-Za-z_][A-Za-z0-9_\.]*)!").unwrap());

fn docx_scope_matches(name: &str, scope: DocxTextScope) -> bool {
    match scope {
        DocxTextScope::Body => name == "word/document.xml",
        DocxTextScope::Headers => name.starts_with("word/header") && name.ends_with(".xml"),
        DocxTextScope::Footers => name.starts_with("word/footer") && name.ends_with(".xml"),
        DocxTextScope::Comments => name == "word/comments.xml",
        DocxTextScope::Footnotes => name == "word/footnotes.xml",
        DocxTextScope::Endnotes => name == "word/endnotes.xml",
        DocxTextScope::MetadataCoreTitle | DocxTextScope::MetadataCore => {
            name == "docProps/core.xml"
        }
        DocxTextScope::MetadataApp => name == "docProps/app.xml",
        DocxTextScope::MetadataCustom => name == "docProps/custom.xml",
        DocxTextScope::MetadataAll => {
            name == "docProps/core.xml"
                || name == "docProps/app.xml"
                || name == "docProps/custom.xml"
        }
        DocxTextScope::AllText => {
            const SCOPES: &[DocxTextScope] = &[
                DocxTextScope::Body,
                DocxTextScope::Headers,
                DocxTextScope::Footers,
                DocxTextScope::Comments,
                DocxTextScope::Footnotes,
                DocxTextScope::Endnotes,
                DocxTextScope::MetadataAll,
            ];
            SCOPES.iter().any(|s| docx_scope_matches(name, *s))
        }
    }
}

fn docx_scope_targets(part_names: &[String], scope: DocxTextScope) -> Vec<String> {
    let mut targets = BTreeSet::new();
    for name in part_names {
        if docx_scope_matches(name, scope) {
            targets.insert(name.clone());
        }
    }
    targets.into_iter().collect()
}

fn pptx_scope_matches(name: &str, scope: PptxTextScope) -> bool {
    match scope {
        PptxTextScope::SlideTitles | PptxTextScope::SlideBody => {
            name.starts_with("ppt/slides/slide") && name.ends_with(".xml")
        }
        PptxTextScope::Notes => {
            name.starts_with("ppt/notesSlides/notesSlide") && name.ends_with(".xml")
        }
        PptxTextScope::Comments => {
            name.starts_with("ppt/comments/comment") && name.ends_with(".xml")
        }
        PptxTextScope::CommentAuthors => name == "ppt/commentAuthors.xml",
        PptxTextScope::MetadataCoreTitle | PptxTextScope::MetadataCore => {
            name == "docProps/core.xml"
        }
        PptxTextScope::MetadataApp => name == "docProps/app.xml",
        PptxTextScope::MetadataCustom => name == "docProps/custom.xml",
        PptxTextScope::MetadataAll => {
            name == "docProps/core.xml"
                || name == "docProps/app.xml"
                || name == "docProps/custom.xml"
        }
        PptxTextScope::AllText => {
            const SCOPES: &[PptxTextScope] = &[
                PptxTextScope::SlideTitles,
                PptxTextScope::Notes,
                PptxTextScope::Comments,
                PptxTextScope::CommentAuthors,
                PptxTextScope::MetadataAll,
            ];
            SCOPES.iter().any(|s| pptx_scope_matches(name, *s))
        }
    }
}

fn pptx_scope_targets(part_names: &[String], scope: PptxTextScope) -> Vec<String> {
    let mut targets = BTreeSet::new();
    for name in part_names {
        if pptx_scope_matches(name, scope) {
            targets.insert(name.clone());
        }
    }
    targets.into_iter().collect()
}

fn xlsx_inline_string_targets(part_names: &[String]) -> Vec<String> {
    part_names
        .iter()
        .filter(|name| name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
        .cloned()
        .collect()
}

fn xlsx_scope_targets(part_names: &[String], scope: XlsxTextScope) -> Vec<String> {
    let mut targets = BTreeSet::new();
    for name in part_names {
        let matches = match scope {
            XlsxTextScope::Comments => {
                name.starts_with("xl/comments") && name.ends_with(".xml")
                    || (name.starts_with("xl/threadedComments/") && name.ends_with(".xml"))
            }
            XlsxTextScope::CommentAuthors => {
                name.starts_with("xl/persons/") && name.ends_with(".xml")
            }
            XlsxTextScope::MetadataCoreTitle | XlsxTextScope::MetadataCore => {
                name == "docProps/core.xml"
            }
            XlsxTextScope::MetadataApp => name == "docProps/app.xml",
            XlsxTextScope::MetadataCustom => name == "docProps/custom.xml",
            XlsxTextScope::MetadataAll => {
                name == "docProps/core.xml"
                    || name == "docProps/app.xml"
                    || name == "docProps/custom.xml"
            }
            _ => false,
        };
        if matches {
            targets.insert(name.clone());
        }
    }
    targets.into_iter().collect()
}

fn xlsx_formula_targets(part_names: &[String]) -> Vec<String> {
    part_names
        .iter()
        .filter(|name| {
            (name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
                || (name.starts_with("xl/charts/chart") && name.ends_with(".xml"))
        })
        .cloned()
        .collect()
}

fn apply_docx_scoped_replace_preserving_formatting(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_names: &[String],
    scoped: &ScopedDocxReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;

    let content_scope = match scoped.scope {
        DocxTextScope::MetadataCoreTitle
        | DocxTextScope::MetadataCore
        | DocxTextScope::MetadataApp
        | DocxTextScope::MetadataCustom
        | DocxTextScope::MetadataAll => None,
        other => Some(other),
    };

    if let Some(scope) = content_scope {
        let targets = docx_scope_targets(part_names, scope);
        matched |= try_apply_preserve_formatting_replace_to_parts(
            parts,
            &targets,
            &WORD_PARAGRAPH_RE,
            &WORD_TEXT_NODE_RE,
            &scoped.replace,
            report,
        );
    }

    if matches!(scoped.scope, DocxTextScope::AllText) {
        let metadata_targets = docx_scope_targets(part_names, DocxTextScope::MetadataAll);
        matched |=
            try_apply_simple_replace_to_parts(parts, &metadata_targets, &scoped.replace, report);
    }

    if matched {
        Ok(())
    } else {
        let targets = docx_scope_targets(part_names, scoped.scope);
        apply_replace_to_parts(parts, &targets, &scoped.replace, report)
    }
}

fn apply_pptx_scoped_replace_preserving_formatting(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_names: &[String],
    scoped: &ScopedPptxReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;

    let slide_targets = match scoped.scope {
        PptxTextScope::SlideTitles | PptxTextScope::SlideBody | PptxTextScope::AllText => {
            pptx_scope_targets(part_names, PptxTextScope::SlideBody)
        }
        _ => Vec::new(),
    };
    matched |= try_apply_preserve_formatting_replace_to_parts(
        parts,
        &slide_targets,
        &DRAWING_PARAGRAPH_RE,
        &DRAWING_TEXT_NODE_RE,
        &scoped.replace,
        report,
    );

    let notes_targets = match scoped.scope {
        PptxTextScope::Notes | PptxTextScope::AllText => {
            pptx_scope_targets(part_names, PptxTextScope::Notes)
        }
        _ => Vec::new(),
    };
    matched |= try_apply_preserve_formatting_replace_to_parts(
        parts,
        &notes_targets,
        &DRAWING_PARAGRAPH_RE,
        &DRAWING_TEXT_NODE_RE,
        &scoped.replace,
        report,
    );

    let comment_targets = match scoped.scope {
        PptxTextScope::Comments | PptxTextScope::AllText => {
            pptx_scope_targets(part_names, PptxTextScope::Comments)
        }
        _ => Vec::new(),
    };
    matched |= try_apply_preserve_formatting_replace_to_parts(
        parts,
        &comment_targets,
        &PPT_COMMENT_RE,
        &PPT_COMMENT_TEXT_NODE_RE,
        &scoped.replace,
        report,
    );

    if matches!(scoped.scope, PptxTextScope::AllText) {
        for scope in [PptxTextScope::CommentAuthors, PptxTextScope::MetadataAll] {
            let targets = pptx_scope_targets(part_names, scope);
            matched |= try_apply_simple_replace_to_parts(parts, &targets, &scoped.replace, report);
        }
    }

    if matched {
        Ok(())
    } else {
        let targets = pptx_scope_targets(part_names, scoped.scope);
        apply_replace_to_parts(parts, &targets, &scoped.replace, report)
    }
}

fn apply_xlsx_scoped_replace_preserving_formatting(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_names: &[String],
    scoped: &ScopedXlsxReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;

    match scoped.scope {
        XlsxTextScope::SharedStrings
        | XlsxTextScope::Headers
        | XlsxTextScope::CellText
        | XlsxTextScope::AllText => {
            matched |= try_apply_preserve_formatting_replace_to_parts(
                parts,
                &["xl/sharedStrings.xml".to_string()],
                &XLSX_SHARED_STRING_ITEM_RE,
                &XML_TEXT_NODE_RE,
                &scoped.replace,
                report,
            );
        }
        _ => {}
    }

    match scoped.scope {
        XlsxTextScope::InlineStrings
        | XlsxTextScope::Headers
        | XlsxTextScope::CellText
        | XlsxTextScope::AllText => {
            let targets = xlsx_inline_string_targets(part_names);
            matched |= try_apply_preserve_formatting_replace_to_parts(
                parts,
                &targets,
                &XLSX_INLINE_STRING_RE,
                &XML_TEXT_NODE_RE,
                &scoped.replace,
                report,
            );
        }
        _ => {}
    }

    match scoped.scope {
        XlsxTextScope::Comments | XlsxTextScope::AllText => {
            let targets = xlsx_scope_targets(part_names, XlsxTextScope::Comments);
            if !targets.is_empty() {
                let normal_comment_targets: Vec<String> = targets
                    .iter()
                    .filter(|name| !name.starts_with("xl/threadedComments/"))
                    .cloned()
                    .collect();
                let threaded_comment_targets: Vec<String> = targets
                    .iter()
                    .filter(|name| name.starts_with("xl/threadedComments/"))
                    .cloned()
                    .collect();

                matched |= try_apply_preserve_formatting_replace_to_parts(
                    parts,
                    &normal_comment_targets,
                    &XLSX_COMMENT_RE,
                    &XML_TEXT_NODE_RE,
                    &scoped.replace,
                    report,
                );

                matched |= try_apply_preserve_formatting_replace_to_parts(
                    parts,
                    &threaded_comment_targets,
                    &XLSX_THREADED_COMMENT_RE,
                    &XML_TEXT_NODE_RE,
                    &scoped.replace,
                    report,
                );
            }
        }
        _ => {}
    }

    if matches!(scoped.scope, XlsxTextScope::AllText) {
        for scope in [XlsxTextScope::CommentAuthors, XlsxTextScope::MetadataAll] {
            let targets = xlsx_scope_targets(part_names, scope);
            matched |= try_apply_simple_replace_to_parts(parts, &targets, &scoped.replace, report);
        }
    }

    if matched {
        Ok(())
    } else {
        match scoped.scope {
            XlsxTextScope::SheetNames => {
                apply_xlsx_sheet_name_replace(parts, &scoped.replace, report)
            }
            XlsxTextScope::MetadataCoreTitle => {
                apply_core_title_replace(parts, &scoped.replace, report)
            }
            _ => {
                let targets = xlsx_scope_targets(part_names, scoped.scope);
                apply_replace_to_parts(parts, &targets, &scoped.replace, report)
            }
        }
    }
}

fn apply_xlsx_workbook_text_replace(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_names: &[String],
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;

    if let Ok(()) = apply_replace_to_xml_text_nodes_named_part(
        parts,
        "xl/sharedStrings.xml",
        &XML_TEXT_NODE_RE,
        replace,
        report,
    ) {
        matched = true;
    }

    let targets = xlsx_inline_string_targets(part_names);
    if !targets.is_empty()
        && apply_replace_to_xml_text_nodes_in_parts(
            parts,
            &targets,
            &XML_TEXT_NODE_RE,
            replace,
            report,
        )
        .is_ok()
    {
        matched = true;
    }

    if matched {
        Ok(())
    } else {
        Err(PatchError::TextNotFound {
            part: "xl/sharedStrings.xml,xl/worksheets/*.xml".to_string(),
            needle: replace.from.clone(),
        })
    }
}

fn apply_xlsx_all_text_replace(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_names: &[String],
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;

    if apply_xlsx_workbook_text_replace(parts, part_names, replace, report).is_ok() {
        matched = true;
    }

    for scope in [
        XlsxTextScope::Comments,
        XlsxTextScope::CommentAuthors,
        XlsxTextScope::MetadataAll,
    ] {
        let targets = xlsx_scope_targets(part_names, scope);
        if !targets.is_empty() && apply_replace_to_parts(parts, &targets, replace, report).is_ok() {
            matched = true;
        }
    }

    if matched {
        Ok(())
    } else {
        Err(PatchError::TextNotFound {
            part: "xl/sharedStrings.xml,xl/worksheets/*.xml,xl/comments*.xml,xl/threadedComments/*.xml,xl/persons/*.xml,docProps/*.xml".to_string(),
            needle: replace.from.clone(),
        })
    }
}

fn apply_xlsx_sheet_name_replace(
    parts: &mut BTreeMap<String, Vec<u8>>,
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    report.parts_scanned += 1;
    let workbook = parts
        .get_mut("xl/workbook.xml")
        .ok_or_else(|| PatchError::MissingPart("xl/workbook.xml".to_string()))?;
    let original = String::from_utf8_lossy(workbook).into_owned();
    let (updated, replacements_applied) = replace_sheet_name_attrs(&original, replace)?;
    if replacements_applied == 0 {
        return Err(PatchError::TextNotFound {
            part: "xl/workbook.xml".to_string(),
            needle: replace.from.clone(),
        });
    }
    *workbook = updated.into_bytes();
    report.parts_modified += 1;
    report.replacements_applied += replacements_applied;
    Ok(())
}

fn rename_exact_sheet_name_attr(
    xml: &str,
    from: &str,
    to: &str,
) -> Result<(String, usize), PatchError> {
    let mut replacements_applied = 0;
    let updated = SHEET_NAME_ATTR_RE.replace_all(xml, |caps: &regex::Captures<'_>| {
        let before = caps.get(1).expect("before").as_str();
        let name = caps.get(2).expect("name").as_str();
        let after = caps.get(3).expect("after").as_str();
        let decoded = xml_unescape(name);
        if decoded == from {
            replacements_applied += 1;
            format!("<sheet{before}name=\"{}\"{after}/>", xml_escape(to))
        } else {
            caps.get(0).expect("whole").as_str().to_string()
        }
    });
    Ok((updated.into_owned(), replacements_applied))
}

fn apply_xlsx_sheet_rename(
    parts: &mut BTreeMap<String, Vec<u8>>,
    rename: &XlsxSheetRename,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    report.parts_scanned += 1;
    let workbook = parts
        .get_mut("xl/workbook.xml")
        .ok_or_else(|| PatchError::MissingPart("xl/workbook.xml".to_string()))?;
    let original = String::from_utf8_lossy(workbook).into_owned();
    let (updated, replacements_applied) =
        rename_exact_sheet_name_attr(&original, &rename.from, &rename.to)?;
    if replacements_applied == 0 {
        return Err(PatchError::TextNotFound {
            part: "xl/workbook.xml".to_string(),
            needle: rename.from.clone(),
        });
    }
    *workbook = updated.into_bytes();
    report.parts_modified += 1;
    report.replacements_applied += replacements_applied;

    if !rename.update_references {
        return Ok(());
    }

    if let Some(workbook) = parts.get_mut("xl/workbook.xml") {
        report.parts_scanned += 1;
        let original = String::from_utf8_lossy(workbook).into_owned();
        let (updated, replacements_applied) =
            rewrite_defined_name_refs(&original, &rename.from, &rename.to)?;
        if replacements_applied > 0 {
            *workbook = updated.into_bytes();
            report.parts_modified += 1;
            report.replacements_applied += replacements_applied;
        }
    }

    let part_names: Vec<String> = parts.keys().cloned().collect();
    let targets = xlsx_formula_targets(&part_names);
    if !targets.is_empty() {
        apply_sheet_ref_rewrite_in_parts(parts, &targets, &rename.from, &rename.to, report)?;
    }
    Ok(())
}

fn apply_sheet_ref_rewrite_in_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_parts: &[String],
    from: &str,
    to: &str,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;
    for part in target_parts {
        report.parts_scanned += 1;
        let Some(data) = parts.get_mut(part) else {
            continue;
        };
        let original = String::from_utf8_lossy(data).into_owned();
        let (updated, replacements_applied) = if part.starts_with("xl/charts/") {
            rewrite_xml_text_nodes(&original, &CHART_FORMULA_NODE_RE, |text| {
                rewrite_formula_sheet_refs(&text, from, to)
            })?
        } else {
            rewrite_xml_text_nodes(&original, &FORMULA_NODE_RE, |text| {
                rewrite_formula_sheet_refs(&text, from, to)
            })?
        };
        if replacements_applied > 0 {
            *data = updated.into_bytes();
            report.parts_modified += 1;
            report.replacements_applied += replacements_applied;
            matched = true;
        }
    }
    let _ = matched;
    Ok(())
}

fn apply_replace_to_xml_text_nodes_named_part(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part: &str,
    node_re: &Regex,
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    report.parts_scanned += 1;
    let data = parts
        .get_mut(part)
        .ok_or_else(|| PatchError::MissingPart(part.to_string()))?;
    let original = String::from_utf8_lossy(data).into_owned();
    let (updated, replacements_applied) = rewrite_xml_text_nodes(&original, node_re, |text| {
        apply_replace_to_text(&text, replace)
    })?;
    if replacements_applied == 0 {
        return Err(PatchError::TextNotFound {
            part: part.to_string(),
            needle: replace.from.clone(),
        });
    }
    *data = updated.into_bytes();
    report.parts_modified += 1;
    report.replacements_applied += replacements_applied;
    Ok(())
}

fn apply_replace_to_xml_text_nodes_in_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_parts: &[String],
    node_re: &Regex,
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let mut matched = false;
    for part in target_parts {
        report.parts_scanned += 1;
        let Some(data) = parts.get_mut(part) else {
            continue;
        };
        let original = String::from_utf8_lossy(data).into_owned();
        let (updated, replacements_applied) = rewrite_xml_text_nodes(&original, node_re, |text| {
            apply_replace_to_text(&text, replace)
        })?;
        if replacements_applied > 0 {
            *data = updated.into_bytes();
            report.parts_modified += 1;
            report.replacements_applied += replacements_applied;
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

fn rewrite_xml_text_nodes<F>(
    xml: &str,
    node_re: &Regex,
    mut rewrite: F,
) -> Result<(String, usize), PatchError>
where
    F: FnMut(String) -> Result<(String, usize), PatchError>,
{
    let mut replacements_applied = 0;
    let updated = node_re.replace_all(xml, |caps: &regex::Captures<'_>| {
        let whole = caps.get(0).expect("whole match").as_str();
        let inner = caps.get(1).expect("inner match").as_str();
        let decoded = xml_unescape(inner);
        match rewrite(decoded) {
            Ok((rewritten, applied)) => {
                replacements_applied += applied;
                whole.replacen(inner, &xml_escape(&rewritten), 1)
            }
            Err(_) => whole.to_string(),
        }
    });
    Ok((updated.into_owned(), replacements_applied))
}

#[derive(Debug, Clone)]
struct XmlTextNodeRef {
    inner_start: usize,
    inner_end: usize,
    text: String,
}

fn rewrite_xml_text_containers_preserving_runs(
    xml: &str,
    container_re: &Regex,
    node_re: &Regex,
    replace: &TextReplace,
) -> Result<(String, usize), PatchError> {
    let mut replacements_applied = 0;
    let updated = container_re.replace_all(xml, |caps: &regex::Captures<'_>| {
        let whole = caps.get(0).expect("whole match").as_str();
        if matches!(replace.mode, ReplaceMode::First) && replacements_applied > 0 {
            return whole.to_string();
        }
        match rewrite_xml_text_nodes_preserving_runs(whole, node_re, replace) {
            Ok((rewritten, applied)) => {
                replacements_applied += applied;
                rewritten
            }
            Err(_) => whole.to_string(),
        }
    });
    Ok((updated.into_owned(), replacements_applied))
}

fn rewrite_xml_text_nodes_preserving_runs(
    xml: &str,
    node_re: &Regex,
    replace: &TextReplace,
) -> Result<(String, usize), PatchError> {
    let mut nodes = Vec::new();
    for caps in node_re.captures_iter(xml) {
        let inner = caps.get(1).expect("inner match");
        nodes.push(XmlTextNodeRef {
            inner_start: inner.start(),
            inner_end: inner.end(),
            text: xml_unescape(inner.as_str()),
        });
    }
    if nodes.is_empty() {
        return Ok((xml.to_string(), 0));
    }

    let mut flat = String::new();
    let mut node_ranges = Vec::with_capacity(nodes.len());
    for node in &nodes {
        let start = flat.len();
        flat.push_str(&node.text);
        node_ranges.push((start, flat.len()));
    }

    let matches = find_match_ranges(&flat, replace)?;
    if matches.is_empty() {
        return Ok((xml.to_string(), 0));
    }

    let mut updated_node_texts: Vec<String> = nodes.iter().map(|node| node.text.clone()).collect();
    for (match_start, match_end) in matches.iter().copied().rev() {
        let first_idx = node_ranges
            .iter()
            .position(|(_, end)| *end > match_start)
            .expect("first matched node");
        let last_idx = node_ranges
            .iter()
            .rposition(|(start, _)| *start < match_end)
            .expect("last matched node");

        let first_local_start = match_start.saturating_sub(node_ranges[first_idx].0);
        let last_local_end = match_end.saturating_sub(node_ranges[last_idx].0);

        if first_idx == last_idx {
            let current = &updated_node_texts[first_idx];
            updated_node_texts[first_idx] = format!(
                "{}{}{}",
                &current[..first_local_start],
                replace.to,
                &current[last_local_end..]
            );
            continue;
        }

        let first_prefix = updated_node_texts[first_idx][..first_local_start].to_string();
        updated_node_texts[first_idx] = format!("{first_prefix}{}", replace.to);

        let last_suffix = updated_node_texts[last_idx][last_local_end..].to_string();
        updated_node_texts[last_idx] = last_suffix;
        for text in &mut updated_node_texts[first_idx + 1..last_idx] {
            text.clear();
        }
    }

    let mut rewritten = String::with_capacity(xml.len());
    let mut cursor = 0;
    for (node, new_text) in nodes.iter().zip(updated_node_texts.iter()) {
        rewritten.push_str(&xml[cursor..node.inner_start]);
        rewritten.push_str(&xml_escape(new_text));
        cursor = node.inner_end;
    }
    rewritten.push_str(&xml[cursor..]);
    Ok((rewritten, matches.len()))
}

fn replace_sheet_name_attrs(
    xml: &str,
    replace: &TextReplace,
) -> Result<(String, usize), PatchError> {
    let mut replacements_applied = 0;
    let updated = SHEET_NAME_ATTR_RE.replace_all(xml, |caps: &regex::Captures<'_>| {
        let before = caps.get(1).expect("before").as_str();
        let name = caps.get(2).expect("name").as_str();
        let after = caps.get(3).expect("after").as_str();
        let decoded = xml_unescape(name);
        match apply_replace_to_text(&decoded, replace) {
            Ok((rewritten, applied)) if applied > 0 => {
                replacements_applied += applied;
                format!("<sheet{before}name=\"{}\"{after}/>", xml_escape(&rewritten))
            }
            _ => caps.get(0).expect("whole").as_str().to_string(),
        }
    });
    Ok((updated.into_owned(), replacements_applied))
}

fn rewrite_defined_name_refs(
    xml: &str,
    from: &str,
    to: &str,
) -> Result<(String, usize), PatchError> {
    rewrite_xml_text_nodes(xml, &DEFINED_NAME_RE, |text| {
        rewrite_formula_sheet_refs(&text, from, to)
    })
}

fn rewrite_formula_sheet_refs(
    formula: &str,
    from: &str,
    to: &str,
) -> Result<(String, usize), PatchError> {
    let mut replacements_applied = 0;
    let mut updated = QUOTED_SHEET_REF_RE
        .replace_all(formula, |caps: &regex::Captures<'_>| {
            let quoted_name = caps.get(1).expect("quoted").as_str();
            let decoded = quoted_name.replace("''", "'");
            if decoded == from {
                replacements_applied += 1;
                format!("{}!", excel_sheet_ref(to))
            } else {
                caps.get(0).expect("whole").as_str().to_string()
            }
        })
        .into_owned();

    updated = UNQUOTED_SHEET_REF_RE
        .replace_all(&updated, |caps: &regex::Captures<'_>| {
            let name = caps.get(1).expect("name").as_str();
            if name == from {
                replacements_applied += 1;
                format!("{}!", excel_sheet_ref(to))
            } else {
                caps.get(0).expect("whole").as_str().to_string()
            }
        })
        .into_owned();

    Ok((updated, replacements_applied))
}

fn excel_sheet_ref(name: &str) -> String {
    if is_valid_unquoted_sheet_name(name) {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn is_valid_unquoted_sheet_name(name: &str) -> bool {
    let mut chars = name.chars();
    match chars.next() {
        Some(ch) if ch.is_ascii_alphabetic() || ch == '_' => {}
        _ => return false,
    }
    chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '_' || ch == '.')
}

fn xml_unescape(text: &str) -> String {
    text.replace("&apos;", "'")
        .replace("&quot;", "\"")
        .replace("&gt;", ">")
        .replace("&lt;", "<")
        .replace("&amp;", "&")
}

fn apply_core_title(
    parts: &mut BTreeMap<String, Vec<u8>>,
    title: Option<&str>,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    let Some(title) = title else {
        return Ok(());
    };

    report.parts_scanned += 1;
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
    report.parts_modified += 1;
    report.replacements_applied += 1;
    Ok(())
}

fn apply_core_title_replace(
    parts: &mut BTreeMap<String, Vec<u8>>,
    replace: &TextReplace,
    report: &mut PatchReport,
) -> Result<(), PatchError> {
    apply_replace_to_named_part(parts, "docProps/core.xml", replace, report)
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
            (
                "word/comments.xml",
                "<w:comment w:author=\"Alice\"><w:t>word comment</w:t></w:comment>",
            ),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title>old word</dc:title></cp:coreProperties>",
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>word company</Company></Properties>",
            ),
            (
                "docProps/custom.xml",
                "<Properties><property name=\"FilePath\"><vt:lpwstr>word.docx</vt:lpwstr></property></Properties>",
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
        assert!(
            String::from_utf8_lossy(&parts["word/comments.xml"]).contains("w:author=\"Alice\"")
        );
        assert!(String::from_utf8_lossy(&parts["docProps/core.xml"]).contains("new core"));
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("term company"));
        assert!(String::from_utf8_lossy(&parts["docProps/custom.xml"]).contains("term.docx"));
    }

    #[test]
    fn typed_docx_patch_replaces_comment_author_and_metadata_properties() {
        let bytes = build_zip(vec![
            (
                "word/comments.xml",
                "<w:comments><w:comment w:id=\"0\" w:author=\"Alice\" w:initials=\"AL\"><w:p><w:r><w:t>Needs review</w:t></w:r></w:p></w:comment></w:comments>",
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>Old Company</Company><Template>Old Template</Template></Properties>",
            ),
            (
                "docProps/custom.xml",
                "<Properties><property name=\"FilePath\"><vt:lpwstr>/tmp/old.docx</vt:lpwstr></property></Properties>",
            ),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title>old</dc:title></cp:coreProperties>",
            ),
        ]);

        let patched = patch_docx(
            &bytes,
            &DocxPatch {
                set_core_title: None,
                replace_body_title: None,
                scoped_replacements: vec![
                    ScopedDocxReplace {
                        scope: DocxTextScope::Comments,
                        replace: TextReplace::all("Alice", "Bob"),
                    },
                    ScopedDocxReplace {
                        scope: DocxTextScope::MetadataApp,
                        replace: TextReplace::all("Old", "New"),
                    },
                    ScopedDocxReplace {
                        scope: DocxTextScope::MetadataCustom,
                        replace: TextReplace::all("/tmp/old.docx", "/tmp/new.docx"),
                    },
                ],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(String::from_utf8_lossy(&parts["word/comments.xml"]).contains("w:author=\"Bob\""));
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("New Company"));
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("New Template"));
        assert!(String::from_utf8_lossy(&parts["docProps/custom.xml"]).contains("/tmp/new.docx"));
    }

    #[test]
    fn typed_pptx_patch_replaces_comments_and_notes() {
        let bytes = build_zip(vec![
            ("ppt/slides/slide1.xml", "<a:t>word slide</a:t>"),
            ("ppt/notesSlides/notesSlide1.xml", "<a:t>word notes</a:t>"),
            (
                "ppt/comments/comment1.xml",
                "<p:cm authorId=\"0\"><p:text>word comment</p:text></p:cm>",
            ),
            (
                "ppt/commentAuthors.xml",
                "<p:cmAuthorLst><p:cmAuthor id=\"0\" name=\"word author\" initials=\"WA\"/></p:cmAuthorLst>",
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>word company</Company></Properties>",
            ),
            (
                "docProps/custom.xml",
                "<Properties><property name=\"FileName\"><vt:lpwstr>word.pptx</vt:lpwstr></property></Properties>",
            ),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title>word deck</dc:title></cp:coreProperties>",
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
        assert!(String::from_utf8_lossy(&parts["ppt/commentAuthors.xml"]).contains("term author"));
        assert!(String::from_utf8_lossy(&parts["docProps/core.xml"]).contains("deck title"));
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("term company"));
        assert!(String::from_utf8_lossy(&parts["docProps/custom.xml"]).contains("term.pptx"));
    }

    #[test]
    fn typed_pptx_patch_replaces_comment_authors_and_metadata_properties() {
        let bytes = build_zip(vec![
            (
                "ppt/commentAuthors.xml",
                "<p:cmAuthorLst><p:cmAuthor id=\"0\" name=\"Alice\" initials=\"AL\"/></p:cmAuthorLst>",
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>Old Company</Company><PresentationFormat>Old Deck</PresentationFormat></Properties>",
            ),
            (
                "docProps/custom.xml",
                "<Properties><property name=\"FileName\"><vt:lpwstr>old.pptx</vt:lpwstr></property></Properties>",
            ),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title>old</dc:title></cp:coreProperties>",
            ),
        ]);

        let patched = patch_pptx(
            &bytes,
            &PptxPatch {
                set_core_title: None,
                scoped_replacements: vec![
                    ScopedPptxReplace {
                        scope: PptxTextScope::CommentAuthors,
                        replace: TextReplace::all("Alice", "Bob"),
                    },
                    ScopedPptxReplace {
                        scope: PptxTextScope::MetadataApp,
                        replace: TextReplace::all("Old", "New"),
                    },
                    ScopedPptxReplace {
                        scope: PptxTextScope::MetadataCustom,
                        replace: TextReplace::all("old.pptx", "new.pptx"),
                    },
                ],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(String::from_utf8_lossy(&parts["ppt/commentAuthors.xml"]).contains("name=\"Bob\""));
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("New Company"));
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("New Deck"));
        assert!(String::from_utf8_lossy(&parts["docProps/custom.xml"]).contains("new.pptx"));
    }

    #[test]
    fn case_insensitive_first_replace_works() {
        let (updated, replacements_applied) = apply_replace_to_text(
            "Word word WORD",
            &TextReplace {
                from: "word".to_string(),
                to: "term".to_string(),
                mode: ReplaceMode::First,
                match_policy: MatchPolicy::CaseInsensitive,
                preserve_formatting: false,
            },
        )
        .unwrap();
        assert_eq!(updated, "term word WORD");
        assert_eq!(replacements_applied, 1);
    }

    #[test]
    fn typed_docx_patch_preserve_formatting_cross_run_first_style_wins() {
        let bytes = build_zip(vec![(
            "word/document.xml",
            concat!(
                "<w:p>",
                "<w:r><w:rPr><w:b/></w:rPr><w:t>Hel</w:t></w:r>",
                "<w:r><w:rPr><w:i/></w:rPr><w:t>lo</w:t></w:r>",
                "</w:p>"
            ),
        )]);

        let patched = patch_docx(
            &bytes,
            &DocxPatch {
                set_core_title: None,
                replace_body_title: None,
                scoped_replacements: vec![ScopedDocxReplace {
                    scope: DocxTextScope::Body,
                    replace: TextReplace::all("Hello", "Hi").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(String::from_utf8_lossy(&parts["word/document.xml"]).contains(
            "<w:r><w:rPr><w:b/></w:rPr><w:t>Hi</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t></w:t></w:r>"
        ));
    }

    #[test]
    fn typed_pptx_patch_preserve_formatting_cross_run_first_style_wins() {
        let bytes = build_zip(vec![(
            "ppt/slides/slide1.xml",
            concat!(
                "<a:p>",
                "<a:r><a:rPr b=\"1\"/><a:t>Hel</a:t></a:r>",
                "<a:r><a:rPr i=\"1\"/><a:t>lo</a:t></a:r>",
                "</a:p>"
            ),
        )]);

        let patched = patch_pptx(
            &bytes,
            &PptxPatch {
                set_core_title: None,
                scoped_replacements: vec![ScopedPptxReplace {
                    scope: PptxTextScope::SlideBody,
                    replace: TextReplace::all("Hello", "Hi").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(
            String::from_utf8_lossy(&parts["ppt/slides/slide1.xml"]).contains(
                "<a:r><a:rPr b=\"1\"/><a:t>Hi</a:t></a:r><a:r><a:rPr i=\"1\"/><a:t></a:t></a:r>"
            )
        );
    }

    #[test]
    fn typed_docx_patch_preserve_formatting_does_not_cross_paragraphs() {
        let bytes = build_zip(vec![(
            "word/document.xml",
            concat!(
                "<w:p><w:r><w:t>Hel</w:t></w:r></w:p>",
                "<w:p><w:r><w:t>lo</w:t></w:r></w:p>"
            ),
        )]);

        let err = patch_docx(
            &bytes,
            &DocxPatch {
                set_core_title: None,
                replace_body_title: None,
                scoped_replacements: vec![ScopedDocxReplace {
                    scope: DocxTextScope::Body,
                    replace: TextReplace::all("Hello", "Hi").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap_err();

        assert!(matches!(err, PatchError::TextNotFound { .. }));
    }

    #[test]
    fn typed_docx_patch_preserve_formatting_first_stops_after_first_paragraph() {
        let bytes = build_zip(vec![(
            "word/document.xml",
            concat!(
                "<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>",
                "<w:p><w:r><w:t>Hello again</w:t></w:r></w:p>"
            ),
        )]);

        let patched = patch_docx(
            &bytes,
            &DocxPatch {
                set_core_title: None,
                replace_body_title: None,
                scoped_replacements: vec![ScopedDocxReplace {
                    scope: DocxTextScope::Body,
                    replace: TextReplace::first("Hello", "Hi").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        let document_xml = String::from_utf8_lossy(&parts["word/document.xml"]);
        assert!(document_xml.contains("<w:p><w:r><w:t>Hi world</w:t></w:r></w:p>"));
        assert!(document_xml.contains("<w:p><w:r><w:t>Hello again</w:t></w:r></w:p>"));
    }

    #[test]
    fn typed_pptx_patch_preserve_formatting_does_not_cross_paragraphs() {
        let bytes = build_zip(vec![(
            "ppt/slides/slide1.xml",
            concat!(
                "<p:sld xmlns:p=\"p\" xmlns:a=\"a\"><p:cSld><p:spTree><p:sp><p:txBody>",
                "<a:p><a:r><a:t>Hel</a:t></a:r></a:p>",
                "<a:p><a:r><a:t>lo</a:t></a:r></a:p>",
                "</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"
            ),
        )]);

        let err = patch_pptx(
            &bytes,
            &PptxPatch {
                set_core_title: None,
                scoped_replacements: vec![ScopedPptxReplace {
                    scope: PptxTextScope::SlideBody,
                    replace: TextReplace::all("Hello", "Hi").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap_err();

        assert!(matches!(err, PatchError::TextNotFound { .. }));
    }

    #[test]
    fn typed_xlsx_patch_preserve_formatting_rich_text_and_empty_string() {
        let bytes = build_zip(vec![
            (
                "xl/sharedStrings.xml",
                concat!(
                    "<sst><si>",
                    "<r><rPr><b/></rPr><t>Confi</t></r>",
                    "<r><rPr><i/></rPr><t>dential</t></r>",
                    "</si></sst>"
                ),
            ),
            (
                "xl/workbook.xml",
                "<workbook><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            ),
        ]);

        let patched = patch_xlsx(
            &bytes,
            &XlsxPatch {
                set_core_title: None,
                rename_sheets: vec![],
                scoped_replacements: vec![ScopedXlsxReplace {
                    scope: XlsxTextScope::SharedStrings,
                    replace: TextReplace::all("Confidential", "").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(
            String::from_utf8_lossy(&parts["xl/sharedStrings.xml"])
                .contains("<r><rPr><b/></rPr><t></t></r><r><rPr><i/></rPr><t></t></r>")
        );
    }

    #[test]
    fn typed_xlsx_patch_preserve_formatting_does_not_cross_string_items() {
        let bytes = build_zip(vec![
            (
                "xl/sharedStrings.xml",
                "<sst><si><t>Hel</t></si><si><t>lo</t></si></sst>",
            ),
            (
                "xl/workbook.xml",
                "<workbook><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            ),
        ]);

        let err = patch_xlsx(
            &bytes,
            &XlsxPatch {
                set_core_title: None,
                rename_sheets: vec![],
                scoped_replacements: vec![ScopedXlsxReplace {
                    scope: XlsxTextScope::SharedStrings,
                    replace: TextReplace::all("Hello", "Hi").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap_err();

        assert!(matches!(err, PatchError::TextNotFound { .. }));
    }

    #[test]
    fn typed_metadata_replace_ignores_preserve_formatting_flag() {
        let bytes = build_zip(vec![(
            "docProps/app.xml",
            "<Properties><Company>Old Company</Company></Properties>",
        )]);

        let patched = patch_docx(
            &bytes,
            &DocxPatch {
                set_core_title: None,
                replace_body_title: None,
                scoped_replacements: vec![ScopedDocxReplace {
                    scope: DocxTextScope::MetadataApp,
                    replace: TextReplace::all("Old", "New").with_preserve_formatting(true),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched).unwrap();
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("New Company"));
    }

    #[test]
    fn patch_docx_batch_works() {
        let bytes = build_zip(vec![("word/document.xml", "<w:t>word</w:t>")]);
        let patched = patch_docx_batch_with_report(
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
            let parts = read_parts(&item.content).unwrap();
            assert!(String::from_utf8_lossy(&parts["word/document.xml"]).contains("term"));
            assert_eq!(item.report.parts_scanned, 1);
            assert_eq!(item.report.parts_modified, 1);
            assert_eq!(item.report.replacements_applied, 1);
        }
    }

    #[test]
    fn typed_xlsx_patch_renames_sheet_and_updates_references() {
        let bytes = build_zip(vec![
            (
                "xl/workbook.xml",
                concat!(
                    r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                    r#"<sheets><sheet name="Sales Data" sheetId="1" r:id="rId1"/>"#,
                    r#"<sheet name="Summary" sheetId="2" r:id="rId2"/></sheets>"#,
                    r#"<definedNames><definedName name="MyRange">'Sales Data'!$A$1:$B$2</definedName></definedNames>"#,
                    r#"</workbook>"#,
                ),
            ),
            (
                "xl/worksheets/sheet1.xml",
                "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Revenue</t></is></c></row></sheetData></worksheet>",
            ),
            (
                "xl/worksheets/sheet2.xml",
                "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f>'Sales Data'!B2</f><v>10</v></c></row></sheetData></worksheet>",
            ),
            (
                "xl/charts/chart1.xml",
                "<c:chartSpace xmlns:c=\"c\"><c:lineChart><c:ser><c:val><c:numRef><c:f>'Sales Data'!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:lineChart></c:chartSpace>",
            ),
        ]);

        let patched = patch_xlsx_with_report(
            &bytes,
            &XlsxPatch {
                set_core_title: None,
                rename_sheets: vec![XlsxSheetRename {
                    from: "Sales Data".to_string(),
                    to: "Revenue Data".to_string(),
                    update_references: true,
                }],
                scoped_replacements: vec![],
            },
        )
        .unwrap();

        let parts = read_parts(&patched.content).unwrap();
        let workbook_xml = String::from_utf8_lossy(&parts["xl/workbook.xml"]);
        let sheet_xml = String::from_utf8_lossy(&parts["xl/worksheets/sheet2.xml"]);
        let chart_xml = String::from_utf8_lossy(&parts["xl/charts/chart1.xml"]);
        assert!(workbook_xml.contains("name=\"Revenue Data\""));
        assert!(
            workbook_xml.contains("'Revenue Data'!$A$1:$B$2")
                || workbook_xml.contains("&apos;Revenue Data&apos;!$A$1:$B$2")
        );
        assert!(
            sheet_xml.contains("'Revenue Data'!B2")
                || sheet_xml.contains("&apos;Revenue Data&apos;!B2")
        );
        assert!(
            chart_xml.contains("'Revenue Data'!$B$2:$B$3")
                || chart_xml.contains("&apos;Revenue Data&apos;!$B$2:$B$3")
        );
        assert!(patched.report.replacements_applied >= 4);
    }

    #[test]
    fn xlsx_sheet_rename_requires_exact_sheet_name_match() {
        let bytes = build_zip(vec![(
            "xl/workbook.xml",
            concat!(
                "<workbook><definedNames>",
                "<definedName name=\"_xlnm.Print_Area\">'Sales Data'!$A$1</definedName>",
                "</definedNames><sheets>",
                "<sheet name=\"Sales Data\" sheetId=\"1\" r:id=\"rId1\"/>",
                "</sheets></workbook>"
            ),
        )]);

        let err = patch_xlsx_with_report(
            &bytes,
            &XlsxPatch {
                set_core_title: None,
                rename_sheets: vec![XlsxSheetRename {
                    from: "Sales".to_string(),
                    to: "Revenue".to_string(),
                    update_references: true,
                }],
                scoped_replacements: vec![],
            },
        )
        .unwrap_err();

        assert!(matches!(err, PatchError::TextNotFound { .. }));
    }

    #[test]
    fn typed_xlsx_patch_replaces_shared_strings_text() {
        let bytes = build_zip(vec![
            (
                "xl/sharedStrings.xml",
                "<sst><si><t>word alpha</t></si><si><t>word beta</t></si></sst>",
            ),
            (
                "xl/workbook.xml",
                "<workbook><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            ),
        ]);

        let patched = patch_xlsx_with_report(
            &bytes,
            &XlsxPatch {
                set_core_title: None,
                rename_sheets: vec![],
                scoped_replacements: vec![ScopedXlsxReplace {
                    scope: XlsxTextScope::AllText,
                    replace: TextReplace::all("word", "term"),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched.content).unwrap();
        assert!(String::from_utf8_lossy(&parts["xl/sharedStrings.xml"]).contains("term alpha"));
        assert!(String::from_utf8_lossy(&parts["xl/sharedStrings.xml"]).contains("term beta"));
        assert!(patched.report.replacements_applied >= 2);
    }

    #[test]
    fn typed_xlsx_patch_replaces_comments_authors_and_metadata_via_all_text() {
        let bytes = build_zip(vec![
            (
                "xl/sharedStrings.xml",
                "<sst><si><t>word cell</t></si></sst>",
            ),
            (
                "xl/comments1.xml",
                "<comments><authors><author>Alice</author></authors><commentList><comment ref=\"A1\" authorId=\"0\"><text><r><t>word comment</t></r></text></comment></commentList></comments>",
            ),
            (
                "xl/persons/person.xml",
                "<personList><person displayName=\"Alice\" id=\"{1}\" userId=\"alice@example.com\" providerId=\"Alice\"/></personList>",
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>word company</Company></Properties>",
            ),
            (
                "docProps/custom.xml",
                "<Properties><property name=\"FilePath\"><vt:lpwstr>word.xlsx</vt:lpwstr></property></Properties>",
            ),
            (
                "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"cp\" xmlns:dc=\"dc\"><dc:title>word workbook</dc:title></cp:coreProperties>",
            ),
            (
                "xl/workbook.xml",
                "<workbook><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            ),
        ]);

        let patched = patch_xlsx_with_report(
            &bytes,
            &XlsxPatch {
                set_core_title: None,
                rename_sheets: vec![],
                scoped_replacements: vec![ScopedXlsxReplace {
                    scope: XlsxTextScope::AllText,
                    replace: TextReplace::all("word", "term"),
                }],
            },
        )
        .unwrap();

        let parts = read_parts(&patched.content).unwrap();
        assert!(String::from_utf8_lossy(&parts["xl/sharedStrings.xml"]).contains("term cell"));
        assert!(String::from_utf8_lossy(&parts["xl/comments1.xml"]).contains("term comment"));
        assert!(
            String::from_utf8_lossy(&parts["xl/persons/person.xml"])
                .contains("displayName=\"Alice\"")
        );
        assert!(String::from_utf8_lossy(&parts["docProps/app.xml"]).contains("term company"));
        assert!(String::from_utf8_lossy(&parts["docProps/custom.xml"]).contains("term.xlsx"));
        assert!(String::from_utf8_lossy(&parts["docProps/core.xml"]).contains("term workbook"));
        assert!(patched.report.replacements_applied >= 5);
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

        let xlsx_patch = XlsxPatch {
            set_core_title: Some("Workbook".to_string()),
            rename_sheets: vec![XlsxSheetRename {
                from: "Sales".to_string(),
                to: "Revenue".to_string(),
                update_references: true,
            }],
            scoped_replacements: vec![ScopedXlsxReplace {
                scope: XlsxTextScope::AllText,
                replace: TextReplace::all("word", "term"),
            }],
        };
        let json = serde_json::to_string(&xlsx_patch).unwrap();
        let decoded: XlsxPatch = serde_json::from_str(&json).unwrap();
        assert_eq!(decoded, xlsx_patch);
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
