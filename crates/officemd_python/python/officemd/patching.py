from __future__ import annotations

from dataclasses import asdict, dataclass, field, is_dataclass
from enum import Enum
import json
from pathlib import Path
from typing import Any, Mapping

from officemd._officemd import _patch_docx_batch_json  # type: ignore[unresolved-import]
from officemd._officemd import _patch_docx_batch_json_with_report  # type: ignore[unresolved-import]
from officemd._officemd import _patch_docx_json  # type: ignore[unresolved-import]
from officemd._officemd import _patch_docx_json_with_report  # type: ignore[unresolved-import]
from officemd._officemd import _patch_pptx_batch_json  # type: ignore[unresolved-import]
from officemd._officemd import _patch_pptx_batch_json_with_report  # type: ignore[unresolved-import]
from officemd._officemd import _patch_pptx_json  # type: ignore[unresolved-import]
from officemd._officemd import _patch_pptx_json_with_report  # type: ignore[unresolved-import]


class ReplaceMode(str, Enum):
    FIRST = "first"
    ALL = "all"


class MatchPolicy(str, Enum):
    EXACT = "exact"
    CASE_INSENSITIVE = "case_insensitive"
    WHOLE_WORD = "whole_word"
    WHOLE_WORD_CASE_INSENSITIVE = "whole_word_case_insensitive"


class DocxTextScope(str, Enum):
    BODY = "body"
    HEADERS = "headers"
    FOOTERS = "footers"
    COMMENTS = "comments"
    FOOTNOTES = "footnotes"
    ENDNOTES = "endnotes"
    METADATA_CORE_TITLE = "metadata_core_title"
    ALL_TEXT = "all_text"


class PptxTextScope(str, Enum):
    SLIDE_TITLES = "slide_titles"
    SLIDE_BODY = "slide_body"
    NOTES = "notes"
    COMMENTS = "comments"
    METADATA_CORE_TITLE = "metadata_core_title"
    ALL_TEXT = "all_text"


@dataclass(frozen=True)
class TextReplace:
    from_text: str
    to_text: str
    mode: ReplaceMode = ReplaceMode.ALL
    match_policy: MatchPolicy = MatchPolicy.EXACT

    def to_dict(self) -> dict[str, Any]:
        return {
            "from": self.from_text,
            "to": self.to_text,
            "mode": self.mode.value,
            "match_policy": self.match_policy.value,
        }


@dataclass(frozen=True)
class ScopedDocxReplace:
    scope: DocxTextScope
    replace: TextReplace

    def to_dict(self) -> dict[str, Any]:
        return {"scope": self.scope.value, "replace": self.replace.to_dict()}


@dataclass(frozen=True)
class ScopedPptxReplace:
    scope: PptxTextScope
    replace: TextReplace

    def to_dict(self) -> dict[str, Any]:
        return {"scope": self.scope.value, "replace": self.replace.to_dict()}


@dataclass(frozen=True)
class DocxPatch:
    set_core_title: str | None = None
    replace_body_title: TextReplace | None = None
    scoped_replacements: list[ScopedDocxReplace] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {}
        if self.set_core_title is not None:
            payload["set_core_title"] = self.set_core_title
        if self.replace_body_title is not None:
            payload["replace_body_title"] = self.replace_body_title.to_dict()
        if self.scoped_replacements:
            payload["scoped_replacements"] = [item.to_dict() for item in self.scoped_replacements]
        return payload


@dataclass(frozen=True)
class PptxPatch:
    set_core_title: str | None = None
    scoped_replacements: list[ScopedPptxReplace] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {}
        if self.set_core_title is not None:
            payload["set_core_title"] = self.set_core_title
        if self.scoped_replacements:
            payload["scoped_replacements"] = [item.to_dict() for item in self.scoped_replacements]
        return payload


@dataclass(frozen=True)
class BatchPatchJob:
    input_path: str | Path
    output_path: str | Path
    patch: DocxPatch | PptxPatch | Mapping[str, Any]
    format: str


@dataclass(frozen=True)
class PatchReport:
    parts_scanned: int
    parts_modified: int
    replacements_applied: int


@dataclass(frozen=True)
class PatchContentResult:
    content: bytes
    report: PatchReport


@dataclass(frozen=True)
class BatchPatchContentResult:
    content: bytes
    report: PatchReport


@dataclass(frozen=True)
class BatchPatchResult:
    input_path: str
    output_path: str
    format: str
    ok: bool
    error: str | None = None
    report: PatchReport | None = None


def _normalize_payload(value: Any) -> Any:
    if hasattr(value, "to_dict"):
        return value.to_dict()
    if isinstance(value, Enum):
        return value.value
    if is_dataclass(value):
        return {k: _normalize_payload(v) for k, v in asdict(value).items()}
    if isinstance(value, Mapping):
        return {str(k): _normalize_payload(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_normalize_payload(v) for v in value]
    return value


def _to_patch_json(patch: DocxPatch | PptxPatch | Mapping[str, Any]) -> str:
    return json.dumps(_normalize_payload(patch), ensure_ascii=False)


def patch_docx(content: bytes, patch: DocxPatch | Mapping[str, Any]) -> bytes:
    return _patch_docx_json(content, _to_patch_json(patch))


def patch_docx_with_report(
    content: bytes, patch: DocxPatch | Mapping[str, Any]
) -> PatchContentResult:
    item = json.loads(_patch_docx_json_with_report(content, _to_patch_json(patch)))
    return PatchContentResult(content=bytes(item["content"]), report=PatchReport(**item["report"]))


def patch_pptx(content: bytes, patch: PptxPatch | Mapping[str, Any]) -> bytes:
    return _patch_pptx_json(content, _to_patch_json(patch))


def patch_pptx_with_report(
    content: bytes, patch: PptxPatch | Mapping[str, Any]
) -> PatchContentResult:
    item = json.loads(_patch_pptx_json_with_report(content, _to_patch_json(patch)))
    return PatchContentResult(content=bytes(item["content"]), report=PatchReport(**item["report"]))


def patch_docx_batch(
    contents: list[bytes], patch: DocxPatch | Mapping[str, Any], workers: int | None = None
) -> list[bytes]:
    return _patch_docx_batch_json(contents, _to_patch_json(patch), workers)


def patch_pptx_batch(
    contents: list[bytes], patch: PptxPatch | Mapping[str, Any], workers: int | None = None
) -> list[bytes]:
    return _patch_pptx_batch_json(contents, _to_patch_json(patch), workers)


def _batch_with_report_from_json(payload: str) -> list[BatchPatchContentResult]:
    raw_items = json.loads(payload)
    return [
        BatchPatchContentResult(
            content=bytes(item["content"]),
            report=PatchReport(**item["report"]),
        )
        for item in raw_items
    ]


def patch_docx_batch_with_report(
    contents: list[bytes], patch: DocxPatch | Mapping[str, Any], workers: int | None = None
) -> list[BatchPatchContentResult]:
    payload = _patch_docx_batch_json_with_report(contents, _to_patch_json(patch), workers)
    return _batch_with_report_from_json(payload)


def patch_pptx_batch_with_report(
    contents: list[bytes], patch: PptxPatch | Mapping[str, Any], workers: int | None = None
) -> list[BatchPatchContentResult]:
    payload = _patch_pptx_batch_json_with_report(contents, _to_patch_json(patch), workers)
    return _batch_with_report_from_json(payload)


def patch_files(jobs: list[BatchPatchJob], workers: int | None = None) -> list[BatchPatchResult]:
    results: list[BatchPatchResult] = []
    grouped: dict[tuple[str, str], list[BatchPatchJob]] = {}

    for job in jobs:
        if job.format not in {"docx", "pptx"}:
            results.append(
                BatchPatchResult(
                    str(job.input_path),
                    str(job.output_path),
                    job.format,
                    False,
                    "format must be 'docx' or 'pptx'",
                )
            )
            continue
        patch_json = _to_patch_json(job.patch)
        grouped.setdefault((job.format, patch_json), []).append(job)

    for (fmt, patch_json), grouped_jobs in grouped.items():
        contents = [Path(job.input_path).read_bytes() for job in grouped_jobs]
        if fmt == "docx":
            patched_contents = _batch_with_report_from_json(
                _patch_docx_batch_json_with_report(contents, patch_json, workers)
            )
        else:
            patched_contents = _batch_with_report_from_json(
                _patch_pptx_batch_json_with_report(contents, patch_json, workers)
            )

        for job, patched in zip(grouped_jobs, patched_contents, strict=True):
            output_path = Path(job.output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_bytes(patched.content)
            results.append(
                BatchPatchResult(
                    str(job.input_path),
                    str(job.output_path),
                    job.format,
                    True,
                    report=patched.report,
                )
            )

    return results
