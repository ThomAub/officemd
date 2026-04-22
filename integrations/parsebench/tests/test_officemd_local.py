"""Unit tests for the `officemd_local` ParseBench provider.

These tests stub subprocess so they run without invoking `cargo` or
ParseBench-driven execution. They assume `parse_bench` is installed in the
test environment; when it is not, the tests are skipped.
"""

from __future__ import annotations

import json
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

import pytest

pytest.importorskip("parse_bench")

from parse_bench.inference.providers.base import (  # noqa: E402
    ProviderConfigError,
    ProviderPermanentError,
)
from parse_bench.schemas.pipeline import PipelineSpec  # noqa: E402
from parse_bench.schemas.pipeline_io import InferenceRequest  # noqa: E402
from parse_bench.schemas.product import ProductType  # noqa: E402

from parse_bench.inference.providers.parse.officemd_local import (  # noqa: E402
    OfficeMDLocalProvider,
)


@pytest.fixture()
def repo_root(tmp_path: Path) -> Path:
    root = tmp_path / "officemd"
    root.mkdir()
    (root / "Cargo.toml").write_text("[workspace]\n")
    return root


@pytest.fixture()
def pdf_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.pdf"
    path.write_bytes(b"%PDF-1.4\n%stub\n")
    return path


def _pipeline(repo_root: Path) -> PipelineSpec:
    return PipelineSpec(
        pipeline_name="officemd_local",
        provider_name="officemd_local",
        product_type=ProductType.PARSE,
        config={
            "cargo_run": True,
            "repo_root": str(repo_root),
            "cargo_profile": "dev",
        },
    )


def _request(pdf_path: Path) -> InferenceRequest:
    return InferenceRequest(
        example_id="ex-1",
        source_file_path=str(pdf_path),
        product_type=ProductType.PARSE,
    )


def _stub_cli(
    monkeypatch: pytest.MonkeyPatch,
    *,
    stdout: str,
    returncode: int = 0,
    stderr: str = "",
) -> list[list[str]]:
    captured: list[list[str]] = []

    def fake_run(argv: list[str], *args: Any, **kwargs: Any) -> subprocess.CompletedProcess[str]:
        captured.append(list(argv))
        return subprocess.CompletedProcess(
            args=argv,
            returncode=returncode,
            stdout=stdout,
            stderr=stderr,
        )

    monkeypatch.setattr(subprocess, "run", fake_run)
    return captured


def _stub_which(monkeypatch: pytest.MonkeyPatch) -> None:
    import shutil

    monkeypatch.setattr(shutil, "which", lambda _: "/usr/bin/cargo")


def test_normalizes_valid_output(
    monkeypatch: pytest.MonkeyPatch, repo_root: Path, pdf_path: Path
) -> None:
    _stub_which(monkeypatch)
    stdout = json.dumps(
        {
            "kind": "Pdf",
            "pdf": {
                "pages": [
                    {"number": 2, "markdown": "# Page 2\n\nBody"},
                    {"number": 1, "markdown": "# Page 1"},
                ],
                "diagnostics": {
                    "classification": "TextBased",
                    "confidence": 0.9,
                    "page_count": 2,
                    "pages_needing_ocr": [],
                    "has_encoding_issues": False,
                },
            },
        }
    )
    captured = _stub_cli(monkeypatch, stdout=stdout)

    provider = OfficeMDLocalProvider(
        "officemd_local", _pipeline(repo_root).config
    )
    result = provider.run_inference_normalized(
        _pipeline(repo_root), _request(pdf_path)
    )

    assert [p.page_index for p in result.output.pages] == [0, 1]
    assert result.output.pages[0].markdown == "# Page 1"
    assert result.output.pages[1].markdown == "# Page 2\n\nBody"
    assert result.output.markdown == "# Page 1\n\n# Page 2\n\nBody"
    assert result.output.layout_pages == []
    assert result.raw_output["document"]["pdf"]["diagnostics"]["classification"] == "TextBased"

    argv = captured[0]
    assert argv[:5] == ["cargo", "run", "--quiet", "-p", "officemd_cli"]
    assert "stream" in argv
    assert "--output-format" in argv and "json" in argv
    assert "--pretty" in argv
    assert str(pdf_path) in argv


def test_preserves_page_order_multipage(
    monkeypatch: pytest.MonkeyPatch, repo_root: Path, pdf_path: Path
) -> None:
    _stub_which(monkeypatch)
    stdout = json.dumps(
        {
            "pdf": {
                "pages": [
                    {"number": i, "markdown": f"page{i}"} for i in range(1, 6)
                ],
                "diagnostics": {
                    "classification": "TextBased",
                    "confidence": 1.0,
                    "page_count": 5,
                    "pages_needing_ocr": [],
                    "has_encoding_issues": False,
                },
            },
        }
    )
    _stub_cli(monkeypatch, stdout=stdout)

    provider = OfficeMDLocalProvider(
        "officemd_local", _pipeline(repo_root).config
    )
    result = provider.run_inference_normalized(
        _pipeline(repo_root), _request(pdf_path)
    )

    assert [p.page_index for p in result.output.pages] == [0, 1, 2, 3, 4]
    assert result.output.markdown.split("\n\n") == [
        "page1",
        "page2",
        "page3",
        "page4",
        "page5",
    ]


def test_invalid_json_raises_permanent(
    monkeypatch: pytest.MonkeyPatch, repo_root: Path, pdf_path: Path
) -> None:
    _stub_which(monkeypatch)
    _stub_cli(monkeypatch, stdout="not json at all")

    provider = OfficeMDLocalProvider(
        "officemd_local", _pipeline(repo_root).config
    )
    with pytest.raises(ProviderPermanentError):
        provider.run_inference(_pipeline(repo_root), _request(pdf_path))


def test_nonzero_exit_raises_permanent(
    monkeypatch: pytest.MonkeyPatch, repo_root: Path, pdf_path: Path
) -> None:
    _stub_which(monkeypatch)
    _stub_cli(
        monkeypatch,
        stdout="",
        returncode=3,
        stderr="Error: something broke\n",
    )

    provider = OfficeMDLocalProvider(
        "officemd_local", _pipeline(repo_root).config
    )
    with pytest.raises(ProviderPermanentError):
        provider.run_inference(_pipeline(repo_root), _request(pdf_path))


def test_missing_pdf_payload_fails_normalize(
    monkeypatch: pytest.MonkeyPatch, repo_root: Path, pdf_path: Path
) -> None:
    _stub_which(monkeypatch)
    _stub_cli(monkeypatch, stdout=json.dumps({"kind": "Docx"}))

    provider = OfficeMDLocalProvider(
        "officemd_local", _pipeline(repo_root).config
    )
    raw = provider.run_inference(_pipeline(repo_root), _request(pdf_path))
    with pytest.raises(ProviderPermanentError):
        provider.normalize(raw)


def test_rejects_non_pdf_input(
    monkeypatch: pytest.MonkeyPatch, repo_root: Path, tmp_path: Path
) -> None:
    _stub_which(monkeypatch)
    docx = tmp_path / "file.docx"
    docx.write_bytes(b"PK\x03\x04")
    provider = OfficeMDLocalProvider(
        "officemd_local", _pipeline(repo_root).config
    )
    request = InferenceRequest(
        example_id="ex-1",
        source_file_path=str(docx),
        product_type=ProductType.PARSE,
    )
    with pytest.raises(ProviderPermanentError):
        provider.run_inference(_pipeline(repo_root), request)


def test_requires_repo_root_when_cargo_run(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("OFFICEMD_REPO_ROOT", raising=False)
    with pytest.raises(ProviderConfigError):
        OfficeMDLocalProvider(
            "officemd_local",
            {"cargo_run": True},
        )


def test_binary_mode_requires_existing_binary() -> None:
    with pytest.raises(ProviderConfigError):
        OfficeMDLocalProvider(
            "officemd_local",
            {"cargo_run": False, "binary": "/nonexistent/officemd"},
        )


def test_binary_mode_builds_argv_from_binary(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, pdf_path: Path
) -> None:
    binary = tmp_path / "officemd"
    binary.write_text("#!/bin/sh\n")
    binary.chmod(0o755)

    stdout = json.dumps(
        {
            "pdf": {
                "pages": [{"number": 1, "markdown": "hi"}],
                "diagnostics": {
                    "classification": "TextBased",
                    "confidence": 1.0,
                    "page_count": 1,
                    "pages_needing_ocr": [],
                    "has_encoding_issues": False,
                },
            }
        }
    )
    captured = _stub_cli(monkeypatch, stdout=stdout)

    provider = OfficeMDLocalProvider(
        "officemd_local",
        {"cargo_run": False, "binary": str(binary)},
    )
    pipeline = PipelineSpec(
        pipeline_name="officemd_local",
        provider_name="officemd_local",
        product_type=ProductType.PARSE,
        config={"cargo_run": False, "binary": str(binary)},
    )
    provider.run_inference_normalized(pipeline, _request(pdf_path))

    argv = captured[0]
    assert argv[0] == str(binary)
    assert argv[1] == "stream"
    assert str(pdf_path) in argv


def test_timing_and_latency(
    monkeypatch: pytest.MonkeyPatch, repo_root: Path, pdf_path: Path
) -> None:
    _stub_which(monkeypatch)
    stdout = json.dumps(
        {
            "pdf": {
                "pages": [{"number": 1, "markdown": "x"}],
                "diagnostics": {
                    "classification": "TextBased",
                    "confidence": 1.0,
                    "page_count": 1,
                    "pages_needing_ocr": [],
                    "has_encoding_issues": False,
                },
            }
        }
    )
    _stub_cli(monkeypatch, stdout=stdout)

    provider = OfficeMDLocalProvider(
        "officemd_local", _pipeline(repo_root).config
    )
    result = provider.run_inference(_pipeline(repo_root), _request(pdf_path))
    assert isinstance(result.started_at, datetime)
    assert isinstance(result.completed_at, datetime)
    assert result.latency_in_ms >= 0
