"""Provider that invokes the local OfficeMD CLI for PARSE tasks.

The provider shells out to `cargo run -p officemd_cli -- stream <path>
--output-format json --pretty` from the OfficeMD workspace root, parses the
resulting JSON document, and normalizes the per-page PDF payload into a
ParseBench `ParseOutput`.

Config keys (read from `PipelineSpec.config` via `base_config`):

- `repo_root` (str, required unless `OFFICEMD_REPO_ROOT` env var is set):
  absolute path to the OfficeMD workspace root (the directory containing the
  top-level `Cargo.toml`).
- `cargo_run` (bool, default True): when True, invoke via `cargo run`. When
  False, invoke the `binary` path directly.
- `binary` (str, optional): absolute path to a prebuilt `officemd` binary.
  Only used when `cargo_run` is False.
- `cargo_profile` (str, default "release"): cargo profile flag passed as
  `--release` when "release", otherwise passed as `--profile <value>`. Set
  to `"dev"` to drop the flag entirely.
- `extra_args` (list[str], optional): additional CLI flags appended after
  the input path (for example `["--no-headers-footers"]`).
- `timeout_seconds` (float | None, default 600): subprocess timeout.

Normalization rules:

- `pdf.pages[].number` maps to `PageIR.page_index = number - 1`.
- `pdf.pages[].markdown` becomes per-page markdown.
- Document markdown is page markdown joined by a single blank line.
- `layout_pages` is left empty in v1.
- The full OfficeMD JSON document is preserved in the raw output for
  downstream analysis.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

from parse_bench.inference.providers.base import (
    Provider,
    ProviderConfigError,
    ProviderPermanentError,
    ProviderTransientError,
)
from parse_bench.inference.providers.registry import register_provider
from parse_bench.schemas.parse_output import PageIR, ParseOutput
from parse_bench.schemas.pipeline import PipelineSpec
from parse_bench.schemas.pipeline_io import (
    InferenceRequest,
    InferenceResult,
    RawInferenceResult,
)
from parse_bench.schemas.product import ProductType

_DEFAULT_TIMEOUT_SECONDS = 600.0


@register_provider("officemd_local")
class OfficeMDLocalProvider(Provider):
    """Run the local OfficeMD CLI as a PARSE provider."""

    def __init__(self, provider_name: str, base_config: dict[str, Any] | None = None):
        super().__init__(provider_name, base_config)
        config = self._base_config

        self._cargo_run: bool = bool(config.get("cargo_run", True))
        self._cargo_profile: str = str(config.get("cargo_profile", "release"))
        self._extra_args: list[str] = list(config.get("extra_args") or [])
        self._timeout_seconds: float | None = config.get(
            "timeout_seconds", _DEFAULT_TIMEOUT_SECONDS
        )

        repo_root = config.get("repo_root") or os.environ.get("OFFICEMD_REPO_ROOT")
        self._repo_root: Path | None = Path(repo_root).resolve() if repo_root else None

        binary = config.get("binary")
        self._binary: Path | None = Path(binary).resolve() if binary else None

        # Validate mode eagerly so misconfiguration surfaces before first call.
        self._resolve_command_prefix()

    def _resolve_command_prefix(self) -> list[str]:
        """Return the argv prefix through `stream` (excluding input path)."""
        if self._cargo_run:
            if self._repo_root is None:
                raise ProviderConfigError(
                    "officemd_local: cargo_run mode requires `repo_root` in config "
                    "or the OFFICEMD_REPO_ROOT environment variable."
                )
            if not (self._repo_root / "Cargo.toml").is_file():
                raise ProviderConfigError(
                    f"officemd_local: repo_root {self._repo_root} does not contain Cargo.toml."
                )
            if shutil.which("cargo") is None:
                raise ProviderConfigError(
                    "officemd_local: `cargo` not found on PATH; install Rust or set "
                    "`cargo_run=False` and provide `binary`."
                )
            argv = ["cargo", "run", "--quiet", "-p", "officemd_cli"]
            profile = self._cargo_profile.lower()
            if profile == "release":
                argv.append("--release")
            elif profile not in ("dev", "debug", ""):
                argv.extend(["--profile", self._cargo_profile])
            argv.extend(["--", "stream"])
            return argv

        if self._binary is None:
            raise ProviderConfigError(
                "officemd_local: cargo_run is False but no `binary` path configured."
            )
        if not self._binary.is_file():
            raise ProviderConfigError(
                f"officemd_local: configured binary {self._binary} does not exist."
            )
        return [str(self._binary), "stream"]

    def _working_directory(self) -> Path | None:
        if self._cargo_run:
            return self._repo_root
        return None

    def _build_argv(self, pdf_path: Path) -> list[str]:
        argv = self._resolve_command_prefix()
        argv.append(str(pdf_path))
        argv.extend(["--output-format", "json", "--pretty"])
        argv.extend(self._extra_args)
        return argv

    def _invoke_cli(self, pdf_path: Path) -> tuple[dict[str, Any], str]:
        argv = self._build_argv(pdf_path)
        cwd = self._working_directory()
        try:
            completed = subprocess.run(  # noqa: S603 - argv is constructed internally
                argv,
                cwd=str(cwd) if cwd is not None else None,
                capture_output=True,
                text=True,
                timeout=self._timeout_seconds,
                check=False,
            )
        except FileNotFoundError as exc:
            raise ProviderConfigError(
                f"officemd_local: failed to launch CLI ({exc}). "
                "Check `cargo` / `binary` configuration."
            ) from exc
        except subprocess.TimeoutExpired as exc:
            raise ProviderTransientError(
                f"officemd_local: CLI exceeded timeout of {self._timeout_seconds}s"
            ) from exc

        if completed.returncode != 0:
            stderr_tail = (completed.stderr or "").strip().splitlines()[-20:]
            raise ProviderPermanentError(
                "officemd_local: CLI exited with code "
                f"{completed.returncode}. stderr (tail):\n"
                + "\n".join(stderr_tail),
                debug_payload={
                    "argv": argv,
                    "returncode": completed.returncode,
                    "stderr": completed.stderr,
                },
            )

        stdout = completed.stdout or ""
        try:
            document = json.loads(stdout)
        except json.JSONDecodeError as exc:
            raise ProviderPermanentError(
                f"officemd_local: failed to parse CLI JSON output: {exc}",
                debug_payload={"argv": argv, "stdout_head": stdout[:2000]},
            ) from exc

        if not isinstance(document, dict):
            raise ProviderPermanentError(
                "officemd_local: CLI JSON output is not a JSON object.",
                debug_payload={"argv": argv},
            )

        return document, " ".join(argv)

    def run_inference(
        self, pipeline: PipelineSpec, request: InferenceRequest
    ) -> RawInferenceResult:
        if request.product_type != ProductType.PARSE:
            raise ProviderPermanentError(
                "officemd_local only supports PARSE product type, got "
                f"{request.product_type}"
            )

        pdf_path = Path(request.source_file_path)
        if pdf_path.suffix.lower() != ".pdf":
            raise ProviderPermanentError(
                f"officemd_local only supports .pdf files, got {pdf_path.suffix}"
            )
        if not pdf_path.exists():
            raise ProviderPermanentError(f"PDF file not found: {pdf_path}")

        started_at = datetime.now()
        document, argv_display = self._invoke_cli(pdf_path)
        completed_at = datetime.now()
        latency_ms = int((completed_at - started_at).total_seconds() * 1000)

        raw_output: dict[str, Any] = {
            "document": document,
            "argv": argv_display,
            "source_file_path": str(pdf_path),
        }

        return RawInferenceResult(
            request=request,
            pipeline=pipeline,
            pipeline_name=pipeline.pipeline_name,
            product_type=request.product_type,
            raw_output=raw_output,
            started_at=started_at,
            completed_at=completed_at,
            latency_in_ms=latency_ms,
        )

    def normalize(self, raw_result: RawInferenceResult) -> InferenceResult:
        if raw_result.product_type != ProductType.PARSE:
            raise ProviderPermanentError(
                "officemd_local only supports PARSE product type, got "
                f"{raw_result.product_type}"
            )

        document = raw_result.raw_output.get("document")
        if not isinstance(document, dict):
            raise ProviderPermanentError(
                "officemd_local: raw_output missing `document` object."
            )

        pdf_payload = document.get("pdf")
        if not isinstance(pdf_payload, dict):
            raise ProviderPermanentError(
                "officemd_local: CLI output missing `pdf` payload; "
                "is the input file a PDF?"
            )

        pages_payload = pdf_payload.get("pages")
        if not isinstance(pages_payload, list):
            raise ProviderPermanentError(
                "officemd_local: `pdf.pages` is missing or not a list."
            )

        pages = _normalize_pages(pages_payload)
        document_markdown = "\n\n".join(page.markdown for page in pages)

        output = ParseOutput(
            task_type="parse",
            example_id=raw_result.request.example_id,
            pipeline_name=raw_result.pipeline_name,
            pages=pages,
            markdown=document_markdown,
        )

        return InferenceResult(
            request=raw_result.request,
            pipeline_name=raw_result.pipeline_name,
            product_type=raw_result.product_type,
            raw_output=raw_result.raw_output,
            output=output,
            started_at=raw_result.started_at,
            completed_at=raw_result.completed_at,
            latency_in_ms=raw_result.latency_in_ms,
        )


def _normalize_pages(pages_payload: list[Any]) -> list[PageIR]:
    """Map OfficeMD `pdf.pages` entries to sorted PageIR (0-indexed)."""
    pages: list[PageIR] = []
    for entry in pages_payload:
        if not isinstance(entry, dict):
            raise ProviderPermanentError(
                "officemd_local: `pdf.pages` entry is not an object."
            )
        raw_number = entry.get("number")
        if not isinstance(raw_number, int) or raw_number < 1:
            raise ProviderPermanentError(
                "officemd_local: `pdf.pages[].number` must be a positive integer, "
                f"got {raw_number!r}."
            )
        markdown = entry.get("markdown")
        if markdown is None:
            markdown = ""
        elif not isinstance(markdown, str):
            raise ProviderPermanentError(
                "officemd_local: `pdf.pages[].markdown` must be a string or null, "
                f"got {type(markdown).__name__}."
            )
        pages.append(PageIR(page_index=raw_number - 1, markdown=markdown))

    pages.sort(key=lambda p: p.page_index)
    return pages
