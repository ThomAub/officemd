#!/usr/bin/env python3
"""Materialize a non-OCR PDF slice using OfficeMD's classifier.

Runs `officemd inspect --output-format json` against a directory of PDFs (or a
list of paths) and emits:

- a JSONL file with one row per PDF listing classification, page count, OCR
  requirement, encoding issues, and pass/fail for the "non-OCR" slice, and
- a plain-text manifest containing only the paths that qualify.

A PDF qualifies for the non-OCR slice when:

- the OfficeMD classification is `TextBased`, AND
- `pages_needing_ocr` is empty.

Usage from the OfficeMD repo root::

    uv run integrations/parsebench/scripts/classify_non_ocr_pdfs.py \\
        --input-dir /path/to/pdfs \\
        --report-jsonl non_ocr_report.jsonl \\
        --manifest non_ocr_manifest.txt

Or explicitly listing files::

    uv run .../classify_non_ocr_pdfs.py file1.pdf file2.pdf \\
        --report-jsonl report.jsonl --manifest manifest.txt

By default, the script uses `cargo run -p officemd_cli --release`. Pass
`--binary /path/to/officemd` to use a prebuilt binary instead.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from collections.abc import Iterable
from pathlib import Path


def _collect_pdfs(
    input_dir: Path | None, explicit_paths: list[Path]
) -> list[Path]:
    pdfs: list[Path] = list(explicit_paths)
    if input_dir is not None:
        pdfs.extend(
            sorted(p for p in input_dir.rglob("*.pdf") if p.is_file())
        )

    # Keep stable order, drop duplicates.
    seen: set[Path] = set()
    deduped: list[Path] = []
    for path in pdfs:
        resolved = path.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        deduped.append(resolved)
    return deduped


def _inspect_pdf(
    pdf: Path,
    *,
    cargo_run: bool,
    repo_root: Path,
    binary: Path | None,
    cargo_profile: str,
) -> dict[str, object]:
    if cargo_run:
        argv = ["cargo", "run", "--quiet", "-p", "officemd_cli"]
        profile = cargo_profile.lower()
        if profile == "release":
            argv.append("--release")
        elif profile not in ("dev", "debug", ""):
            argv.extend(["--profile", cargo_profile])
        argv.extend(["--", "inspect", str(pdf), "--output-format", "json"])
        cwd: str | None = str(repo_root)
    else:
        assert binary is not None
        argv = [str(binary), "inspect", str(pdf), "--output-format", "json"]
        cwd = None

    completed = subprocess.run(  # noqa: S603 - argv constructed internally
        argv,
        cwd=cwd,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        return {
            "error": "inspect_failed",
            "returncode": completed.returncode,
            "stderr_tail": (completed.stderr or "").strip().splitlines()[-5:],
        }

    try:
        payload = json.loads(completed.stdout or "")
    except json.JSONDecodeError as exc:
        return {"error": f"invalid_json: {exc}"}

    pdf_info = payload.get("pdf") if isinstance(payload, dict) else None
    if not isinstance(pdf_info, dict):
        return {"error": "missing_pdf_info"}
    return pdf_info


def _is_non_ocr(pdf_info: dict[str, object]) -> bool:
    classification = pdf_info.get("classification")
    pages_needing_ocr = pdf_info.get("pages_needing_ocr")
    return (
        classification == "TextBased"
        and isinstance(pages_needing_ocr, list)
        and len(pages_needing_ocr) == 0
    )


def _default_repo_root() -> Path:
    env = os.environ.get("OFFICEMD_REPO_ROOT")
    if env:
        return Path(env).resolve()
    # Walk upwards from this script until we find a Cargo.toml.
    here = Path(__file__).resolve()
    for candidate in here.parents:
        if (candidate / "Cargo.toml").is_file():
            return candidate
    return Path.cwd().resolve()


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "pdfs",
        nargs="*",
        type=Path,
        help="Explicit PDF paths to classify (combined with --input-dir).",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=None,
        help="Directory to scan recursively for *.pdf files.",
    )
    parser.add_argument(
        "--report-jsonl",
        type=Path,
        required=True,
        help="Path to write the per-PDF classification report (JSONL).",
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        required=True,
        help="Path to write the plain-text list of non-OCR PDFs (one per line).",
    )
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=_default_repo_root(),
        help="OfficeMD workspace root used when `--cargo-run` is set.",
    )
    parser.add_argument(
        "--binary",
        type=Path,
        default=None,
        help="Path to a prebuilt officemd binary. Disables --cargo-run when set.",
    )
    parser.add_argument(
        "--cargo-profile",
        default="release",
        help="Cargo profile for `cargo run`. Default: release.",
    )
    parser.add_argument(
        "--no-cargo-run",
        dest="cargo_run",
        action="store_false",
        help="Require --binary instead of invoking cargo.",
    )
    parser.set_defaults(cargo_run=True)

    args = parser.parse_args(list(argv) if argv is not None else None)

    cargo_run = args.cargo_run and args.binary is None
    if cargo_run and shutil.which("cargo") is None:
        print("ERROR: cargo not found on PATH.", file=sys.stderr)
        return 2
    if not cargo_run and args.binary is None:
        print("ERROR: must provide --binary when --no-cargo-run is used.", file=sys.stderr)
        return 2
    if not cargo_run and not args.binary.is_file():
        print(f"ERROR: binary {args.binary} not found.", file=sys.stderr)
        return 2

    pdfs = _collect_pdfs(args.input_dir, args.pdfs)
    if not pdfs:
        print("ERROR: no PDFs provided (use positional args or --input-dir).", file=sys.stderr)
        return 2

    args.report_jsonl.parent.mkdir(parents=True, exist_ok=True)
    args.manifest.parent.mkdir(parents=True, exist_ok=True)

    selected: list[Path] = []
    with args.report_jsonl.open("w", encoding="utf-8") as report_fh:
        for pdf in pdfs:
            pdf_info = _inspect_pdf(
                pdf,
                cargo_run=cargo_run,
                repo_root=args.repo_root,
                binary=args.binary,
                cargo_profile=args.cargo_profile,
            )
            non_ocr = "error" not in pdf_info and _is_non_ocr(pdf_info)
            record = {
                "path": str(pdf),
                "non_ocr": non_ocr,
                "pdf": pdf_info,
            }
            report_fh.write(json.dumps(record) + "\n")
            if non_ocr:
                selected.append(pdf)

    with args.manifest.open("w", encoding="utf-8") as manifest_fh:
        for pdf in selected:
            manifest_fh.write(f"{pdf}\n")

    print(
        f"Classified {len(pdfs)} PDFs; {len(selected)} in non-OCR slice.",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
