# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Download a small PDF sample from the surgeai/GDP.pdf Hugging Face dataset.

Usage:
    uv run benchmark/setup-gdp-pdf-corpus.py
    uv run benchmark/setup-gdp-pdf-corpus.py --limit 20
"""

import argparse
import json
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any


DATASET_ID = "surgeai/GDP.pdf"
REVISION = "main"
TREE_URL = f"https://huggingface.co/api/datasets/{DATASET_ID}/tree/{REVISION}?recursive=true"
RESOLVE_BASE = f"https://huggingface.co/datasets/{DATASET_ID}/resolve/{REVISION}/"


def default_output_dir() -> Path:
    return Path(__file__).resolve().parent / "corpus" / "gdp-pdf"


def human_bytes(size: int | None) -> str:
    if size is None:
        return "unknown size"
    value = float(size)
    for unit in ["B", "KB", "MB", "GB"]:
        if value < 1024 or unit == "GB":
            return f"{value:.1f}{unit}" if unit != "B" else f"{int(value)}B"
        value /= 1024
    raise AssertionError("unreachable")


def load_dataset_tree() -> list[dict[str, Any]]:
    request = urllib.request.Request(
        TREE_URL,
        headers={"User-Agent": "officemd-benchmark-downloader/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            payload = json.load(response)
    except urllib.error.URLError as exc:
        raise RuntimeError(f"failed to fetch dataset tree: {exc}") from exc

    if not isinstance(payload, list):
        raise RuntimeError("unexpected Hugging Face tree response")
    return [entry for entry in payload if isinstance(entry, dict)]


def discover_pdfs() -> list[dict[str, Any]]:
    entries = load_dataset_tree()
    pdfs = [
        entry
        for entry in entries
        if entry.get("type") == "file"
        and isinstance(entry.get("path"), str)
        and entry["path"].startswith("pdfs/")
        and entry["path"].lower().endswith(".pdf")
    ]
    if not pdfs:
        raise RuntimeError("no PDFs found in the Hugging Face dataset tree")
    return pdfs


def source_url(path: str) -> str:
    return RESOLVE_BASE + urllib.parse.quote(path)


def has_pdf_header(path: Path) -> bool:
    try:
        with path.open("rb") as handle:
            return handle.read(5) == b"%PDF-"
    except OSError:
        return False


def download(entry: dict[str, Any], output_dir: Path) -> Path:
    source_path = str(entry["path"])
    target = output_dir / Path(source_path).name
    expected_size = entry.get("size")
    if not isinstance(expected_size, int):
        expected_size = None

    if target.exists() and has_pdf_header(target):
        actual_size = target.stat().st_size
        if expected_size is None or actual_size == expected_size:
            print(f"  skip {target.name} ({human_bytes(actual_size)})")
            return target

    request = urllib.request.Request(
        source_url(source_path),
        headers={"User-Agent": "officemd-benchmark-downloader/1.0"},
    )
    print(f"  downloading {target.name} ({human_bytes(expected_size)})")
    with tempfile.NamedTemporaryFile(dir=output_dir, delete=False) as tmp:
        tmp_path = Path(tmp.name)
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    tmp.write(chunk)
        except Exception:
            tmp_path.unlink(missing_ok=True)
            raise

    if expected_size is not None and tmp_path.stat().st_size != expected_size:
        actual = human_bytes(tmp_path.stat().st_size)
        tmp_path.unlink(missing_ok=True)
        raise RuntimeError(f"{target.name} downloaded with unexpected size {actual}")
    if not has_pdf_header(tmp_path):
        tmp_path.unlink(missing_ok=True)
        raise RuntimeError(f"{target.name} is not a valid PDF")

    tmp_path.replace(target)
    return target


def write_manifest(output_dir: Path, downloaded: list[dict[str, Any]]) -> None:
    manifest = {
        "dataset": DATASET_ID,
        "revision": REVISION,
        "source": f"https://huggingface.co/datasets/{DATASET_ID}",
        "files": downloaded,
    }
    with (output_dir / "manifest.json").open("w") as handle:
        json.dump(manifest, handle, indent=2)
        handle.write("\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=default_output_dir(),
        help="Directory for downloaded PDFs",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="Number of PDFs to download",
    )
    parser.add_argument(
        "--sort",
        choices=["size", "path"],
        default="size",
        help="Selection order before applying --limit",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print selected files without downloading",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.limit < 1:
        print("ERROR: --limit must be at least 1", file=sys.stderr)
        return 2

    pdfs = discover_pdfs()
    if args.sort == "size":
        pdfs.sort(key=lambda entry: (entry.get("size", sys.maxsize), entry["path"]))
    else:
        pdfs.sort(key=lambda entry: entry["path"])

    selected = pdfs[: args.limit]
    print(f"Found {len(pdfs)} PDFs in {DATASET_ID}; selected {len(selected)}.")

    if args.dry_run:
        for entry in selected:
            print(f"  {entry['path']} ({human_bytes(entry.get('size'))})")
        return 0

    args.output_dir.mkdir(parents=True, exist_ok=True)
    downloaded = []
    for entry in selected:
        target = download(entry, args.output_dir)
        downloaded.append(
            {
                "source_path": entry["path"],
                "local_path": str(target),
                "size": target.stat().st_size,
            }
        )

    write_manifest(args.output_dir, downloaded)
    print(f"Corpus ready: {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
