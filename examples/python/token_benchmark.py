#!/usr/bin/env python3
import argparse
import json
import os
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable

SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".csv", ".pptx", ".pdf"}


@dataclass
class BenchmarkRow:
    file: str
    compact_tokens: int
    human_tokens: int
    token_delta: int
    token_savings_pct: float
    compact_chars: int
    human_chars: int


TokenCounter = Callable[[str], int]


def repo_root() -> Path:
    return Path(__file__).resolve().parents[3]


def default_inputs(root: Path) -> list[Path]:
    data_dir = root / "crates" / "examples" / "data"
    if not data_dir.exists():
        return []
    return sorted(
        p
        for p in data_dir.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def build_token_counter(
    tokenizer: str, encoding: str, model: str | None
) -> tuple[str, TokenCounter]:
    if tokenizer == "anthropic":
        if not model:
            raise SystemExit("--model is required when --tokenizer anthropic")
        return build_anthropic_counter(model)
    if tokenizer == "gemini":
        if not model:
            raise SystemExit("--model is required when --tokenizer gemini")
        return build_gemini_counter(model)

    if tokenizer == "approx":
        return "approx(chars/4)", lambda text: 0 if not text else max(1, len(text) // 4)

    try:
        import tiktoken
    except ModuleNotFoundError:
        print(
            "warning: tiktoken not installed, falling back to approx(chars/4)",
            file=sys.stderr,
        )
        return "approx(chars/4)", lambda text: 0 if not text else max(1, len(text) // 4)

    try:
        enc = tiktoken.get_encoding(encoding)
    except Exception:
        try:
            enc = tiktoken.encoding_for_model(encoding)
        except Exception as exc:
            raise SystemExit(
                f"Could not load tokenizer '{encoding}'. Use --tokenizer approx or a valid tiktoken encoding/model."
            ) from exc

    return f"tiktoken:{encoding}", lambda text: len(enc.encode(text))


def _http_json_post(url: str, payload: dict, headers: dict[str, str]) -> dict:
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} on {url}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Request failed for {url}: {exc.reason}") from exc

    try:
        return json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Non-JSON response from {url}: {body[:300]}") from exc


def build_anthropic_counter(model: str) -> tuple[str, TokenCounter]:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise SystemExit("ANTHROPIC_API_KEY is required when --tokenizer anthropic")

    endpoint = "https://api.anthropic.com/v1/messages/count_tokens"

    def counter(text: str) -> int:
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": text}],
        }
        data = _http_json_post(
            endpoint,
            payload,
            {
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
        )
        if "input_tokens" not in data:
            raise RuntimeError(f"Unexpected Anthropic response: {data}")
        return int(data["input_tokens"])

    return f"anthropic:{model}", counter


def build_gemini_counter(model: str) -> tuple[str, TokenCounter]:
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise SystemExit(
            "GEMINI_API_KEY or GOOGLE_API_KEY is required when --tokenizer gemini"
        )

    model_name = model if model.startswith("models/") else f"models/{model}"
    endpoint = (
        "https://generativelanguage.googleapis.com/v1beta/"
        f"{model_name}:countTokens?key={urllib.parse.quote(api_key, safe='')}"
    )

    def counter(text: str) -> int:
        payload = {"contents": [{"parts": [{"text": text}]}]}
        data = _http_json_post(endpoint, payload, {"content-type": "application/json"})
        if "totalTokens" not in data:
            raise RuntimeError(f"Unexpected Gemini response: {data}")
        return int(data["totalTokens"])

    return f"gemini:{model_name}", counter


def render_markdown_via_cli(
    root: Path,
    path: Path,
    markdown_style: str,
    include_document_properties: bool,
) -> str:
    cmd = [
        "cargo",
        "run",
        "-q",
        "-p",
        "officemd_cli",
        "--",
        "stream",
        str(path),
        "--markdown-style",
        markdown_style,
    ]
    if include_document_properties:
        cmd.append("--include-document-properties")

    proc = subprocess.run(
        cmd,
        cwd=root,
        text=True,
        capture_output=True,
        check=False,
    )
    if proc.returncode != 0:
        detail = proc.stderr.strip() or proc.stdout.strip()
        raise RuntimeError(f"CLI failed for {path} ({markdown_style}): {detail}")
    return proc.stdout


def benchmark_file(
    path: Path,
    counter: TokenCounter,
    include_document_properties: bool,
    root: Path,
) -> BenchmarkRow:
    compact = render_markdown_via_cli(
        root=root,
        path=path,
        markdown_style="compact",
        include_document_properties=include_document_properties,
    )
    human = render_markdown_via_cli(
        root=root,
        path=path,
        markdown_style="human",
        include_document_properties=include_document_properties,
    )

    compact_tokens = counter(compact)
    human_tokens = counter(human)
    delta = human_tokens - compact_tokens
    savings_pct = (delta / human_tokens * 100.0) if human_tokens else 0.0

    return BenchmarkRow(
        file=str(path.relative_to(root)),
        compact_tokens=compact_tokens,
        human_tokens=human_tokens,
        token_delta=delta,
        token_savings_pct=savings_pct,
        compact_chars=len(compact),
        human_chars=len(human),
    )


def print_table(rows: list[BenchmarkRow], tokenizer_name: str) -> None:
    headers = [
        "file",
        "compact_tok",
        "human_tok",
        "delta",
        "saved%",
        "compact_chr",
        "human_chr",
    ]
    data = [
        [
            row.file,
            str(row.compact_tokens),
            str(row.human_tokens),
            str(row.token_delta),
            f"{row.token_savings_pct:.1f}",
            str(row.compact_chars),
            str(row.human_chars),
        ]
        for row in rows
    ]

    widths = [len(h) for h in headers]
    for row in data:
        for idx, col in enumerate(row):
            widths[idx] = max(widths[idx], len(col))

    print(f"tokenizer={tokenizer_name}")
    print(" ".join(headers[idx].ljust(widths[idx]) for idx in range(len(headers))))
    print(" ".join("-" * widths[idx] for idx in range(len(headers))))
    for row in data:
        print(
            " ".join(
                row[idx].ljust(widths[idx]) if idx == 0 else row[idx].rjust(widths[idx])
                for idx in range(len(row))
            )
        )

    total_compact = sum(r.compact_tokens for r in rows)
    total_human = sum(r.human_tokens for r in rows)
    total_delta = total_human - total_compact
    total_saved = (total_delta / total_human * 100.0) if total_human else 0.0
    print()
    print(
        f"TOTAL compact={total_compact} human={total_human} delta={total_delta} saved={total_saved:.1f}%"
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare compact vs human markdown token counts on fixture files."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        type=Path,
        help="Files to benchmark (defaults to supported files in examples/data)",
    )
    parser.add_argument(
        "--tokenizer",
        choices=["auto", "approx", "anthropic", "gemini"],
        default="auto",
        help=(
            "Token counter: auto uses tiktoken if available, approx uses chars/4, "
            "anthropic uses Messages count_tokens API, gemini uses models.countTokens API."
        ),
    )
    parser.add_argument(
        "--encoding",
        default="o200k_base",
        help="tiktoken encoding/model when --tokenizer auto (default: o200k_base)",
    )
    parser.add_argument(
        "--model",
        default=None,
        help="Provider model id for --tokenizer anthropic/gemini (for example claude-opus-4-1 or gemini-2.5-pro).",
    )
    parser.add_argument(
        "--include-document-properties",
        action="store_true",
        help="Include document properties in both markdown outputs.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit JSON instead of table output.",
    )
    parser.add_argument(
        "--fail-if-compact-worse",
        action="store_true",
        help="Exit non-zero if compact has more tokens than human for any file.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root = repo_root()

    if args.inputs:
        inputs = [p if p.is_absolute() else root / p for p in args.inputs]
    else:
        inputs = default_inputs(root)

    if not inputs:
        print("No input files found.", file=sys.stderr)
        return 1

    unknown = [p for p in inputs if p.suffix.lower() not in SUPPORTED_EXTENSIONS]
    if unknown:
        print(
            "Unsupported file type(s): " + ", ".join(str(p) for p in unknown),
            file=sys.stderr,
        )
        return 2

    missing = [p for p in inputs if not p.exists()]
    if missing:
        print("Missing file(s): " + ", ".join(str(p) for p in missing), file=sys.stderr)
        return 3

    tokenizer_name, counter = build_token_counter(
        args.tokenizer, args.encoding, args.model
    )

    rows: list[BenchmarkRow] = []
    for path in inputs:
        try:
            rows.append(
                benchmark_file(
                    path,
                    counter,
                    include_document_properties=args.include_document_properties,
                    root=root,
                )
            )
        except Exception as exc:
            print(str(exc), file=sys.stderr)
            return 5

    if args.json:
        payload = {
            "tokenizer": tokenizer_name,
            "rows": [asdict(r) for r in rows],
            "totals": {
                "compact_tokens": sum(r.compact_tokens for r in rows),
                "human_tokens": sum(r.human_tokens for r in rows),
                "token_delta": sum(r.token_delta for r in rows),
            },
        }
        print(json.dumps(payload, indent=2))
    else:
        print_table(rows, tokenizer_name)

    if args.fail_if_compact_worse:
        regressions = [r for r in rows if r.compact_tokens > r.human_tokens]
        if regressions:
            print("\ncompact-worse files:", file=sys.stderr)
            for row in regressions:
                print(
                    f"- {row.file}: compact={row.compact_tokens}, human={row.human_tokens}",
                    file=sys.stderr,
                )
            return 4

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
