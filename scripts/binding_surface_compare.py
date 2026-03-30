#!/usr/bin/env python3
"""Compare JS and Python binding surfaces for officemd.

This script validates JS, Python, and both Rust binding implementations against a
canonical contract:
- exported function set
- parameter names and order
- Python default values
- JS declared return types
- Rust binding function names / parameter shapes / return types
- basic runtime return types for Python, and optionally JS when a built native
  module is available
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parent.parent
JS_DTS = REPO_ROOT / "crates" / "officemd_js" / "index.d.ts"
JS_RUST = REPO_ROOT / "crates" / "officemd_js" / "src" / "lib.rs"
PYTHON_CRATE = REPO_ROOT / "crates" / "officemd_python"
PYTHON_RUST = REPO_ROOT / "crates" / "officemd_python" / "src" / "lib.rs"

FUNCTION_RE = re.compile(r"export declare function (\w+)\((.*?)\):\s*([^\n]+)", re.S)
PARAM_RE = re.compile(r"\s*(\w+)\??:\s*(.+)")
JS_RUST_FN_RE = re.compile(r"#\[napi\][^\n]*\n(?:#\[[^\n]+\]\n)*pub fn (\w+)\((.*?)\) -> Result<(.+?)>\s*\{", re.S)
PY_RUST_ATTR_RE = re.compile(r"#\[pyfunction(?:\(signature = \((.*?)\)\))?\]\nfn (\w+)\((.*?)\) -> PyResult<(.+?)>\s*\{", re.S)


@dataclass(frozen=True)
class ContractFunction:
    name: str
    params: tuple[str, ...]
    python_defaults: dict[str, object]
    js_return_type: str
    python_runtime_type: str | None = None
    js_runtime_type: str | None = None


@dataclass(frozen=True)
class JsFunction:
    name: str
    params: tuple[str, ...]
    optional_params: frozenset[str]
    return_type: str


@dataclass(frozen=True)
class PythonFunction:
    name: str
    params: tuple[str, ...]
    defaults: dict[str, object]


@dataclass(frozen=True)
class RustBindingFunction:
    name: str
    params: tuple[str, ...]
    return_type: str


CONTRACT: dict[str, ContractFunction] = {
    "apply_ooxml_patch_json": ContractFunction(
        name="apply_ooxml_patch_json",
        params=("content", "patch_json"),
        python_defaults={},
        js_return_type="bytes",
        python_runtime_type="bytes",
        js_runtime_type="bytes",
    ),
    "create_document_from_markdown": ContractFunction(
        name="create_document_from_markdown",
        params=("markdown", "format"),
        python_defaults={},
        js_return_type="bytes",
        python_runtime_type="bytes",
        js_runtime_type="bytes",
    ),
    "detect_format": ContractFunction(
        name="detect_format",
        params=("content",),
        python_defaults={},
        js_return_type="str",
        python_runtime_type="str",
        js_runtime_type="str",
    ),
    "docling_from_bytes": ContractFunction(
        name="docling_from_bytes",
        params=("content", "format"),
        python_defaults={"format": None},
        js_return_type="str",
    ),
    "extract_csv_tables_ir_json": ContractFunction(
        name="extract_csv_tables_ir_json",
        params=("content", "delimiter", "include_document_properties"),
        python_defaults={"delimiter": ",", "include_document_properties": False},
        js_return_type="str",
    ),
    "extract_ir_json": ContractFunction(
        name="extract_ir_json",
        params=("content", "format"),
        python_defaults={"format": None},
        js_return_type="str",
    ),
    "extract_sheet_names": ContractFunction(
        name="extract_sheet_names",
        params=("content",),
        python_defaults={},
        js_return_type="list[str]",
    ),
    "extract_tables_ir_json": ContractFunction(
        name="extract_tables_ir_json",
        params=("content", "style_aware_values", "streaming_rows", "include_document_properties"),
        python_defaults={
            "style_aware_values": False,
            "streaming_rows": False,
            "include_document_properties": False,
        },
        js_return_type="str",
    ),
    "inspect_pdf_fonts_json": ContractFunction(
        name="inspect_pdf_fonts_json",
        params=("content",),
        python_defaults={},
        js_return_type="str",
    ),
    "inspect_pdf_json": ContractFunction(
        name="inspect_pdf_json",
        params=("content",),
        python_defaults={},
        js_return_type="str",
    ),
    "markdown_from_bytes": ContractFunction(
        name="markdown_from_bytes",
        params=(
            "content",
            "format",
            "include_document_properties",
            "use_first_row_as_header",
            "include_headers_footers",
            "include_formulas",
            "markdown_style",
            "force_extract",
        ),
        python_defaults={
            "format": None,
            "include_document_properties": False,
            "use_first_row_as_header": True,
            "include_headers_footers": True,
            "include_formulas": True,
            "markdown_style": None,
            "force_extract": False,
        },
        js_return_type="str",
    ),
    "markdown_from_bytes_batch": ContractFunction(
        name="markdown_from_bytes_batch",
        params=(
            "contents",
            "format",
            "workers",
            "include_document_properties",
            "use_first_row_as_header",
            "include_headers_footers",
            "include_formulas",
            "markdown_style",
        ),
        python_defaults={
            "format": None,
            "workers": None,
            "include_document_properties": False,
            "use_first_row_as_header": True,
            "include_headers_footers": True,
            "include_formulas": True,
            "markdown_style": None,
        },
        js_return_type="list[str]",
    ),
}


def camel_to_snake(name: str) -> str:
    return re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()


def normalize_js_type(type_expr: str) -> str:
    type_expr = " ".join(type_expr.split())
    type_expr = type_expr.replace("Buffer", "bytes")
    type_expr = type_expr.replace("Array<string>", "list[str]")
    type_expr = type_expr.replace("Array<bytes>", "list[bytes]")
    type_expr = type_expr.replace("Vec<String>", "list[str]")
    type_expr = type_expr.replace("Vec<u8>", "bytes")
    type_expr = type_expr.replace("String", "str")
    type_expr = type_expr.replace("u32", "int")
    type_expr = type_expr.replace("usize", "int")
    type_expr = type_expr.replace("bool", "bool")
    type_expr = type_expr.replace("string", "str")
    type_expr = type_expr.replace("number", "int")
    type_expr = type_expr.replace("boolean", "bool")
    type_expr = re.sub(r"\s*\|\s*undefined\s*\|\s*null", " | None", type_expr)
    return type_expr.rstrip(";")


def normalize_python_default(value: object) -> object:
    if isinstance(value, (str, bool, int)) or value is None:
        return value
    return repr(value)


def parse_signature_params(raw: str) -> list[str]:
    params: list[str] = []
    for raw_param in [chunk.strip() for chunk in raw.split(",") if chunk.strip()]:
        name = raw_param.split(":", 1)[0].strip().rstrip("?")
        params.append(camel_to_snake(name))
    return params


def load_js_surface() -> dict[str, JsFunction]:
    text = JS_DTS.read_text()
    surface: dict[str, JsFunction] = {}
    for match in FUNCTION_RE.finditer(text):
        name = camel_to_snake(match.group(1))
        raw_params = match.group(2)
        params: list[str] = []
        optional_params: set[str] = set()
        for raw_param in [chunk.strip() for chunk in raw_params.split(",") if chunk.strip()]:
            param_match = PARAM_RE.fullmatch(raw_param)
            if not param_match:
                raise ValueError(f"Could not parse JS parameter: {raw_param!r}")
            original_name = param_match.group(1)
            name_snake = camel_to_snake(original_name)
            params.append(name_snake)
            if "?" in raw_param.split(":", 1)[0]:
                optional_params.add(name_snake)
        surface[name] = JsFunction(
            name=name,
            params=tuple(params),
            optional_params=frozenset(optional_params),
            return_type=normalize_js_type(match.group(3).strip()),
        )
    return surface


def load_python_surface() -> dict[str, PythonFunction]:
    script = r'''
import inspect
import json
import officemd

names = [
    "apply_ooxml_patch_json",
    "create_document_from_markdown",
    "detect_format",
    "docling_from_bytes",
    "extract_csv_tables_ir_json",
    "extract_ir_json",
    "extract_sheet_names",
    "extract_tables_ir_json",
    "inspect_pdf_fonts_json",
    "inspect_pdf_json",
    "markdown_from_bytes",
    "markdown_from_bytes_batch",
]

payload = {}
for name in names:
    signature = inspect.signature(getattr(officemd, name))
    defaults = {}
    for param in signature.parameters.values():
        if param.default is not inspect._empty:
            defaults[param.name] = param.default
    payload[name] = {
        "params": [param.name for param in signature.parameters.values()],
        "defaults": defaults,
    }
print(json.dumps(payload, sort_keys=True))
'''
    completed = subprocess.run(
        ["uv", "run", "python", "-c", script],
        cwd=PYTHON_CRATE,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        sys.stderr.write(completed.stderr)
        raise SystemExit(completed.returncode)

    payload = json.loads(completed.stdout.strip().splitlines()[-1])
    surface: dict[str, PythonFunction] = {}
    for name, info in payload.items():
        defaults = {
            key: normalize_python_default(value) for key, value in info["defaults"].items()
        }
        surface[name] = PythonFunction(name=name, params=tuple(info["params"]), defaults=defaults)
    return surface


def load_python_runtime_types() -> dict[str, str]:
    script = r'''
import json
import officemd

sample_docx = officemd.create_document_from_markdown("## Section: body\n\nHello\n", "docx")
patched_docx = officemd.apply_ooxml_patch_json(
    sample_docx,
    '{"edits":[{"part":"word/document.xml","from":"Hello","to":"Hello runtime"}]}',
)
values = {
    "apply_ooxml_patch_json": patched_docx,
    "create_document_from_markdown": sample_docx,
    "detect_format": officemd.detect_format(sample_docx),
}
print(json.dumps({name: type(value).__name__ for name, value in values.items()}, sort_keys=True))
'''
    completed = subprocess.run(
        ["uv", "run", "python", "-c", script],
        cwd=PYTHON_CRATE,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        sys.stderr.write(completed.stderr)
        raise SystemExit(completed.returncode)
    return json.loads(completed.stdout.strip().splitlines()[-1])


def load_js_runtime_types() -> dict[str, str] | None:
    candidates = sorted((REPO_ROOT / "crates" / "officemd_js").glob("*.node"))
    if not candidates:
        return None

    script = r'''
const fs = require("node:fs");
const path = require("node:path");
const packageDir = path.resolve(process.cwd(), "crates", "officemd_js");
const entries = fs.readdirSync(packageDir).filter((name) => name.endsWith(".node"));
const bindingPath = path.join(packageDir, entries[0]);
const mod = require(bindingPath);
const docx = mod.createDocumentFromMarkdown("## Section: body\n\nHello\n", "docx");
const patched = mod.applyOoxmlPatchJson(
  docx,
  JSON.stringify({
    edits: [{ part: "word/document.xml", from: "Hello", to: "Hello runtime" }],
  }),
);
const payload = {
  apply_ooxml_patch_json: Buffer.isBuffer(patched) ? "bytes" : typeof patched,
  create_document_from_markdown: Buffer.isBuffer(docx) ? "bytes" : typeof docx,
  detect_format: typeof mod.detectFormat(docx),
};
console.log(JSON.stringify(payload));
'''
    completed = subprocess.run(
        ["node", "-e", script],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        sys.stderr.write(completed.stderr)
        raise SystemExit(completed.returncode)
    payload = json.loads(completed.stdout.strip().splitlines()[-1])
    if payload.get("detect_format") == "string":
        payload["detect_format"] = "str"
    return payload


def extract_js_rust_surface() -> dict[str, RustBindingFunction]:
    text = JS_RUST.read_text()
    surface: dict[str, RustBindingFunction] = {}
    for match in JS_RUST_FN_RE.finditer(text):
        name = match.group(1)
        params = []
        for raw_param in [chunk.strip() for chunk in match.group(2).split(",") if chunk.strip()]:
            params.append(camel_to_snake(raw_param.split(":", 1)[0].strip()))
        surface[name] = RustBindingFunction(
            name=name,
            params=tuple(params),
            return_type=normalize_js_type(match.group(3).strip()),
        )
    return surface


def parse_python_rust_params(fn_args: str) -> tuple[str, ...]:
    params: list[str] = []
    for raw_param in [chunk.strip() for chunk in fn_args.split(",") if chunk.strip()]:
        name = raw_param.split(":", 1)[0].strip()
        if name == "py":
            continue
        params.append(camel_to_snake(name))
    return tuple(params)


def extract_python_rust_surface() -> dict[str, RustBindingFunction]:
    text = PYTHON_RUST.read_text()
    surface: dict[str, RustBindingFunction] = {}
    for match in PY_RUST_ATTR_RE.finditer(text):
        name = match.group(2)
        if name.startswith("_"):
            continue
        params = parse_python_rust_params(match.group(3))
        surface[name] = RustBindingFunction(
            name=name,
            params=params,
            return_type=normalize_js_type(match.group(4).strip()),
        )
    return surface


def compare_surfaces(
    js_surface: dict[str, JsFunction],
    py_surface: dict[str, PythonFunction],
    js_rust_surface: dict[str, RustBindingFunction],
    py_rust_surface: dict[str, RustBindingFunction],
    py_runtime_types: dict[str, str],
    js_runtime_types: dict[str, str] | None,
) -> list[str]:
    problems: list[str] = []

    for label, surface in (
        ("JS", js_surface),
        ("Python", py_surface),
        ("JS Rust", js_rust_surface),
        ("Python Rust", py_rust_surface),
    ):
        missing = sorted(set(CONTRACT) - set(surface))
        extra = sorted(set(surface) - set(CONTRACT))
        if missing:
            problems.append(f"Missing in {label}: {', '.join(missing)}")
        if extra:
            problems.append(f"Unexpected in {label}: {', '.join(extra)}")

    for name, contract in CONTRACT.items():
        js_fn = js_surface.get(name)
        py_fn = py_surface.get(name)
        js_rust_fn = js_rust_surface.get(name)
        py_rust_fn = py_rust_surface.get(name)
        if None in (js_fn, py_fn, js_rust_fn, py_rust_fn):
            continue

        for label, fn_params in (
            ("JS", js_fn.params),
            ("Python", py_fn.params),
            ("JS Rust", js_rust_fn.params),
            ("Python Rust", py_rust_fn.params),
        ):
            if fn_params != contract.params:
                problems.append(f"{label} parameters for {name}: expected {list(contract.params)}, got {list(fn_params)}")

        expected_optional = frozenset(contract.python_defaults)
        if js_fn.optional_params != expected_optional:
            problems.append(
                f"JS optional parameters for {name}: expected {sorted(expected_optional)}, got {sorted(js_fn.optional_params)}"
            )

        for label, ret_type in (
            ("JS", js_fn.return_type),
            ("JS Rust", js_rust_fn.return_type),
            ("Python Rust", py_rust_fn.return_type),
        ):
            if ret_type != contract.js_return_type:
                problems.append(f"{label} return type for {name}: expected {contract.js_return_type}, got {ret_type}")

        for param, expected in contract.python_defaults.items():
            observed = py_fn.defaults.get(param, object())
            if observed != expected:
                problems.append(
                    f"Python default for {name}.{param}: expected {expected!r}, got {observed!r}"
                )

        required_params = [param for param in contract.params if param not in contract.python_defaults]
        for param in required_params:
            if param in py_fn.defaults:
                problems.append(f"Python parameter {name}.{param} should be required")

        if contract.python_runtime_type is not None:
            observed_type = py_runtime_types.get(name)
            if observed_type != contract.python_runtime_type:
                problems.append(
                    f"Python runtime type for {name}: expected {contract.python_runtime_type}, got {observed_type}"
                )

        if contract.js_runtime_type is not None and js_runtime_types is not None:
            observed_js_type = js_runtime_types.get(name)
            if observed_js_type != contract.js_runtime_type:
                problems.append(
                    f"JS runtime type for {name}: expected {contract.js_runtime_type}, got {observed_js_type}"
                )

    return problems


def print_surfaces(
    js_surface: dict[str, JsFunction],
    py_surface: dict[str, PythonFunction],
    js_rust_surface: dict[str, RustBindingFunction],
    py_rust_surface: dict[str, RustBindingFunction],
) -> None:
    print("## js")
    for name in sorted(js_surface):
        fn = js_surface[name]
        optionals = sorted(fn.optional_params)
        print(f"- {name}({', '.join(fn.params)}) -> {fn.return_type}" + (f" [optional: {', '.join(optionals)}]" if optionals else ""))
    print()
    print("## python")
    for name in sorted(py_surface):
        fn = py_surface[name]
        defaults = ", ".join(f"{k}={v!r}" for k, v in fn.defaults.items())
        print(f"- {name}({', '.join(fn.params)})" + (f" [defaults: {defaults}]" if defaults else ""))
    print()
    print("## js rust")
    for name in sorted(js_rust_surface):
        fn = js_rust_surface[name]
        print(f"- {name}({', '.join(fn.params)}) -> {fn.return_type}")
    print()
    print("## python rust")
    for name in sorted(py_rust_surface):
        fn = py_rust_surface[name]
        print(f"- {name}({', '.join(fn.params)}) -> {fn.return_type}")
    print()


def parse_args() -> Any:
    parser = argparse.ArgumentParser(description="Compare officemd JS/Python binding surfaces.")
    parser.add_argument("--print", action="store_true", help="Print normalized surfaces")
    parser.add_argument(
        "--require-js-runtime",
        action="store_true",
        help="Fail if a built JS native module is not available for runtime checks",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    js_surface = load_js_surface()
    py_surface = load_python_surface()
    js_rust_surface = extract_js_rust_surface()
    py_rust_surface = extract_python_rust_surface()
    py_runtime_types = load_python_runtime_types()
    js_runtime_types = load_js_runtime_types()

    if args.require_js_runtime and js_runtime_types is None:
        print("JS runtime checks requested but no built .node module was found", file=sys.stderr)
        return 1

    if args.print:
        print_surfaces(js_surface, py_surface, js_rust_surface, py_rust_surface)

    problems = compare_surfaces(
        js_surface,
        py_surface,
        js_rust_surface,
        py_rust_surface,
        py_runtime_types,
        js_runtime_types,
    )
    if problems:
        for problem in problems:
            print(problem, file=sys.stderr)
        return 1

    runtime_note = " with JS runtime checks" if js_runtime_types is not None else ""
    print(f"Binding surfaces match canonical contract across {len(CONTRACT)} functions{runtime_note}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
