# OfficeMD ↔ ParseBench integration

A drop-in [ParseBench][parsebench] `PARSE` provider that invokes the local
OfficeMD CLI on a PDF and normalizes the JSON document into ParseBench's
`ParseOutput`. Shipped with a helper script to materialize a non-OCR slice of
a document corpus using OfficeMD's own classifier.

The first benchmark pass targets **text-based, non-OCR PDFs** so results are
comparable with ParseBench's local baselines (`pypdf_baseline`,
`pymupdf_text`).

[parsebench]: https://github.com/run-llama/ParseBench

## Layout

```
integrations/parsebench/
├── pyproject.toml
├── src/parse_bench/
│   └── inference/
│       ├── providers/parse/officemd_local.py   # Provider (drop-in module path)
│       └── pipelines/officemd_pipelines.py     # register_officemd_pipelines(register_fn)
├── scripts/classify_non_ocr_pdfs.py            # Materialize the non-OCR slice
└── tests/test_officemd_local.py                # Provider unit tests
```

The provider module lives under the `parse_bench.inference.providers.parse`
package path so it can be imported as-is from a ParseBench checkout. The
package is distributed as `parse-bench-officemd` and cohabits the
`parse_bench` namespace via a Hatch wheel.

## Installation

Install into the same virtualenv as your ParseBench checkout. From the
OfficeMD repo root:

```sh
uv pip install -e integrations/parsebench
# or:
pip install -e integrations/parsebench
```

Then patch the ParseBench pipeline registry to register the new pipeline.
Edit `src/parse_bench/inference/pipelines/parse.py` (in the ParseBench
checkout) to call `register_officemd_pipelines`:

```python
from parse_bench.inference.pipelines.officemd_pipelines import (
    register_officemd_pipelines,
)

def register_parse_pipelines(register_fn):
    # ...existing registrations...
    register_officemd_pipelines(register_fn)
```

No upstream changes are required beyond this single call; the provider
auto-registers via `@register_provider("officemd_local")` when the pipeline
module is imported.

## Running

Point the provider at an OfficeMD checkout (either via env var or pipeline
config) and run the pipeline:

```sh
export OFFICEMD_REPO_ROOT=/path/to/officemd
uv run parse-bench pipelines                          # confirms officemd_local is listed
uv run parse-bench run officemd_local --test --group text_content
uv run parse-bench run officemd_local --test --group text_formatting
```

By default the provider invokes `cargo run --release -p officemd_cli --
stream <pdf> --output-format json --pretty` from `repo_root`. To use a
prebuilt binary instead:

```python
PipelineSpec(
    pipeline_name="officemd_local_binary",
    provider_name="officemd_local",
    product_type=ProductType.PARSE,
    config={
        "cargo_run": False,
        "binary": "/path/to/target/release/officemd",
        "extra_args": ["--no-headers-footers"],
    },
)
```

### Provider config reference

| Key | Type | Default | Description |
|---|---|---|---|
| `cargo_run` | bool | `True` | Invoke `cargo run` from `repo_root` |
| `repo_root` | str | `$OFFICEMD_REPO_ROOT` | Absolute path to the workspace root (required when `cargo_run=True`) |
| `cargo_profile` | str | `"release"` | `release`, `dev`, or a custom profile name |
| `binary` | str | `None` | Prebuilt binary path (required when `cargo_run=False`) |
| `extra_args` | list[str] | `[]` | Extra CLI flags appended after the input path |
| `timeout_seconds` | float | `600` | Per-file subprocess timeout |

## Normalization rules

- `pdf.pages[].number` → `PageIR.page_index = number - 1`
- `pdf.pages[].markdown` → `PageIR.markdown`
- Document `markdown` is page markdown joined with a single blank line
- `layout_pages` is intentionally left empty in v1; layout attribution is
  not wired through OfficeMD yet
- The full OfficeMD JSON document (including `pdf.diagnostics`) is preserved
  in `raw_output["document"]` for downstream analysis

## Slicing the non-OCR benchmark subset

After the dataset documents are present, run the classifier to select PDFs
that OfficeMD considers `TextBased` with no pages requiring OCR:

```sh
uv run integrations/parsebench/scripts/classify_non_ocr_pdfs.py \
    --input-dir ~/parsebench/data/documents \
    --report-jsonl non_ocr_report.jsonl \
    --manifest non_ocr_manifest.txt
```

Outputs:

- `non_ocr_report.jsonl` — one row per PDF with classification, confidence,
  page count, `pages_needing_ocr`, and a `non_ocr` boolean. Keep this
  alongside the benchmark run for later drill-down.
- `non_ocr_manifest.txt` — the filtered list of qualifying PDFs; feed it to
  ParseBench (or a small wrapper) to restrict the run rather than editing
  the upstream dataset files.

A PDF is selected when the classification is `TextBased` **and**
`pages_needing_ocr` is empty, matching OfficeMD's own definition of
"pure text-based".

## Comparison workflow

Run the same non-OCR slice through OfficeMD and the local baselines:

```sh
uv run parse-bench run officemd_local  --test --group text_content
uv run parse-bench run pypdf_baseline  --test --group text_content
uv run parse-bench run pymupdf_text    --test --group text_content
```

Repeat for `text_formatting` and `table`. `chart` is intentionally skipped
in the first pass — OfficeMD emits text markdown but no chart-specific
structured payload yet, so a chart comparison would degenerate into a text
comparison. Revisit once chart metadata is part of the PDF JSON payload.

## Testing

```sh
uv pip install -e integrations/parsebench[dev]
uv run pytest integrations/parsebench/tests -q
```

The unit tests stub `subprocess.run` and do not require `cargo`, a real PDF,
or a live OfficeMD checkout. They do require `parse_bench` to be installed
in the environment; otherwise they are skipped via `pytest.importorskip`.

## Assumptions and non-goals

- Invocation uses `cargo run` by default so benchmark runs always exercise
  the current tree. Switch to `binary` mode for stable runs.
- ParseBench layout attribution is **not** wired up in v1; the goal is to
  improve non-OCR parse quality, not overlay reconstruction.
- Cleaner page-boundary or formatting semantics than the current CLI JSON
  exposes should be added to the OfficeMD PDF payload, **not** papered over
  with ParseBench-specific post-processing.
