#!/usr/bin/env bash
# Compare officemd vs markitdown vs docling conversion speed with hyperfine.
#
# Prerequisites:
#   brew install hyperfine   (or cargo install hyperfine)
#   cd crates/officemd_python
#   uv sync --group bench
#   uv run maturin develop --release
#
# Usage from repo root:
#   bash crates/tests/benchmarks/run_bench.sh [file ...]
#
# Defaults to showcase.docx, showcase.xlsx, showcase.pptx if no args given.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
BENCH_DIR="$REPO_ROOT/crates/tests/benchmarks"
PYTHON_DIR="$REPO_ROOT/crates/officemd_python"
DATA_DIR="$REPO_ROOT/examples/data"

if ! command -v hyperfine &>/dev/null; then
    echo "hyperfine not found. Install with: brew install hyperfine" >&2
    exit 1
fi

FILES=("$@")
if [ ${#FILES[@]} -eq 0 ]; then
    FILES=(
        "$DATA_DIR/showcase.docx"
        "$DATA_DIR/showcase.xlsx"
        "$DATA_DIR/showcase.pptx"
    )
fi

for FILE in "${FILES[@]}"; do
    if [ ! -f "$FILE" ]; then
        # Try resolving relative to DATA_DIR
        if [ -f "$DATA_DIR/$FILE" ]; then
            FILE="$DATA_DIR/$FILE"
        else
            echo "File not found: $FILE" >&2
            exit 1
        fi
    fi

    BASENAME="$(basename "$FILE")"
    echo ""
    echo "=== Benchmarking: $BASENAME ==="
    echo ""

    hyperfine \
        --warmup 2 \
        --export-markdown "/tmp/bench_${BASENAME%.*}.md" \
        --command-name "officemd" \
        "cd $PYTHON_DIR && uv run python $BENCH_DIR/bench_officemd.py $FILE > /dev/null" \
        --command-name "markitdown" \
        "uv run $BENCH_DIR/bench_markitdown.py $FILE > /dev/null" \
        --command-name "docling" \
        "uv run $BENCH_DIR/bench_docling.py $FILE > /dev/null"

    echo ""
    echo "Results saved to /tmp/bench_${BASENAME%.*}.md"
done
