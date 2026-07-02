#!/usr/bin/env bash
# Benchmark officemd vs markitdown vs markit vs liteparse (PDF only) using hyperfine.
#
# Prerequisites:
#   brew install hyperfine
#   cargo build --release -p officemd_cli
#   npm i -g markit-ai                (optional, for markit comparison)
#   npm i -g @llamaindex/liteparse    (optional, for PDF comparison)
#   bash benchmark/setup-corpus.sh
#   uv run benchmark/setup-gdp-pdf-corpus.py  (optional GDP.pdf sample)
#
# Usage from repo root:
#   bash benchmark/run.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
CORPUS="$SCRIPT_DIR/corpus"
DATA_DIR="$REPO_ROOT/examples/data"
RESULTS="$SCRIPT_DIR/results"

# Files to skip from examples/data (edge-case fixtures, not meaningful benchmarks)
SKIP_PATTERN="^(trim_|encoding_heuristic|ocr_)"

# --- Tool detection ---

if ! command -v hyperfine &>/dev/null; then
    echo "ERROR: hyperfine not found. Install with: brew install hyperfine" >&2
    exit 1
fi

# Find officemd binary
OFFICEMD=""
if [ -x "$REPO_ROOT/target/release/officemd" ]; then
    OFFICEMD="$REPO_ROOT/target/release/officemd"
elif command -v officemd &>/dev/null; then
    OFFICEMD="$(command -v officemd)"
else
    echo "ERROR: officemd not found. Build with: cargo build --release -p officemd_cli" >&2
    exit 1
fi
echo "officemd: $OFFICEMD"

# Check markitdown
if ! uv run --with 'markitdown[all]==0.1.5' -- markitdown --help &>/dev/null; then
    echo "ERROR: markitdown not available via uv" >&2
    exit 1
fi
echo "markitdown: via uv run --with 'markitdown[all]==0.1.5'"

# Check markit (optional)
HAS_MARKIT=false
if command -v markit &>/dev/null; then
    HAS_MARKIT=true
    echo "markit: $(command -v markit)"
else
    echo "markit: not found (skipping)"
    echo "  Install with: npm i -g markit-ai"
fi

# Check liteparse (optional, PDF only)
HAS_LITEPARSE=false
if command -v lit &>/dev/null; then
    HAS_LITEPARSE=true
    echo "liteparse: $(command -v lit)"
else
    echo "liteparse: not found (skipping for PDF files)"
    echo "  Install with: npm i -g @llamaindex/liteparse"
fi

# Check corpus
if [ ! -d "$CORPUS" ] || [ -z "$(ls -A "$CORPUS" 2>/dev/null)" ]; then
    echo "ERROR: corpus not found. Run: bash benchmark/setup-corpus.sh" >&2
    exit 1
fi

mkdir -p "$RESULTS"

# Collect benchmark files from corpus + examples/data
FILES=()
add_corpus_file() {
    local file="$1"
    local basename
    basename="$(basename "$file")"
    case "$basename" in
        *.csv|*.CSV|*.docx|*.DOCX|*.pdf|*.PDF|*.pptx|*.PPTX|*.xlsx|*.XLSX)
            FILES+=("$file")
            ;;
    esac
}

while IFS= read -r FILE; do
    [ -f "$FILE" ] && add_corpus_file "$FILE"
done < <(find "$CORPUS" -type f | sort)
if [ -d "$DATA_DIR" ]; then
    for FILE in "$DATA_DIR"/*; do
        [ -f "$FILE" ] || continue
        BASENAME="$(basename "$FILE")"
        # Skip non-document files and edge-case fixtures
        [[ "$BASENAME" == *.md ]] && continue
        [[ "$BASENAME" =~ $SKIP_PATTERN ]] && continue
        FILES+=("$FILE")
    done
fi

file_id() {
    local file="$1"
    local rel="${file#$CORPUS/}"
    if [ "$rel" = "$file" ]; then
        rel="$(basename "$file")"
    fi
    rel="${rel//\//__}"
    rel="${rel// /_}"
    printf "%s" "$rel"
}

echo ""
echo "=== Running benchmarks (${#FILES[@]} files) ==="

for FILE in "${FILES[@]}"; do
    BASENAME="$(basename "$FILE")"
    EXT="${BASENAME##*.}"
    FILE_ID="$(file_id "$FILE")"

    echo ""
    echo "--- $BASENAME ---"

    CMDS=(
        --command-name "officemd"
        "$OFFICEMD markdown $FILE"
        --command-name "markitdown"
        "uv run --with 'markitdown[all]==0.1.5' -- markitdown $FILE"
    )

    # Add markit (all formats)
    if [ "$HAS_MARKIT" = true ]; then
        CMDS+=(
            --command-name "markit"
            "markit $FILE"
        )
    fi

    # Add liteparse (PDF only)
    if [ "$EXT" = "pdf" ] && [ "$HAS_LITEPARSE" = true ]; then
        CMDS+=(
            --command-name "liteparse"
            "lit parse $FILE"
        )
    fi

    hyperfine \
        --warmup 3 \
        --min-runs 5 \
        --export-json "$RESULTS/${FILE_ID}.json" \
        --export-markdown "$RESULTS/${FILE_ID}.md" \
        "${CMDS[@]}"
done

# --- Generate report ---

echo ""
echo "=== Generating report ==="
uv run "$SCRIPT_DIR/report.py" --results-dir "$RESULTS" --output "$RESULTS/summary.md"

echo ""
echo "Results saved to $RESULTS/summary.md"
