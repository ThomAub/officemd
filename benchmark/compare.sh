#!/usr/bin/env bash
# Quality comparison: convert each corpus file with all tools and compare output.
#
# Usage from repo root:
#   bash benchmark/compare.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
CORPUS="$SCRIPT_DIR/corpus"
DATA_DIR="$REPO_ROOT/examples/data"
RESULTS="$SCRIPT_DIR/results"

# Files to skip from examples/data (edge-case fixtures, not meaningful benchmarks)
SKIP_PATTERN="^(trim_|encoding_heuristic|ocr_)"

# --- Tool detection ---

OFFICEMD=""
if [ -x "$REPO_ROOT/target/release/officemd" ]; then
    OFFICEMD="$REPO_ROOT/target/release/officemd"
elif command -v officemd &>/dev/null; then
    OFFICEMD="$(command -v officemd)"
else
    echo "ERROR: officemd not found. Build with: cargo build --release -p officemd_cli" >&2
    exit 1
fi

HAS_MARKIT=false
if command -v markit &>/dev/null; then
    HAS_MARKIT=true
fi

HAS_LITEPARSE=false
if command -v lit &>/dev/null; then
    HAS_LITEPARSE=true
fi

if [ ! -d "$CORPUS" ] || [ -z "$(ls -A "$CORPUS" 2>/dev/null)" ]; then
    echo "ERROR: corpus not found. Run: bash benchmark/setup-corpus.sh" >&2
    exit 1
fi

# --- Output directories ---

OUT_OFFICEMD="$RESULTS/officemd"
OUT_MARKITDOWN="$RESULTS/markitdown"
OUT_MARKIT="$RESULTS/markit"
OUT_LITEPARSE="$RESULTS/liteparse"

mkdir -p "$OUT_OFFICEMD" "$OUT_MARKITDOWN" "$OUT_MARKIT" "$OUT_LITEPARSE"
rm -f "$OUT_OFFICEMD"/*.md "$OUT_MARKITDOWN"/*.md "$OUT_MARKIT"/*.md "$OUT_LITEPARSE"/*.md "$RESULTS/quality.md"

file_id() {
    local file="$1"
    local rel="${file#$REPO_ROOT/}"
    rel="${rel//\//__}"
    rel="${rel// /_}"
    printf "%s" "$rel"
}

display_name() {
    local file="$1"
    printf "%s" "${file#$REPO_ROOT/}"
}

run_converter() {
    local output="$1"
    shift

    local tmp
    tmp="$(mktemp "$RESULTS/.compare.XXXXXX")"
    if "$@" > "$tmp" 2>/dev/null; then
        mv "$tmp" "$output"
    else
        rm -f "$tmp" "$output"
    fi
}

# Collect benchmark files from corpus + examples/data
FILES=()
for FILE in "$CORPUS"/*; do
    [ -f "$FILE" ] && FILES+=("$FILE")
done
if [ -d "$DATA_DIR" ]; then
    for FILE in "$DATA_DIR"/*; do
        [ -f "$FILE" ] || continue
        BASENAME="$(basename "$FILE")"
        [[ "$BASENAME" == *.md ]] && continue
        [[ "$BASENAME" =~ $SKIP_PATTERN ]] && continue
        FILES+=("$FILE")
    done
fi

# --- Convert all files ---

echo "=== Converting files (${#FILES[@]}) ==="

for FILE in "${FILES[@]}"; do
    BASENAME="$(basename "$FILE")"
    EXT="${BASENAME##*.}"
    FILE_ID="$(file_id "$FILE")"

    echo "  $BASENAME"

    # officemd
    run_converter "$OUT_OFFICEMD/${FILE_ID}.md" "$OFFICEMD" markdown "$FILE"

    # markitdown
    run_converter "$OUT_MARKITDOWN/${FILE_ID}.md" \
        uv run --with 'markitdown[all]==0.1.5' -- markitdown "$FILE"

    # markit (all formats)
    if [ "$HAS_MARKIT" = true ]; then
        run_converter "$OUT_MARKIT/${FILE_ID}.md" markit "$FILE"
    fi

    # liteparse (PDF only)
    if [ "$EXT" = "pdf" ] && [ "$HAS_LITEPARSE" = true ]; then
        run_converter "$OUT_LITEPARSE/${FILE_ID}.md" lit parse "$FILE"
    fi
done

# --- Build comparison table ---

echo ""
echo "=== Quality comparison - Output size ==="
echo ""

# Header
FMT="| %-30s | %14s | %14s | %14s | %14s |\n"
REPORT="$RESULTS/quality.md"

print_header() {
    printf "$FMT" "file" "officemd" "markitdown" "markit" "liteparse"
    printf "$FMT" "------------------------------" "--------------" "--------------" "--------------" "--------------"
}

{
    echo "# Quality comparison - Output size (bytes)"
    echo ""
    print_header
} > "$REPORT"

echo ""
print_header

for FILE in "${FILES[@]}"; do
    FILE_LABEL="$(display_name "$FILE")"
    FILE_ID="$(file_id "$FILE")"
    BASENAME="$(basename "$FILE")"
    EXT="${BASENAME##*.}"

    o_file="$OUT_OFFICEMD/${FILE_ID}.md"
    m_file="$OUT_MARKITDOWN/${FILE_ID}.md"
    k_file="$OUT_MARKIT/${FILE_ID}.md"
    l_file="$OUT_LITEPARSE/${FILE_ID}.md"

    o_val="-"; m_val="-"; k_val="-"; l_val="-"

    [ -s "$o_file" ] && o_val="$(wc -c < "$o_file" | tr -d ' ')"
    [ -s "$m_file" ] && m_val="$(wc -c < "$m_file" | tr -d ' ')"
    [ -s "$k_file" ] && k_val="$(wc -c < "$k_file" | tr -d ' ')"
    [ "$EXT" = "pdf" ] && [ -s "$l_file" ] && l_val="$(wc -c < "$l_file" | tr -d ' ')"

    ROW="$(printf "$FMT" "$FILE_LABEL" "$o_val" "$m_val" "$k_val" "$l_val")"

    echo -n "$ROW"
    echo "$ROW" >> "$REPORT"
done

echo ""
echo "Report saved to $REPORT"
