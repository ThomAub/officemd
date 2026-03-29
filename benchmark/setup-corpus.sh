#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CORPUS="$SCRIPT_DIR/corpus"

mkdir -p "$CORPUS"

echo "=== Setting up benchmark corpus ==="

# --- Downloads (skip if already present) ---

download() {
    local name="$1" url="$2"
    if [ -f "$CORPUS/$name" ]; then
        echo "  skip $name (exists)"
        return
    fi
    echo "  downloading $name"
    curl -fsSL "$url" -o "$CORPUS/$name"
}

download "bitcoin-whitepaper.pdf" \
    "https://bitcoin.org/bitcoin.pdf"

download "us-constitution.pdf" \
    "https://constitutioncenter.org/media/files/constitution.pdf"

download "calibre-demo.docx" \
    "https://calibre-ebook.com/downloads/demos/demo.docx"

download "financial-sample.xlsx" \
    "https://go.microsoft.com/fwlink/?LinkID=521962"

download "titanic.csv" \
    "https://raw.githubusercontent.com/datasciencedojo/datasets/master/titanic.csv"

# --- Generated files ---

echo ""
echo "Generating synthetic files:"
uv run "$SCRIPT_DIR/generate-corpus.py" --output-dir "$CORPUS"

# --- Summary ---

echo ""
echo "=== Corpus ready ==="
ls -lhS "$CORPUS"
