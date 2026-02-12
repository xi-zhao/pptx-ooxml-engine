#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT_DIR="${1:-$ROOT_DIR/examples/generated}"

mkdir -p "$OUT_DIR"
PYTHONPATH="$ROOT_DIR/src" python -m pptx_ooxml_engine.examples_runner --output-dir "$OUT_DIR"

echo "Generated example PPTX files in: $OUT_DIR"
