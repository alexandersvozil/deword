#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
INPUT="$ROOT/input/board-memo-input.docx"
PATCH="$ROOT/patch.json"

cp "$INPUT" "$ROOT/candidate/board-memo-output.docx"
deword patch "$ROOT/candidate/board-memo-output.docx" -p "$PATCH"

echo "Built $ROOT/candidate/board-memo-output.docx"
