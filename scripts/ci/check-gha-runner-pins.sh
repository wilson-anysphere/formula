#!/usr/bin/env bash
set -euo pipefail

# Guardrail: avoid using GitHub Actions' moving `*-latest` hosted runner labels
# in the Desktop Release workflow.
#
# Why:
# `macos-latest`, `windows-latest`, and `ubuntu-latest` aliases periodically move
# to newer OS images. That can introduce unexpected build breakages for tagged
# releases and make it harder to reproduce historical artifacts.
#
# Update guidance:
# When bumping runner pins (e.g., macos-14 -> macos-15), validate the release
# workflow on the new images first, then update the pinned versions + this guard
# if needed.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

workflow="${1:-.github/workflows/release.yml}"
if [ ! -f "$workflow" ]; then
  echo "error: workflow file not found: ${workflow}" >&2
  exit 2
fi

# Ignore matches that only appear in comments. (We still match `runs-on: ubuntu-latest # ...`.)
matches="$(grep -nE '^[^#]*\\b(macos|windows|ubuntu)-latest\\b' "$workflow" || true)"
if [ -n "$matches" ]; then
  echo "error: ${workflow} uses forbidden GitHub Actions runner '*-latest' labels." >&2
  echo "Pin runner images instead (e.g., macos-14, windows-2022, ubuntu-24.04) to avoid alias drift." >&2
  echo "" >&2
  echo "Found:" >&2
  echo "$matches" >&2
  exit 1
fi
