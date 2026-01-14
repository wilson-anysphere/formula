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

# Ignore matches that only appear in comments and block scalars (e.g. within a `run: |` script body).
# We still match `runs-on: ubuntu-latest # ...`.
matches="$(
  awk '
    function indent(s) {
      match(s, /^[ ]*/);
      return RLENGTH;
    }

    BEGIN {
      in_block = 0;
      block_indent = 0;
      re = "(macos-latest|windows-latest|ubuntu-latest)";
    }

    {
      raw = $0;
      ind = indent(raw);

      if (in_block) {
        # Blank/whitespace-only lines can appear inside block scalars with any indentation.
        # Treat them as part of the scalar so we do not accidentally stop skipping early.
        if (raw ~ /^[[:space:]]*$/) {
          next;
        }
        if (ind > block_indent) {
          next;
        }
        in_block = 0;
      }

      # Strip YAML comments (best-effort; avoids tripping on documentation).
      line = raw;
      sub(/#.*/, "", line);

      if (line ~ re) {
        printf "%d:%s\n", NR, raw;
      }

      # Detect YAML block scalars (e.g. `run: |` / `releaseBody: >-`) so we can skip
      # their content lines.
      # YAML allows both orders for chomping/indentation indicators (e.g. `|2-`, `|-2`).
      if (line ~ /:[[:space:]]*[>|][0-9+-]*[[:space:]]*$/) {
        in_block = 1;
        block_indent = ind;
      }
    }
  ' "$workflow"
)"
if [ -n "$matches" ]; then
  echo "error: ${workflow} uses forbidden GitHub Actions runner '*-latest' labels." >&2
  echo "Pin runner images instead (e.g., macos-14, windows-2022, ubuntu-24.04) to avoid alias drift." >&2
  echo "" >&2
  echo "Found:" >&2
  echo "$matches" >&2
  exit 1
fi
