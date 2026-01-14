#!/usr/bin/env bash
set -euo pipefail

# Guardrail: avoid using GitHub Actions' moving `*-latest` hosted runner labels
# in workflow configuration (`runs-on`, matrices, etc).
#
# Why:
# `macos-latest`, `windows-latest`, and `ubuntu-latest` aliases periodically move
# to newer OS images. That can introduce unexpected build breakages for tagged
# releases and make it harder to reproduce historical artifacts.
#
# Usage:
#   bash scripts/ci/check-gha-runner-pins.sh [workflow.yml ...]
#   bash scripts/ci/check-gha-runner-pins.sh .github/workflows
#
# Update guidance:
# When bumping runner pins (e.g., macos-14 -> macos-15), validate the release
# workflow (and any other affected workflows) on the new images first, then
# update the pinned versions.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

inputs=("$@")
if [ "${#inputs[@]}" -eq 0 ]; then
  inputs=(".github/workflows/release.yml")
fi

workflows=()
for path in "${inputs[@]}"; do
  if [ -d "$path" ]; then
    # Expand a directory into the workflow files it contains.
    while IFS= read -r file; do
      [ -z "$file" ] && continue
      workflows+=("$file")
    done < <(find "$path" -maxdepth 4 -type f \( -name '*.yml' -o -name '*.yaml' \) | sort)
    continue
  fi

  if [ -f "$path" ]; then
    workflows+=("$path")
    continue
  fi

  echo "error: workflow file/directory not found: ${path}" >&2
  exit 2
done

if [ "${#workflows[@]}" -eq 0 ]; then
  echo "error: no workflow files found from inputs: ${inputs[*]}" >&2
  exit 2
fi

# Ignore matches that only appear in comments and block scalars (e.g. within a `run: |` script body).
# We still match `runs-on: ubuntu-latest # ...`.
all_matches=""
for workflow in "${workflows[@]}"; do
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
        block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";
      }

      {
        raw = $0;
        sub(/\r$/, "", raw);
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

        is_block = (line ~ block_re);

        # Ignore single-line `run:` steps. We only want to catch runner label usage
        # in the workflow configuration, not within inline shell snippets.
        if (!is_block && line ~ /^[[:space:]]*-?[[:space:]]*run:[[:space:]]+/) {
          next;
        }

        if (line ~ re) {
          printf "%s:%d:%s\n", FILENAME, NR, raw;
        }

        # Detect YAML block scalars (e.g. `run: |` / `releaseBody: >-`) so we can skip
        # their content lines.
        # YAML allows both orders for chomping/indentation indicators (e.g. `|2-`, `|-2`).
        if (is_block) {
          in_block = 1;
          block_indent = ind;
        }
      }
    ' "$workflow"
  )"

  if [ -n "$matches" ]; then
    if [ -n "$all_matches" ]; then
      all_matches+=$'\n'
    fi
    all_matches+="$matches"
  fi
done

if [ -n "$all_matches" ]; then
  echo "error: found forbidden GitHub Actions runner '*-latest' labels in workflow configuration." >&2
  echo "Pin runner images instead (e.g., macos-14, windows-2022, ubuntu-24.04) to avoid alias drift." >&2
  echo "See https://github.com/actions/runner-images for supported images and deprecation notices." >&2
  echo "" >&2
  echo "Found:" >&2
  echo "$all_matches" >&2
  exit 1
fi
