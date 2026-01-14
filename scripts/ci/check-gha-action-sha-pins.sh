#!/usr/bin/env bash
set -euo pipefail

# Guardrail: ensure GitHub Actions in critical workflows are pinned to immutable
# commit SHAs (not floating tags/branches).
#
# Why:
# - Prevents compromised/malicious action updates from affecting signed release artifacts.
# - Makes releases reproducible by ensuring the workflow always runs the same action code.
#
# Usage:
#   bash scripts/ci/check-gha-action-sha-pins.sh [workflow.yml ...]
#   bash scripts/ci/check-gha-action-sha-pins.sh .github/workflows
#
# If no workflow paths are provided, defaults to `.github/workflows/release.yml`.

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

fail=0

trim() {
  local s="$1"
  # ltrim
  s="${s#"${s%%[![:space:]]*}"}"
  # rtrim
  s="${s%"${s##*[![:space:]]}"}"
  printf '%s' "$s"
}

extract_uses_lines() {
  local workflow="$1"
  # Print lines in the format: file:line:content
  #
  # Important: avoid matching `uses:` strings inside YAML block scalars (e.g. `run: |` scripts)
  # to prevent false positives.
  awk '
    function indent(line) {
      match(line, /^[ \t]*/);
      return RLENGTH;
    }
    {
      raw=$0;
      sub(/\r$/, "", raw);

      ind=indent(raw);

      trimmed=raw;
      sub(/^[ \t]*/, "", trimmed);

      # Best-effort strip trailing comments for block-scalar detection.
      code=trimmed;
      sub(/[ \t]+#.*/, "", code);

      if (in_block) {
        # YAML block scalars end when indentation returns to the block-start level (or less).
        if (trimmed != "" && ind <= block_indent) {
          in_block=0;
        } else {
          next;
        }
      }

      # Detect the start of a block scalar (e.g. `run: |`, `script: >-`, `run: |2`, `run: |2-`).
      # GitHub YAML parser accepts indentation/chomping indicators in either order, so treat any
      # combination of digits/+/- after `|`/`>` as a block scalar start.
      if (code ~ /^[^:#]+:[ \t]*[>|][0-9+-]*[ \t]*$/) {
        in_block=1;
        block_indent=ind;
      }

      # Match YAML keys (including job-level reusable workflows) and step entries.
      if (trimmed ~ /^-?[ \t]*uses:[ \t]*/) {
        print FILENAME ":" NR ":" raw;
      }
    }
  ' "$workflow"
}

for workflow in "${workflows[@]}"; do

while IFS= read -r match; do
  [ -z "$match" ] && continue

  file="${match%%:*}"
  rest="${match#*:}"
  line_no="${rest%%:*}"
  line="${match#*:*:}"

  value="${line#*uses:}"
  comment=""
  if [[ "$value" == *"#"* ]]; then
    comment="${value#*#}"
    comment="$(trim "$comment")"
  fi
  value="${value%%#*}" # strip comment
  value="$(trim "$value")"
  value="${value%$'\r'}"

  # Strip surrounding quotes if present.
  if [[ "$value" == \"*\" ]]; then
    value="${value#\"}"
    value="${value%\"}"
  elif [[ "$value" == \'*\' ]]; then
    value="${value#\'}"
    value="${value%\'}"
  fi

  # Local actions are safe to reference by path.
  if [[ "$value" == ./* ]]; then
    continue
  fi

  # Container actions referenced by image digest are out of scope here.
  # (There are none in our release workflows today.)
  if [[ "$value" == docker://* ]]; then
    continue
  fi

  if [[ "$value" != *"@"* ]]; then
    echo "error: ${file}:${line_no} has an action/workflow 'uses:' without an @ref:" >&2
    echo "  uses: ${value}" >&2
    fail=1
    continue
  fi

  ref="${value##*@}"
  if ! [[ "$ref" =~ ^[0-9a-fA-F]{40}$ ]]; then
    echo "error: ${file}:${line_no} must pin actions/workflows to a full 40-char commit SHA:" >&2
    echo "  Found: uses: ${value}" >&2
    echo "  Fix: replace the ref after '@' with an immutable commit SHA, and keep a trailing comment with the original tag (e.g. # v4)." >&2
    fail=1
  else
    # Maintainability: require a trailing comment indicating the original tag/branch.
    # Example:
    #   uses: actions/checkout@<sha> # v4.3.1
    if [ -z "$comment" ]; then
      echo "error: ${file}:${line_no} pins an action to a SHA but is missing a trailing version comment:" >&2
      echo "  Found: uses: ${value}" >&2
      echo "  Fix: add a trailing comment with the original upstream ref (e.g. # v4.3.1 or # master)." >&2
      fail=1
    else
      # Dependabot uses the trailing comment to infer the intended version channel/tag for SHA pins.
      # Validate that the comment begins with something tag-like (semver-ish) or a common branch name.
      token="${comment%%[[:space:]]*}"
      if ! [[ "$token" =~ ^(v?[0-9]+(\.[0-9]+){0,3}|master|main|stable|beta|nightly)$ ]]; then
        echo "error: ${file}:${line_no} action pin comment should start with an upstream ref (e.g. v4.3.1):" >&2
        echo "  Found: uses: ${value} # ${comment}" >&2
        echo "  Fix: start the comment with the upstream tag/branch (e.g. # v4 or # v4.3.1)." >&2
        fail=1
      fi
    fi
  fi
done <<<"$(extract_uses_lines "$workflow" || true)"

done

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "GitHub Actions pins: OK (all checked workflows use commit SHAs)."
