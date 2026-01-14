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
#
# If no workflow paths are provided, defaults to `.github/workflows/release.yml`.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

workflows=("$@")
if [ "${#workflows[@]}" -eq 0 ]; then
  workflows=(".github/workflows/release.yml")
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

for workflow in "${workflows[@]}"; do
  if [ ! -f "$workflow" ]; then
    echo "error: workflow file not found: ${workflow}" >&2
    exit 2
  fi

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
      echo "  Fix: add a trailing comment with the original upstream ref (e.g. # v4)." >&2
      fail=1
    fi
  fi
done <<<"$(grep -nHE '^[[:space:]]*-?[[:space:]]*uses:' "$workflow" || true)"

done

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "GitHub Actions pins: OK (all checked workflows use commit SHAs)."
