#!/usr/bin/env bash
set -euo pipefail

# Ensure TAURI_CLI_VERSION pins are consistent across workflows.
#
# Rationale:
# - We install tauri-cli (cargo tauri) in multiple workflows (CI, release, bundle-size, smoke).
# - Drifting pins can cause "CI green, release red" (or vice versa) due to toolchain differences.
# - This script fails fast when the pins diverge, so version bumps are an explicit PR with CI signal.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

workflows=(
  ".github/workflows/ci.yml"
  ".github/workflows/release.yml"
  ".github/workflows/desktop-bundle-size.yml"
  ".github/workflows/windows-arm64-smoke.yml"
)

extract_version() {
  local file="$1"
  local line=""
  line="$(grep -E '^[[:space:]]*TAURI_CLI_VERSION[[:space:]]*:' "$file" | head -n 1 || true)"
  if [ -z "$line" ]; then
    return 0
  fi
  local value="${line#*:}"
  value="${value%%#*}"
  value="${value#"${value%%[![:space:]]*}"}"
  value="${value%"${value##*[![:space:]]}"}"
  if [[ "$value" == \"*\" ]]; then
    value="${value#\"}"
    value="${value%\"}"
  elif [[ "$value" == \'*\' ]]; then
    value="${value#\'}"
    value="${value%\'}"
  fi
  value="${value#v}"
  printf '%s' "$value"
}

baseline=""
for wf in "${workflows[@]}"; do
  if [ ! -f "$wf" ]; then
    echo "Missing workflow file: ${wf}" >&2
    exit 1
  fi
  v="$(extract_version "$wf")"
  if [ -z "$v" ]; then
    echo "Failed to find TAURI_CLI_VERSION in ${wf}" >&2
    exit 1
  fi
  if ! [[ "$v" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "Expected TAURI_CLI_VERSION in ${wf} to be a pinned patch version (e.g. 2.9.5); got ${v}" >&2
    exit 1
  fi
  if [ -z "$baseline" ]; then
    baseline="$v"
  elif [ "$v" != "$baseline" ]; then
    echo "Tauri CLI version pin mismatch across workflows:" >&2
    echo "  Expected: ${baseline}" >&2
    for wf2 in "${workflows[@]}"; do
      echo "  - ${wf2}: $(extract_version "$wf2")" >&2
    done
    echo "" >&2
    echo "Fix: update TAURI_CLI_VERSION in the workflows above so they match exactly." >&2
    exit 1
  fi
done

echo "Tauri CLI version pins match (${baseline})."
