#!/usr/bin/env bash
set -euo pipefail

# Ensure TAURI_CLI_VERSION pins are consistent across workflows.
#
# Rationale:
# - We install tauri-cli (cargo tauri) in multiple workflows (CI, release, dry-run, bundle-size, smoke).
# - Drifting pins can cause "CI green, release red" (or vice versa) due to toolchain differences.
# - This script fails fast when the pins diverge, so version bumps are an explicit PR with CI signal.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

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
workflow_files=()
while IFS= read -r file; do
  [ -z "$file" ] && continue
  workflow_files+=("$file")
done < <(git ls-files .github/workflows | grep -E '\.(yml|yaml)$' || true)

if [ "${#workflow_files[@]}" -eq 0 ]; then
  echo "No workflow files found under .github/workflows" >&2
  exit 2
fi

matched=0
for wf in "${workflow_files[@]}"; do
  v="$(extract_version "$wf")"
  if [ -z "$v" ]; then
    continue
  fi
  matched=1
  if ! [[ "$v" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "Expected TAURI_CLI_VERSION in ${wf} to be a pinned patch version (e.g. 2.9.5); got ${v}" >&2
    exit 1
  fi
  if [ -z "$baseline" ]; then
    baseline="$v"
  elif [ "$v" != "$baseline" ]; then
    echo "Tauri CLI version pin mismatch across workflows:" >&2
    echo "  Expected: ${baseline}" >&2
    for wf2 in "${workflow_files[@]}"; do
      wf2_v="$(extract_version "$wf2")"
      if [ -n "$wf2_v" ]; then
        echo "  - ${wf2}: ${wf2_v}" >&2
      fi
    done
    echo "" >&2
    echo "Fix: update TAURI_CLI_VERSION in the workflows above so they match exactly." >&2
    exit 1
  fi
done

if [ "$matched" -eq 0 ]; then
  echo "No TAURI_CLI_VERSION pins found in any workflow under .github/workflows." >&2
  echo "If tauri-cli is not used in CI, delete this guard script and remove it from CI." >&2
  exit 1
fi

echo "Tauri CLI version pins match (${baseline})."
