#!/usr/bin/env bash
set -euo pipefail

# Ensure WASM_PACK_VERSION pins are consistent across workflows.
#
# Rationale:
# - We install wasm-pack in CI, release builds, dry-run packaging, bundle-size checks,
#   perf runs, and Windows ARM64 smoke.
# - wasm-pack upgrades can change build output or break compatibility; we pin for determinism.
# - This script fails fast when the pins diverge.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

workflows=(
  ".github/workflows/ci.yml"
  ".github/workflows/release.yml"
  ".github/workflows/desktop-bundle-dry-run.yml"
  ".github/workflows/desktop-bundle-size.yml"
  ".github/workflows/desktop-perf-platform-matrix.yml"
  ".github/workflows/perf.yml"
  ".github/workflows/windows-arm64-smoke.yml"
)

extract_version() {
  local file="$1"
  local line=""
  line="$(grep -E '^[[:space:]]*WASM_PACK_VERSION[[:space:]]*:' "$file" | head -n 1 || true)"
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
    echo "Failed to find WASM_PACK_VERSION in ${wf}" >&2
    exit 1
  fi
  if ! [[ "$v" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "Expected WASM_PACK_VERSION in ${wf} to be a pinned patch version (e.g. 0.13.1); got ${v}" >&2
    exit 1
  fi
  if [ -z "$baseline" ]; then
    baseline="$v"
  elif [ "$v" != "$baseline" ]; then
    echo "wasm-pack version pin mismatch across workflows:" >&2
    echo "  Expected: ${baseline}" >&2
    for wf2 in "${workflows[@]}"; do
      echo "  - ${wf2}: $(extract_version "$wf2")" >&2
    done
    echo "" >&2
    echo "Fix: update WASM_PACK_VERSION in the workflows above so they match exactly." >&2
    exit 1
  fi
done

echo "wasm-pack version pins match (${baseline})."
