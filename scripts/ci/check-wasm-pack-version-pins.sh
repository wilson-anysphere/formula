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

extract_version() {
  local file="$1"
  local line=""
  # Ignore matches inside YAML block scalars (e.g. `run: |`) so non-semantic script content
  # can't satisfy or fail this guardrail.
  line="$(
    awk '
      function indent(s) {
        match(s, /^[ ]*/);
        return RLENGTH;
      }

      BEGIN {
        in_block = 0;
        block_indent = 0;
        block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";
      }

      {
        raw = $0;
        sub(/\r$/, "", raw);
        ind = indent(raw);

        if (in_block) {
          if (raw ~ /^[[:space:]]*$/) next;
          if (ind > block_indent) next;
          in_block = 0;
        }

        trimmed = raw;
        sub(/^[[:space:]]*/, "", trimmed);
        if (trimmed ~ /^#/) next;

        line = raw;
        sub(/#.*/, "", line);
        is_block = (line ~ block_re);

        if (line ~ /^[[:space:]]*WASM_PACK_VERSION[[:space:]]*:/) {
          print raw;
          exit;
        }

        if (is_block) {
          in_block = 1;
          block_indent = ind;
        }
      }
    ' "$file"
  )"
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
  value="${value#[vV]}"
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
    echo "Expected WASM_PACK_VERSION in ${wf} to be a pinned patch version (e.g. 0.13.1); got ${v}" >&2
    exit 1
  fi
  if [ -z "$baseline" ]; then
    baseline="$v"
  elif [ "$v" != "$baseline" ]; then
    echo "wasm-pack version pin mismatch across workflows:" >&2
    echo "  Expected: ${baseline}" >&2
    for wf2 in "${workflow_files[@]}"; do
      wf2_v="$(extract_version "$wf2")"
      if [ -n "$wf2_v" ]; then
        echo "  - ${wf2}: ${wf2_v}" >&2
      fi
    done
    echo "" >&2
    echo "Fix: update WASM_PACK_VERSION in the workflows above so they match exactly." >&2
    exit 1
  fi
done

if [ "$matched" -eq 0 ]; then
  echo "No WASM_PACK_VERSION pins found in any workflow under .github/workflows." >&2
  echo "If wasm-pack is not used in CI, delete this guard script and remove it from CI." >&2
  exit 1
fi

echo "wasm-pack version pins match (${baseline})."
