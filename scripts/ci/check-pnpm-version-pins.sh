#!/usr/bin/env bash
set -euo pipefail

# Ensure CI and release workflows pin pnpm to the same patch version as the repo.
#
# Rationale:
# - pnpm patch versions can change lockfile resolution and install behavior.
# - We pin pnpm in package.json (`packageManager`) and in GitHub Actions workflows for
#   deterministic installs.
# - This script fails fast when the workflow pins drift from the repo's intended pnpm.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

package_json="package.json"
ci_workflow=".github/workflows/ci.yml"
release_workflow=".github/workflows/release.yml"
windows_arm64_smoke_workflow=".github/workflows/windows-arm64-smoke.yml"
bundle_size_workflow=".github/workflows/desktop-bundle-size.yml"
desktop_perf_platform_matrix_workflow=".github/workflows/desktop-perf-platform-matrix.yml"

extract_package_manager_pnpm_version() {
  local file="$1"
  local line=""
  line="$(grep -E '^[[:space:]]*"packageManager"[[:space:]]*:' "$file" | head -n 1 || true)"
  if [ -z "$line" ]; then
    return 0
  fi
  # Extract the quoted value.
  local value
  value="$(echo "$line" | sed -E 's/.*"packageManager"[[:space:]]*:[[:space:]]*"([^"]+)".*/\1/')"
  # Expect format pnpm@X.Y.Z
  if [[ "$value" =~ ^pnpm@([0-9]+\.[0-9]+\.[0-9]+)$ ]]; then
    printf '%s' "${BASH_REMATCH[1]}"
  else
    printf '%s' ""
  fi
}

expected_pnpm_version="$(extract_package_manager_pnpm_version "$package_json")"
if [ -z "$expected_pnpm_version" ]; then
  echo "Failed to parse pnpm version from ${package_json} packageManager field (expected pnpm@X.Y.Z)." >&2
  exit 1
fi

extract_workflow_env_pnpm_version() {
  local file="$1"
  local line=""
  line="$(grep -E '^[[:space:]]*PNPM_VERSION[[:space:]]*:' "$file" | head -n 1 || true)"
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
  printf '%s' "$value"
}

check_workflow_pnpm_pins() {
  local file="$1"
  local env_pnpm_version=""
  env_pnpm_version="$(extract_workflow_env_pnpm_version "$file")"
  local in_action=0
  local action_line=""
  local found_any=0
  local validated_any=0

  while IFS= read -r line; do
    if [[ "$line" =~ uses:[[:space:]]*pnpm/action-setup@ ]]; then
      in_action=1
      action_line="$line"
      found_any=1
      continue
    fi

    if [ "$in_action" -eq 1 ]; then
      # Look for the version pin under the action's "with:" block.
      if [[ "$line" =~ ^[[:space:]]*version:[[:space:]]*(.+)$ ]]; then
        local value="${BASH_REMATCH[1]}"
        # Strip trailing comment.
        value="${value%%#*}"
        # Trim whitespace.
        value="${value#"${value%%[![:space:]]*}"}"
        value="${value%"${value##*[![:space:]]}"}"
        # Strip quotes if present.
        if [[ "$value" == \"*\" ]]; then
          value="${value#\"}"
          value="${value%\"}"
        elif [[ "$value" == \'*\' ]]; then
          value="${value#\'}"
          value="${value%\'}"
        fi

        # Allow env indirection if it equals the expected version (rare in these workflows,
        # but used in some auxiliary workflows).
        if [[ "$value" == "\${{ env.PNPM_VERSION }}" || "$value" == "\${{ env.PNPM_VERSION }}"* ]]; then
          if [ -z "$env_pnpm_version" ]; then
            echo "pnpm version pin mismatch in ${file}:" >&2
            echo "  Found: version: \${{ env.PNPM_VERSION }} (but PNPM_VERSION is not set in the workflow env)" >&2
            echo "  Expected pnpm version: ${expected_pnpm_version} (from ${package_json} packageManager)" >&2
            exit 1
          fi
          if [ "$env_pnpm_version" != "$expected_pnpm_version" ]; then
            echo "pnpm version pin mismatch in ${file}:" >&2
            echo "  Expected pnpm version: ${expected_pnpm_version} (from ${package_json} packageManager)" >&2
            echo "  Found workflow PNPM_VERSION: ${env_pnpm_version}" >&2
            echo "  In action: ${action_line}" >&2
            exit 1
          fi
          validated_any=1
          in_action=0
          continue
        fi

        if [ "$value" != "$expected_pnpm_version" ]; then
          echo "pnpm version pin mismatch in ${file}:" >&2
          echo "  Expected pnpm version: ${expected_pnpm_version} (from ${package_json} packageManager)" >&2
          echo "  Found: version: ${value}" >&2
          echo "  In action: ${action_line}" >&2
          exit 1
        fi

        validated_any=1
        in_action=0
      fi
    fi
  done <"$file"

  if [ "$in_action" -eq 1 ]; then
    echo "pnpm/action-setup in ${file} is missing a 'version:' pin." >&2
    exit 1
  fi

  if [ "$found_any" -eq 0 ]; then
    echo "No pnpm/action-setup steps found in ${file} (expected at least one for JS workflows)." >&2
    exit 1
  fi

  if [ "$validated_any" -eq 0 ]; then
    echo "pnpm/action-setup steps found in ${file} but no version pin was validated." >&2
    echo "Ensure each pnpm/action-setup has a 'version:' key pinned to ${expected_pnpm_version} (directly or via env.PNPM_VERSION)." >&2
    exit 1
  fi
}

check_workflow_pnpm_pins "$ci_workflow"
check_workflow_pnpm_pins "$release_workflow"
check_workflow_pnpm_pins "$windows_arm64_smoke_workflow"
check_workflow_pnpm_pins "$bundle_size_workflow"
check_workflow_pnpm_pins "$desktop_perf_platform_matrix_workflow"

echo "pnpm version pins match package.json (pnpm@${expected_pnpm_version})."
