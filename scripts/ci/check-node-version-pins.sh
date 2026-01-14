#!/usr/bin/env bash
set -euo pipefail

# Ensure all Node-consuming workflows (and local tooling pins) use the same
# pinned Node.js major version.
#
# Rationale:
# - We run different kinds of builds (web/desktop, tagged desktop releases, perf,
#   security scans, etc).
# - A Node major mismatch between workflows can cause "CI green, release red"
#   failures (or worse, subtly different artifacts).
# - This script fails fast when the Node major pin diverges, so version bumps are
#   an explicit, coordinated PR with CI signal.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

ci_workflow=".github/workflows/ci.yml"
release_workflow=".github/workflows/release.yml"
dry_run_workflow=".github/workflows/desktop-bundle-dry-run.yml"
bundle_size_workflow=".github/workflows/desktop-bundle-size.yml"
desktop_perf_platform_matrix_workflow=".github/workflows/desktop-perf-platform-matrix.yml"
collab_perf_workflow=".github/workflows/collab-perf.yml"
perf_workflow=".github/workflows/perf.yml"
security_workflow=".github/workflows/security.yml"
windows_arm64_smoke_workflow=".github/workflows/windows-arm64-smoke.yml"

extract_node_major() {
  local file="$1"
  local line=""
  line="$(grep -E '^[[:space:]]*NODE_VERSION[[:space:]]*:' "$file" | head -n 1 || true)"
  if [ -z "$line" ]; then
    return 0
  fi

  # Remove key + colon.
  local value="${line#*:}"
  # Strip trailing comments.
  value="${value%%#*}"
  # Trim whitespace.
  value="${value#"${value%%[![:space:]]*}"}"
  value="${value%"${value##*[![:space:]]}"}"
  # Strip surrounding quotes if present.
  if [[ "$value" == \"*\" ]]; then
    value="${value#\"}"
    value="${value%\"}"
  elif [[ "$value" == \'*\' ]]; then
    value="${value#\'}"
    value="${value%\'}"
  fi

  printf '%s' "$value"
}

extract_first_numeric_major() {
  local raw="$1"
  # Trim whitespace.
  raw="${raw#"${raw%%[![:space:]]*}"}"
  raw="${raw%"${raw##*[![:space:]]}"}"
  # Strip surrounding quotes if present.
  if [[ "$raw" == \"*\" ]]; then
    raw="${raw#\"}"
    raw="${raw%\"}"
  elif [[ "$raw" == \'*\' ]]; then
    raw="${raw#\'}"
    raw="${raw%\'}"
  fi
  # Strip a leading "v" (Node version strings sometimes use v22.0.0).
  raw="${raw#v}"
  raw="${raw#V}"
  # Capture leading digits as the major.
  if [[ "$raw" =~ ^([0-9]+) ]]; then
    printf '%s' "${BASH_REMATCH[1]}"
    return 0
  fi
  printf '%s' ""
}

extract_nvmrc_node_major() {
  local file="$1"
  if [ ! -f "$file" ]; then
    return 0
  fi
  local line=""
  # `.nvmrc` is typically a single line; ignore blank lines and comments.
  line="$(grep -E '^[[:space:]]*[^#[:space:]]' "$file" | head -n 1 || true)"
  extract_first_numeric_major "$line"
}

extract_mise_node_major() {
  local file="$1"
  if [ ! -f "$file" ]; then
    return 0
  fi
  local line=""
  line="$(grep -E '^[[:space:]]*node[[:space:]]*=' "$file" | head -n 1 || true)"
  if [ -z "$line" ]; then
    printf '%s' ""
    return 0
  fi
  # Remove key + '='.
  local value="${line#*=}"
  # Strip trailing comments.
  value="${value%%#*}"
  extract_first_numeric_major "$value"
}

require_env_pin_usage() {
  local file="$1"
  local fail=0

  # Fail fast if a workflow uses node-version-file (that would bypass the explicit env pin).
  set +e
  local file_pins
  file_pins="$(git grep -n "node-version-file:" -- "$file" 2>/dev/null)"
  local file_pin_status=$?
  set -e
  if [ "$file_pin_status" -eq 2 ]; then
    echo "git grep failed while scanning ${file} for node-version-file pins" >&2
    exit 2
  fi
  if [ -n "$file_pins" ]; then
    echo "Node workflow pin check failed: ${file} uses node-version-file (unsupported in this repo)." >&2
    echo "Use: node-version: \${{ env.NODE_VERSION }} (and keep NODE_VERSION in sync between CI and release)." >&2
    echo "" >&2
    echo "$file_pins" >&2
    exit 1
  fi

  # We expect workflows to reference the pinned Node major via env.NODE_VERSION.
  # (This makes it harder to accidentally update one job but not the others.)
  set +e
  local matches
  matches="$(git grep -n "node-version:" -- "$file" 2>/dev/null)"
  local status=$?
  set -e

  if [ "$status" -eq 2 ]; then
    echo "git grep failed while scanning ${file}" >&2
    exit 2
  fi

  if [ -z "$matches" ]; then
    echo "Node workflow pin check failed: no node-version pins found in ${file}." >&2
    echo "Expected actions/setup-node to use: node-version: \${{ env.NODE_VERSION }}" >&2
    exit 1
  fi

  while IFS= read -r match; do
    [ -z "$match" ] && continue
    # match format: "<file>:<line>:<content>"
    local content="${match#*:*:}"
    # Ignore commented-out lines.
    local trimmed="$content"
    trimmed="${trimmed#"${trimmed%%[![:space:]]*}"}"
    case "$trimmed" in
      \#*) continue ;;
    esac

    if [[ "$content" != *"node-version: \${{ env.NODE_VERSION }}"* ]]; then
      echo "Node version pin mismatch in ${file}:"
      echo "  Expected: node-version: \${{ env.NODE_VERSION }}"
      echo "  Found:    ${match}"
      echo
      fail=1
    fi
  done <<<"$matches"

  if [ "$fail" -ne 0 ]; then
    exit 1
  fi
}

ci_node_major="$(extract_node_major "$ci_workflow")"
release_node_major="$(extract_node_major "$release_workflow")"
dry_run_node_major="$(extract_node_major "$dry_run_workflow")"
bundle_node_major="$(extract_node_major "$bundle_size_workflow")"
desktop_perf_platform_matrix_node_major="$(extract_node_major "$desktop_perf_platform_matrix_workflow")"
collab_perf_node_major="$(extract_node_major "$collab_perf_workflow")"
perf_node_major="$(extract_node_major "$perf_workflow")"
security_node_major="$(extract_node_major "$security_workflow")"
smoke_node_major="$(extract_node_major "$windows_arm64_smoke_workflow")"

if [ -z "$ci_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${ci_workflow}" >&2
  exit 1
fi
if [ -z "$release_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${release_workflow}" >&2
  exit 1
fi
if [ -z "$bundle_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${bundle_size_workflow}" >&2
  exit 1
fi
if [ -z "$desktop_perf_platform_matrix_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${desktop_perf_platform_matrix_workflow}" >&2
  exit 1
fi
if [ -z "$collab_perf_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${collab_perf_workflow}" >&2
  exit 1
fi
if [ -z "$perf_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${perf_workflow}" >&2
  exit 1
fi
if [ -z "$security_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${security_workflow}" >&2
  exit 1
fi
if [ -z "$smoke_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${windows_arm64_smoke_workflow}" >&2
  exit 1
fi
if [ -z "$dry_run_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${dry_run_workflow}" >&2
  exit 1
fi

if ! [[ "$ci_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${ci_workflow} to be a Node major (e.g. 22); got ${ci_node_major}" >&2
  exit 1
fi
if ! [[ "$release_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${release_workflow} to be a Node major (e.g. 22); got ${release_node_major}" >&2
  exit 1
fi
if ! [[ "$bundle_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${bundle_size_workflow} to be a Node major (e.g. 22); got ${bundle_node_major}" >&2
  exit 1
fi
if ! [[ "$desktop_perf_platform_matrix_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${desktop_perf_platform_matrix_workflow} to be a Node major (e.g. 22); got ${desktop_perf_platform_matrix_node_major}" >&2
  exit 1
fi
if ! [[ "$collab_perf_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${collab_perf_workflow} to be a Node major (e.g. 22); got ${collab_perf_node_major}" >&2
  exit 1
fi
if ! [[ "$perf_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${perf_workflow} to be a Node major (e.g. 22); got ${perf_node_major}" >&2
  exit 1
fi
if ! [[ "$security_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${security_workflow} to be a Node major (e.g. 22); got ${security_node_major}" >&2
  exit 1
fi
if ! [[ "$smoke_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${windows_arm64_smoke_workflow} to be a Node major (e.g. 22); got ${smoke_node_major}" >&2
  exit 1
fi
if ! [[ "$dry_run_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${dry_run_workflow} to be a Node major (e.g. 22); got ${dry_run_node_major}" >&2
  exit 1
fi

if [ "$ci_node_major" != "$release_node_major" ]; then
  echo "Node major pin mismatch between CI and release workflows:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${release_workflow}: NODE_VERSION=${release_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi
if [ "$ci_node_major" != "$bundle_node_major" ]; then
  echo "Node major pin mismatch between CI and desktop bundle-size workflows:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${bundle_size_workflow}: NODE_VERSION=${bundle_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi
if [ "$ci_node_major" != "$desktop_perf_platform_matrix_node_major" ]; then
  echo "Node major pin mismatch between CI and desktop perf platform matrix workflow:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${desktop_perf_platform_matrix_workflow}: NODE_VERSION=${desktop_perf_platform_matrix_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi
if [ "$ci_node_major" != "$collab_perf_node_major" ]; then
  echo "Node major pin mismatch between CI and collab perf workflow:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${collab_perf_workflow}: NODE_VERSION=${collab_perf_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi
if [ "$ci_node_major" != "$perf_node_major" ]; then
  echo "Node major pin mismatch between CI and perf workflow:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${perf_workflow}: NODE_VERSION=${perf_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi
if [ "$ci_node_major" != "$security_node_major" ]; then
  echo "Node major pin mismatch between CI and security workflow:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${security_workflow}: NODE_VERSION=${security_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi
if [ "$ci_node_major" != "$smoke_node_major" ]; then
  echo "Node major pin mismatch between CI and Windows ARM64 smoke workflows:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${windows_arm64_smoke_workflow}: NODE_VERSION=${smoke_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi
if [ "$ci_node_major" != "$dry_run_node_major" ]; then
  echo "Node major pin mismatch between CI and desktop dry-run workflow:" >&2
  echo "  ${ci_workflow}: NODE_VERSION=${ci_node_major}" >&2
  echo "  ${dry_run_workflow}: NODE_VERSION=${dry_run_node_major}" >&2
  echo "" >&2
  echo "Fix: update one of the workflows so both use the same Node major." >&2
  exit 1
fi

# Also ensure the workflows actually use the env pin consistently.
require_env_pin_usage "$ci_workflow"
require_env_pin_usage "$release_workflow"
require_env_pin_usage "$dry_run_workflow"
require_env_pin_usage "$bundle_size_workflow"
require_env_pin_usage "$desktop_perf_platform_matrix_workflow"
require_env_pin_usage "$collab_perf_workflow"
require_env_pin_usage "$perf_workflow"
require_env_pin_usage "$security_workflow"
require_env_pin_usage "$windows_arm64_smoke_workflow"

# Optional local tooling pins (keep local release builds aligned with CI).
nvmrc_major="$(extract_nvmrc_node_major ".nvmrc")"
if [ -n "$nvmrc_major" ] && [ "$nvmrc_major" != "$ci_node_major" ]; then
  echo "Node major pin mismatch between workflows and .nvmrc:" >&2
  echo "  workflows: NODE_VERSION=${ci_node_major}" >&2
  echo "  .nvmrc:     ${nvmrc_major}" >&2
  echo "" >&2
  echo "Fix: update .nvmrc or the workflows so they agree." >&2
  exit 1
fi

mise_node_major="$(extract_mise_node_major "mise.toml")"
if [ -n "$mise_node_major" ] && [ "$mise_node_major" != "$ci_node_major" ]; then
  echo "Node major pin mismatch between workflows and mise.toml:" >&2
  echo "  workflows: NODE_VERSION=${ci_node_major}" >&2
  echo "  mise.toml:  node=${mise_node_major}" >&2
  echo "" >&2
  echo "Fix: update mise.toml or the workflows so they agree." >&2
  exit 1
fi

echo "Node version pins match (NODE_VERSION=${ci_node_major})."
