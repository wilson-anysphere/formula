#!/usr/bin/env bash
set -euo pipefail

# Ensure CI and release workflows use the same pinned Node.js major version.
#
# Rationale:
# - We run different kinds of builds (web/desktop, tagged desktop releases).
# - A Node major mismatch between CI and release can cause "CI green, release red"
#   failures (or worse, subtly different bundles).
# - This script fails fast when the Node major pin diverges, so version bumps are
#   an explicit PR with CI signal.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

ci_workflow=".github/workflows/ci.yml"
release_workflow=".github/workflows/release.yml"
bundle_size_workflow=".github/workflows/desktop-bundle-size.yml"

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

require_env_pin_usage() {
  local file="$1"
  local fail=0

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
bundle_node_major="$(extract_node_major "$bundle_size_workflow")"

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

# Also ensure the workflows actually use the env pin consistently.
require_env_pin_usage "$ci_workflow"
require_env_pin_usage "$release_workflow"
require_env_pin_usage "$bundle_size_workflow"

echo "Node version pins match (NODE_VERSION=${ci_node_major})."
