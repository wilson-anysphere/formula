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

extract_node_major() {
  local file="$1"
  local line=""
  # Ignore matches inside YAML block scalars (e.g. `run: |`) so non-semantic script content can't
  # satisfy or fail this guardrail.
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

        if (line ~ /^[[:space:]]*NODE_VERSION[[:space:]]*:/) {
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
  # Strip a leading "v" (Node version strings sometimes use vX.Y.Z).
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

extract_node_version_file_major() {
  local file="$1"
  if [ ! -f "$file" ]; then
    return 0
  fi
  local line=""
  # `.node-version` is typically a single line; ignore blank lines and comments.
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

workflow_uses_node_tooling() {
  local file="$1"
  # Only treat workflow YAML configuration and executable `run:` scripts as semantic:
  # - Ignore commented-out YAML.
  # - Ignore non-`run:` YAML block scalar bodies (e.g. env vars, action inputs) so strings like
  #   "actions/setup-node@" or "node -v" embedded in documentation can't accidentally classify a
  #   workflow as Node-consuming.
  awk '
    function indent(s) {
      match(s, /^[ ]*/);
      return RLENGTH;
    }

    BEGIN {
      in_block = 0;
      block_indent = 0;
      block_is_run = 0;
      block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";
      found = 0;
    }

    {
      raw = $0;
      sub(/\r$/, "", raw);
      ind = indent(raw);

      if (in_block) {
        # Blank/whitespace-only lines are always part of the scalar.
        if (raw ~ /^[[:space:]]*$/) next;
        if (ind > block_indent) {
          if (block_is_run) {
            trimmed = raw;
            sub(/^[[:space:]]*/, "", trimmed);
            # Ignore comment-only script lines.
            if (trimmed ~ /^#/) next;
            if (trimmed ~ /^node([[:space:]]|$)/) {
              found = 1;
              exit;
            }
          }
          next;
        }
        in_block = 0;
        block_is_run = 0;
      }

      trimmed = raw;
      sub(/^[[:space:]]*/, "", trimmed);
      if (trimmed ~ /^#/) next;

      line = raw;
      sub(/#.*/, "", line);
      is_block = (line ~ block_re);

      if (line ~ /^[[:space:]]*-?[[:space:]]*uses:[[:space:]]*actions\/setup-node@/) {
        found = 1;
        exit;
      }
      if (line ~ /^[[:space:]]*-?[[:space:]]*uses:[[:space:]]*pnpm\/action-setup@/) {
        found = 1;
        exit;
      }

      # Inline run steps: `run: node ...`.
      if (!is_block && line ~ /^[[:space:]]*-?[[:space:]]*run:[[:space:]]+/) {
        cmd = line;
        sub(/^[[:space:]]*-?[[:space:]]*run:[[:space:]]*/, "", cmd);
        if (cmd ~ /^node([[:space:]]|$)/) {
          found = 1;
          exit;
        }
      }

      # Track YAML block scalars so we can selectively scan only `run:` script bodies.
      if (is_block) {
        block_is_run = (line ~ /^[[:space:]]*-?[[:space:]]*run:[[:space:]]*[>|]/);
        in_block = 1;
        block_indent = ind;
      }
    }

    END {
      exit found ? 0 : 1;
    }
  ' "$file"
}

require_node_env_pins_match() {
  local file="$1"
  local expected_major="$2"

  local pins=""
  # Ignore matches inside YAML block scalars (e.g. `run: |`) so non-semantic script content can't
  # satisfy or fail this guardrail.
  pins="$(
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

        if (line ~ /^[[:space:]]*NODE_VERSION[[:space:]]*:/) {
          print raw;
        }

        if (is_block) {
          in_block = 1;
          block_indent = ind;
        }
      }
    ' "$file"
  )"
  if [ -z "$pins" ]; then
    echo "Node workflow pin check failed: Failed to find NODE_VERSION in ${file}" >&2
    exit 1
  fi

  while IFS= read -r line; do
    [ -z "$line" ] && continue
    local value="${line#*:}"
    value="${value%%#*}"
    value="$(extract_first_numeric_major "$value")"
    if [ -z "$value" ]; then
      echo "Node workflow pin check failed: Could not parse NODE_VERSION major from ${file}:" >&2
      echo "  ${line}" >&2
      exit 1
    fi
    if [ "$value" != "$expected_major" ]; then
      echo "Node major pin mismatch between CI and ${file}:" >&2
      echo "  ${ci_workflow}: NODE_VERSION=${expected_major}" >&2
      echo "  ${file}: NODE_VERSION=${value}" >&2
      echo "" >&2
      echo "Fix: update ${file} so NODE_VERSION matches CI/release." >&2
      exit 1
    fi
  done <<<"$pins"
}

require_env_pin_usage() {
  local file="$1"
  local fail=0
  local validated_any=0

  # Fail fast if a workflow uses node-version-file (that would bypass the explicit env pin).
  # Ignore matches inside YAML block scalars (e.g. `run: |` bodies) so arbitrary script content
  # can't satisfy or fail this guardrail.
  local file_pins
  file_pins="$(
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

        line = raw;
        sub(/#.*/, "", line);
        is_block = (line ~ block_re);

        # Ignore single-line run steps (inline shell snippets).
        if (!is_block && line ~ /^[[:space:]]*-?[[:space:]]*run:[[:space:]]+/) next;

        if (line ~ /node-version-file:/) {
          printf "%s:%d:%s\n", FILENAME, NR, raw;
        }

        if (is_block) {
          in_block = 1;
          block_indent = ind;
        }
      }
    ' "$file"
  )"

  if [ -n "$file_pins" ]; then
    echo "Node workflow pin check failed: ${file} uses node-version-file (unsupported in this repo)." >&2
    echo "Use: node-version: \${{ env.NODE_VERSION }} (and keep NODE_VERSION in sync across workflows)." >&2
    echo "" >&2
    echo "$file_pins" >&2
    exit 1
  fi

  # We expect workflows to reference the pinned Node major via env.NODE_VERSION.
  # (This makes it harder to accidentally update one job but not the others.)
  #
  # Ignore matches inside YAML block scalars (e.g. `run: |` script bodies) and inline
  # `run:` steps so script content can't satisfy or fail this guardrail.
  local matches
  matches="$(
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

        line = raw;
        sub(/#.*/, "", line);
        is_block = (line ~ block_re);

        # Ignore single-line run steps (inline shell snippets).
        if (!is_block && line ~ /^[[:space:]]*-?[[:space:]]*run:[[:space:]]+/) next;

        if (line ~ /node-version:/) {
          printf "%s:%d:%s\n", FILENAME, NR, raw;
        }

        if (is_block) {
          in_block = 1;
          block_indent = ind;
        }
      }
    ' "$file"
  )"

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

    validated_any=1

    if [[ "$content" != *"node-version: \${{ env.NODE_VERSION }}"* ]]; then
      echo "Node version pin mismatch in ${file}:" >&2
      echo "  Expected: node-version: \${{ env.NODE_VERSION }}" >&2
      echo "  Found:    ${match}" >&2
      echo >&2
      fail=1
    fi
  done <<<"$matches"

  if [ "$fail" -ne 0 ]; then
    exit 1
  fi

  if [ "$validated_any" -eq 0 ]; then
    echo "Node workflow pin check failed: ${file} contains only commented-out node-version pins." >&2
    echo "Expected actions/setup-node to use: node-version: \${{ env.NODE_VERSION }}" >&2
    exit 1
  fi
}

ci_node_major="$(extract_node_major "$ci_workflow")"
if [ -z "$ci_node_major" ]; then
  echo "Failed to find NODE_VERSION in ${ci_workflow}" >&2
  exit 1
fi

if ! [[ "$ci_node_major" =~ ^[0-9]+$ ]]; then
  echo "Expected NODE_VERSION in ${ci_workflow} to be a numeric Node major; got ${ci_node_major}" >&2
  exit 1
fi
# Discover workflows that depend on Node (setup-node, pnpm, or direct `node` invocation)
# and ensure they follow the same pinning rules as CI/release.
workflow_files=()
while IFS= read -r file; do
  [ -z "$file" ] && continue
  workflow_files+=("$file")
done < <(git ls-files .github/workflows | grep -E '\.(yml|yaml)$' || true)

node_workflows=()
for workflow in "${workflow_files[@]}"; do
  if workflow_uses_node_tooling "$workflow"; then
    node_workflows+=("$workflow")
  fi
done

# Always include CI + release workflows (they establish the canonical env pin for the repo).
node_workflows+=("$ci_workflow" "$release_workflow")
mapfile -t node_workflows < <(printf '%s\n' "${node_workflows[@]}" | sort -u)

if [ "${#node_workflows[@]}" -eq 0 ]; then
  echo "Node workflow pin check failed: no workflows appear to use Node tooling." >&2
  exit 1
fi

for workflow in "${node_workflows[@]}"; do
  require_node_env_pins_match "$workflow" "$ci_node_major"
  require_env_pin_usage "$workflow"
done

# Optional local tooling pins (keep local release builds aligned with CI).
nvmrc_major="$(extract_nvmrc_node_major ".nvmrc")"
node_version_major="$(extract_node_version_file_major ".node-version")"
mise_node_major="$(extract_mise_node_major "mise.toml")"

if [ -z "$nvmrc_major" ] && [ -z "$node_version_major" ] && [ -z "$mise_node_major" ]; then
  echo "Node workflow pin check failed: no repo-local Node version pin found." >&2
  echo "Add either:" >&2
  echo "  - .nvmrc (recommended), or" >&2
  echo "  - .node-version, or" >&2
  echo "  - mise.toml [tools] node = \"<major>\"" >&2
  exit 1
fi

if [ -n "$nvmrc_major" ] && [ "$nvmrc_major" != "$ci_node_major" ]; then
  echo "Node major pin mismatch between workflows and .nvmrc:" >&2
  echo "  workflows: NODE_VERSION=${ci_node_major}" >&2
  echo "  .nvmrc:     ${nvmrc_major}" >&2
  echo "" >&2
  echo "Fix: update .nvmrc or the workflows so they agree." >&2
  exit 1
fi

if [ -n "$node_version_major" ] && [ "$node_version_major" != "$ci_node_major" ]; then
  echo "Node major pin mismatch between workflows and .node-version:" >&2
  echo "  workflows: NODE_VERSION=${ci_node_major}" >&2
  echo "  .node-version: ${node_version_major}" >&2
  echo "" >&2
  echo "Fix: update .node-version or the workflows so they agree." >&2
  exit 1
fi

if [ -n "$mise_node_major" ] && [ "$mise_node_major" != "$ci_node_major" ]; then
  echo "Node major pin mismatch between workflows and mise.toml:" >&2
  echo "  workflows: NODE_VERSION=${ci_node_major}" >&2
  echo "  mise.toml:  node=${mise_node_major}" >&2
  echo "" >&2
  echo "Fix: update mise.toml or the workflows so they agree." >&2
  exit 1
fi

# Best-effort docs check: keep documented example workflows in sync with CI/release pins.
# These docs explicitly instruct readers to keep NODE_VERSION aligned with `.nvmrc` and CI.
docs_files=(
  "docs/13-testing-validation.md"
  "docs/16-performance-targets.md"
)
for doc in "${docs_files[@]}"; do
  if [ ! -f "$doc" ]; then
    continue
  fi
  found=0
  while IFS= read -r match; do
    [ -z "$match" ] && continue
    found=1

    line_no="${match%%:*}"
    content="${match#*:}"

    # Ignore commented lines.
    trimmed="$content"
    trimmed="${trimmed#"${trimmed%%[![:space:]]*}"}"
    case "$trimmed" in
      \#*) continue ;;
    esac

    value="${content#*:}"
    value="${value%%#*}"
    major="$(extract_first_numeric_major "$value")"
    if [ -z "$major" ]; then
      echo "Failed to parse NODE_VERSION in ${doc}:${line_no}:" >&2
      echo "  ${content}" >&2
      exit 1
    fi
    if [ "$major" != "$ci_node_major" ]; then
      echo "Node major pin mismatch between workflows and ${doc}:" >&2
      echo "  workflows: NODE_VERSION=${ci_node_major}" >&2
      echo "  ${doc}:${line_no}: NODE_VERSION=${major}" >&2
      echo "" >&2
      echo "Fix: update ${doc} so NODE_VERSION matches ${ci_node_major} (and keep it in sync with .nvmrc / workflows)." >&2
      exit 1
    fi
  done < <(grep -n -E '^[[:space:]]*NODE_VERSION[[:space:]]*:' "$doc" || true)
  # If the docs file doesn't mention NODE_VERSION, ignore; some downstream forks may not carry these docs.
  [ "$found" -eq 0 ] && true
done

echo "Node version pins match (NODE_VERSION=${ci_node_major})."
