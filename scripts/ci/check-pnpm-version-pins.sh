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
mise_file="mise.toml"

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

extract_mise_pnpm_version() {
  local file="$1"
  if [ ! -f "$file" ]; then
    printf '%s' ""
    return 0
  fi

  local line=""
  line="$(grep -E '^[[:space:]]*pnpm[[:space:]]*=' "$file" | head -n 1 || true)"
  if [ -z "$line" ]; then
    printf '%s' ""
    return 0
  fi

  local value="${line#*=}"
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
  value="${value#V}"
  printf '%s' "$value"
}

extract_workflow_env_pnpm_version() {
  local file="$1"
  local line=""
  # Ignore matches inside YAML block scalars (e.g. `run: |`) so non-semantic script content
  # can't satisfy or fail env indirection checks.
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

        if (line ~ /^[[:space:]]*PNPM_VERSION[[:space:]]*:/) {
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
  printf '%s' "$value"
}

workflow_has_pnpm_action_setup() {
  local file="$1"
  awk '
    function indent(s) {
      match(s, /^[ ]*/);
      return RLENGTH;
    }

    BEGIN {
      in_block = 0;
      block_indent = 0;
      block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";
      found = 0;
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

      if (line ~ /^[[:space:]]*-?[[:space:]]*uses:[[:space:]]*pnpm\/action-setup@/) {
        found = 1;
        exit;
      }

      if (is_block) {
        in_block = 1;
        block_indent = ind;
      }
    }

    END {
      exit found ? 0 : 1;
    }
  ' "$file"
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
    # Ignore YAML comment lines.
    local trimmed="$line"
    trimmed="${trimmed#"${trimmed%%[![:space:]]*}"}"
    case "$trimmed" in
      \#*) continue ;;
    esac

    if [[ "$line" =~ ^[[:space:]]*-?[[:space:]]*uses:[[:space:]]*pnpm/action-setup@ ]]; then
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
  done < <(
    # Skip YAML block scalar bodies (e.g. `run: |` scripts) so non-semantic text can't satisfy or fail these checks.
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
        if (line ~ block_re) {
          in_block = 1;
          block_indent = ind;
        }

        print raw;
      }
    ' "$file"
  )

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

check_workflow_corepack_pnpm_pins() {
  local file="$1"
  local env_pnpm_version=""
  env_pnpm_version="$(extract_workflow_env_pnpm_version "$file")"

  local found_any=0

  while IFS= read -r match; do
    [ -z "$match" ] && continue
    found_any=1

    # match format: "<line_no>:<content>"
    local line_no="${match%%:*}"
    local content="${match#*:}"

    # Ignore commented lines.
    local trimmed="$content"
    trimmed="${trimmed#"${trimmed%%[![:space:]]*}"}"
    case "$trimmed" in
      \#*) continue ;;
    esac

    # Supported forms:
    # - corepack prepare pnpm@9.0.0 --activate
    # - corepack prepare pnpm@${{ env.PNPM_VERSION }} --activate
    if [[ "$trimmed" =~ pnpm@([0-9]+\.[0-9]+\.[0-9]+) ]]; then
      local version="${BASH_REMATCH[1]}"
      if [ "$version" != "$expected_pnpm_version" ]; then
        echo "pnpm version pin mismatch in ${file}:${line_no} (corepack prepare):" >&2
        echo "  Expected pnpm version: ${expected_pnpm_version} (from ${package_json} packageManager)" >&2
        echo "  Found: pnpm@${version}" >&2
        echo "  Line: ${content}" >&2
        exit 1
      fi
      continue
    fi

    if [[ "$trimmed" =~ pnpm@\$\{\{[[:space:]]*env\.PNPM_VERSION[[:space:]]*\}\} ]]; then
      if [ -z "$env_pnpm_version" ]; then
        echo "pnpm version pin mismatch in ${file}:${line_no} (corepack prepare):" >&2
        echo "  Found: pnpm@\${{ env.PNPM_VERSION }} (but PNPM_VERSION is not set in the workflow env)" >&2
        echo "  Expected pnpm version: ${expected_pnpm_version} (from ${package_json} packageManager)" >&2
        echo "  Line: ${content}" >&2
        exit 1
      fi
      if [ "$env_pnpm_version" != "$expected_pnpm_version" ]; then
        echo "pnpm version pin mismatch in ${file}:${line_no} (corepack prepare):" >&2
        echo "  Expected pnpm version: ${expected_pnpm_version} (from ${package_json} packageManager)" >&2
        echo "  Found workflow PNPM_VERSION: ${env_pnpm_version}" >&2
        echo "  Line: ${content}" >&2
        exit 1
      fi
      continue
    fi

    echo "pnpm version pin mismatch in ${file}:${line_no} (corepack prepare):" >&2
    echo "  Expected pnpm version: ${expected_pnpm_version} (from ${package_json} packageManager)" >&2
    echo "  Found an unrecognized pnpm ref in: ${content}" >&2
    echo "  Fix: use corepack prepare pnpm@${expected_pnpm_version} --activate (or pnpm@\\\${{ env.PNPM_VERSION }} with PNPM_VERSION pinned)." >&2
    exit 1
  done < <(grep -n -E 'corepack[[:space:]]+prepare[[:space:]]+pnpm@' "$file" || true)

  if [ "$found_any" -eq 0 ]; then
    echo "No corepack prepare pnpm@... steps found in ${file} (expected at least one match when checking corepack pins)." >&2
    exit 1
  fi
}

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
for workflow in "${workflow_files[@]}"; do
  if workflow_has_pnpm_action_setup "$workflow"; then
    matched=1
    check_workflow_pnpm_pins "$workflow"
  fi
done

if [ "$matched" -eq 0 ]; then
  echo "No pnpm/action-setup steps found in any workflow under .github/workflows." >&2
  echo "If pnpm is not used in CI, delete this guard script and remove it from CI." >&2
  exit 1
fi

corepack_matched=0
for workflow in "${workflow_files[@]}"; do
  if grep -q -E 'corepack[[:space:]]+prepare[[:space:]]+pnpm@' "$workflow"; then
    corepack_matched=1
    check_workflow_corepack_pnpm_pins "$workflow"
  fi
done

mise_pnpm_version="$(extract_mise_pnpm_version "$mise_file")"
if [ -n "$mise_pnpm_version" ]; then
  if ! [[ "$mise_pnpm_version" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "mise.toml pnpm tool version must be patch-pinned (X.Y.Z); found pnpm=${mise_pnpm_version}" >&2
    exit 1
  fi
  if [ "$mise_pnpm_version" != "$expected_pnpm_version" ]; then
    echo "pnpm version pin mismatch between ${package_json} and ${mise_file}:" >&2
    echo "  ${package_json}: pnpm@${expected_pnpm_version}" >&2
    echo "  ${mise_file}: pnpm=${mise_pnpm_version}" >&2
    echo "" >&2
    echo "Fix: update ${mise_file} pnpm pin to ${expected_pnpm_version} (or update ${package_json} packageManager)." >&2
    exit 1
  fi
fi

echo "pnpm version pins match package.json (pnpm@${expected_pnpm_version})."
