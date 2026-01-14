#!/usr/bin/env bash
set -euo pipefail

# Ensure CI/release workflows install the same Rust toolchain as rust-toolchain.toml.
#
# Rationale:
# - rust-toolchain.toml is the single "intended" Rust version for the repo.
# - GitHub Actions workflows install Rust via dtolnay/rust-toolchain pinned to a commit SHA
#   (supply-chain hardening). When pinning by SHA, upstream recommends selecting a SHA that is in
#   master history to avoid garbage collection of old branch-only commits.
# - Because the action is pinned by SHA, workflows must pass the Rust version via `with: toolchain:`.
# - This script fails fast if workflows drift from rust-toolchain.toml, making Rust upgrades an
#   explicit PR/change with CI validation.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

toolchain_file="rust-toolchain.toml"
if [ ! -f "$toolchain_file" ]; then
  echo "Missing ${toolchain_file} at repo root" >&2
  exit 1
fi

channel="$(
  awk -F'"' '
    $1 ~ /^[[:space:]]*channel[[:space:]]*=[[:space:]]*/ {
      print $2;
      exit;
    }
  ' "$toolchain_file"
)"

if [ -z "$channel" ]; then
  echo "Failed to parse toolchain channel from ${toolchain_file}" >&2
  exit 1
fi

if ! [[ "$channel" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
  echo "rust-toolchain.toml must pin an explicit Rust release (e.g. 1.92.0); found channel=${channel}" >&2
  exit 1
fi

fail=0

workflow_files=()
while IFS= read -r file; do
  [ -z "$file" ] && continue
  workflow_files+=("$file")
done < <(git ls-files .github/workflows | grep -E '\.(yml|yaml)$' || true)

if [ "${#workflow_files[@]}" -eq 0 ]; then
  echo "No workflow files found under .github/workflows" >&2
  exit 2
fi

for workflow in "${workflow_files[@]}"; do
  # Parse each workflow and validate every dtolnay/rust-toolchain step declares a matching toolchain.
  awk -v workflow="$workflow" -v expected="$channel" '
    function ltrim(s) { sub(/^[[:space:]]+/, "", s); return s }
    function rtrim(s) { sub(/[[:space:]]+$/, "", s); return s }
    function trim(s) { return rtrim(ltrim(s)) }
    function strip_quotes(s) {
      s = trim(s)
      if (s ~ /^"/) { sub(/^"/, "", s); sub(/"$/, "", s) }
      else if (s ~ /^'\''/) { sub(/^'\''/, "", s); sub(/'\''$/, "", s) }
      return s
    }
    function normalize_toolchain(s) {
      s = strip_quotes(s)
      sub(/^v/, "", s)
      return s
    }
    function finalize_step() {
      if (!step_dtolnay) return
      if (step_toolchain == "") {
        printf("Rust toolchain workflow step is missing an explicit toolchain input:\n") > "/dev/stderr"
        printf("  rust-toolchain.toml channel = %s\n", expected) > "/dev/stderr"
        printf("  %s:%d uses dtolnay/rust-toolchain but no `with: toolchain:` was found\n", workflow, step_uses_line) > "/dev/stderr"
        printf("  Fix: add:\n    with:\n      toolchain: %s\n\n", expected) > "/dev/stderr"
        fail = 1
        return
      }
      if (step_toolchain !~ /^[0-9]+\.[0-9]+\.[0-9]+$/) {
        printf("Rust toolchain must be patch-pinned (X.Y.Z):\n") > "/dev/stderr"
        printf("  rust-toolchain.toml channel = %s\n", expected) > "/dev/stderr"
        printf("  %s:%d toolchain: %s\n", workflow, step_toolchain_line, step_toolchain) > "/dev/stderr"
        printf("  Fix: set toolchain: %s\n\n", expected) > "/dev/stderr"
        fail = 1
        return
      }
      if (step_toolchain != expected) {
        printf("Rust toolchain pin mismatch:\n") > "/dev/stderr"
        printf("  rust-toolchain.toml channel = %s\n", expected) > "/dev/stderr"
        printf("  %s:%d toolchain: %s\n", workflow, step_toolchain_line, step_toolchain) > "/dev/stderr"
        printf("  Fix: update the workflow to use toolchain: %s (or update rust-toolchain.toml).\n\n", expected) > "/dev/stderr"
        fail = 1
      }
    }

    BEGIN {
      fail = 0
      in_block = 0
      block_indent = 0
      in_steps = 0
      steps_indent = 0
      step_item_indent = 0
      in_with = 0
      with_indent = 0
      step_dtolnay = 0
      step_toolchain = ""
      step_uses_line = 0
      step_toolchain_line = 0
    }

    {
      line = $0
      sub(/\r$/, "", line)
      trimmed = trim(line)
      match(line, /^[ ]*/)
      indent = RLENGTH

      # Skip YAML literal/folded block scalar contents (run: |, restore-keys: |, etc).
      if (in_block) {
        if (trimmed != "" && indent <= block_indent) {
          in_block = 0
        } else {
          next
        }
      }
      if (!in_block && trimmed ~ /^[^#].*:[[:space:]]*[|>][+-]?[[:space:]]*($|#)/) {
        in_block = 1
        block_indent = indent
        next
      }

      # Enter a steps: block.
      if (!in_steps) {
        if (trimmed ~ /^steps:[[:space:]]*($|#)/) {
          in_steps = 1
          steps_indent = indent
          step_item_indent = steps_indent + 2
          in_with = 0
          with_indent = 0
          step_dtolnay = 0
          step_toolchain = ""
          step_uses_line = 0
          step_toolchain_line = 0
        }
        next
      }

      # Leaving a steps: block.
      if (trimmed != "" && trimmed !~ /^#/) {
        if (indent <= steps_indent) {
          finalize_step()
          in_steps = 0
          in_with = 0
          step_dtolnay = 0
          step_toolchain = ""
          next
        }
      }

      # New step item (only when indent matches the list under steps:).
      if (indent == step_item_indent && trimmed ~ /^-[[:space:]]/) {
        finalize_step()
        in_with = 0
        with_indent = 0
        step_dtolnay = 0
        step_toolchain = ""
        step_uses_line = 0
        step_toolchain_line = 0

        # Inline "- uses: ..." form.
        step_line = trimmed
        sub(/^-+/, "", step_line)
        step_line = trim(step_line)
        if (step_line ~ /^uses:[[:space:]]*/) {
          value = step_line
          sub(/^uses:[[:space:]]*/, "", value)
          value = trim(value)
          # Strip YAML comments.
          sub(/[[:space:]]+#.*/, "", value)
          value = strip_quotes(value)
          if (value ~ /^dtolnay\/rust-toolchain@/) {
            step_dtolnay = 1
            step_uses_line = NR
          }
        }
        next
      }

      # Within a step item.
      if (indent < step_item_indent) {
        next
      }

      # Track with: block scope.
      if (in_with && trimmed != "" && indent <= with_indent) {
        in_with = 0
        with_indent = 0
      }
      if (trimmed ~ /^with:[[:space:]]*($|#)/) {
        in_with = 1
        with_indent = indent
        next
      }

      # Detect uses: dtolnay/rust-toolchain inside a named step.
      if (trimmed ~ /^uses:[[:space:]]*/) {
        value = trimmed
        sub(/^uses:[[:space:]]*/, "", value)
        value = trim(value)
        sub(/[[:space:]]+#.*/, "", value)
        value = strip_quotes(value)
        if (value ~ /^dtolnay\/rust-toolchain@/) {
          step_dtolnay = 1
          step_uses_line = NR
        }
      }

      if (step_dtolnay && in_with && trimmed ~ /^toolchain:[[:space:]]*/) {
        value = trimmed
        sub(/^toolchain:[[:space:]]*/, "", value)
        value = trim(value)
        sub(/[[:space:]]+#.*/, "", value)
        value = normalize_toolchain(value)
        step_toolchain = value
        step_toolchain_line = NR
      }
    }

    END {
      if (in_steps) {
        finalize_step()
      }
      exit fail
    }
  ' "$workflow" || fail=1
done

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "Rust toolchain pins match rust-toolchain.toml (channel=${channel})."
