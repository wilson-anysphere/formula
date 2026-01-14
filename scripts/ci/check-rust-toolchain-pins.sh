#!/usr/bin/env bash
set -euo pipefail

# Ensure CI/release workflows install the same Rust toolchain as rust-toolchain.toml.
#
# This script also validates any optional local tooling pin in `mise.toml` (if present) so
# developers using mise install the same Rust release as CI/release.
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
  # Parse each workflow and validate:
  # - Each job that runs Rust tooling installs the pinned toolchain (dtolnay/rust-toolchain),
  #   and does so *before* any cargo/rustup/rustc usage.
  # - Every dtolnay/rust-toolchain step declares a matching `with: toolchain: <pinned>` input.
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
    function note_rust_use(snippet) {
      if (steps_saw_dtolnay) return
      if (steps_rust_before_toolchain_line != 0) return
      steps_rust_before_toolchain_line = NR
      steps_rust_before_toolchain_snippet = snippet
    }
    function line_looks_like_rust_tooling(s) {
      # True for:
      # - `cargo ...`, `rustup ...`, `rustc ...` (including `cargo +nightly ...`)
      # - invocations of the repo cargo wrapper
      # - scripts that invoke cargo (CI preflight helpers)
      # - but not `.cargo/` paths or `cargo-foo` cache keys.
      if (s ~ /cargo_agent\.sh/) return 1
      if (s ~ /check-cargo-lock-reproducible\.sh/) return 1
      if (s ~ /check-tauri-permissions\.mjs/) return 1
      if (s ~ /build-formula-wasm-node\.mjs/) return 1
      if (s ~ /packages\/engine\/scripts\/build-wasm\.mjs/) return 1
      if (s ~ /generate-function-catalog\.js/) return 1
      if (s ~ /excel-oracle\/compat_gate\.py/) return 1
      if (s ~ /tools\.corpus\.triage/) return 1
      if (s ~ /tools\/corpus\/triage\.py/) return 1
      if (s ~ /scripts\/security\/ci\.sh/) return 1
      if (s ~ /desktop_size_report\.py/) return 1
      if (s ~ /desktop_binary_size_report\.py/) return 1
      # pnpm wrapper scripts that transitively run Rust tooling (e.g. desktop prebuild hooks).
      if (s ~ /pnpm[[:space:]]+build:desktop([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+dev:desktop([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+test:e2e([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+test:node$/) return 1
      if (s ~ /pnpm[[:space:]]+build:wasm([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+-w[[:space:]]+build:wasm([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+-C[[:space:]]+apps\/desktop[[:space:]]+build([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+-C[[:space:]]+apps\/desktop[[:space:]]+dev([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+-C[[:space:]]+apps\/desktop[[:space:]]+test:e2e([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+-C[[:space:]]+packages\/engine[[:space:]]+build:wasm([[:space:]]|$)/) return 1
      if (s ~ /pnpm[[:space:]]+benchmark([[:space:]]|$)/) return 1
      if (s ~ /(^|[[:space:];&|()])cargo([[:space:]]|$)/) return 1
      if (s ~ /(^|[[:space:];&|()])wasm-pack([[:space:]]|$)/) return 1
      if (s ~ /(^|[[:space:];&|()])rustup([[:space:]]|$)/) return 1
      if (s ~ /(^|[[:space:];&|()])rustc([[:space:]]|$)/) return 1
      return 0
    }
    function finalize_steps_block() {
      if (!in_steps) return
      if (steps_rust_before_toolchain_line == 0) return
      printf("Rust toolchain pin check failed:\n") > "/dev/stderr"
      printf("  rust-toolchain.toml channel = %s\n", expected) > "/dev/stderr"
      printf("  %s:%d uses Rust tooling before installing the pinned toolchain\n", workflow, steps_rust_before_toolchain_line) > "/dev/stderr"
      if (steps_rust_before_toolchain_snippet != "") {
        printf("  First Rust usage in this job: %s\n", steps_rust_before_toolchain_snippet) > "/dev/stderr"
      }
      printf("  Fix: add a dtolnay/rust-toolchain step (pinned to a commit SHA) *before* any Rust tooling in this job, with:\n    with:\n      toolchain: %s\n\n", expected) > "/dev/stderr"
      fail = 1
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
      block_is_run = 0
      in_steps = 0
      steps_indent = 0
      step_item_indent = 0
      in_with = 0
      with_indent = 0
      steps_saw_dtolnay = 0
      steps_rust_before_toolchain_line = 0
      steps_rust_before_toolchain_snippet = ""
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
          block_is_run = 0
        } else {
          # Only treat `run: |` blocks as executable shell scripts; other block scalars like
          # `with: script: |` (github-script) and `with: restore-keys: |` contain structured data
          # and may mention cargo/rustup in strings without actually invoking Rust.
          if (in_steps && block_is_run) {
            # Ignore comment-only lines in scripts (they are not executed).
            if (trimmed ~ /^#/) next
            if (line_looks_like_rust_tooling(trimmed)) {
              note_rust_use(trimmed)
            }
          }
          next
        }
      }
      if (!in_block && trimmed ~ /^[^#].*:[[:space:]]*[|>][0-9+-]*[[:space:]]*($|#)/) {
        code = trimmed
        # Best-effort strip trailing YAML comments.
        sub(/[[:space:]]+#.*/, "", code)
        block_is_run = (code ~ /^run:[[:space:]]*[|>]/ || code ~ /^-[[:space:]]*run:[[:space:]]*[|>]/)
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
          steps_saw_dtolnay = 0
          steps_rust_before_toolchain_line = 0
          steps_rust_before_toolchain_snippet = ""
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
          finalize_steps_block()
          in_steps = 0
          in_with = 0
          steps_saw_dtolnay = 0
          steps_rust_before_toolchain_line = 0
          steps_rust_before_toolchain_snippet = ""
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

        # Inline "- uses: ..." / "- run: ..." forms.
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
            steps_saw_dtolnay = 1
          } else if (value ~ /^tauri-apps\/tauri-action@/ || value ~ /^Swatinem\/rust-cache@/) {
            note_rust_use("uses: " value)
          }
        } else if (step_line ~ /^run:[[:space:]]*/) {
          value = step_line
          sub(/^run:[[:space:]]*/, "", value)
          value = trim(value)
          sub(/[[:space:]]+#.*/, "", value)
          value = strip_quotes(value)
          if (line_looks_like_rust_tooling(value)) {
            note_rust_use("run: " value)
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
          steps_saw_dtolnay = 1
        } else if (value ~ /^tauri-apps\/tauri-action@/ || value ~ /^Swatinem\/rust-cache@/) {
          note_rust_use("uses: " value)
        }
      }

      # Detect run commands that invoke Rust tooling.
      if (trimmed ~ /^run:[[:space:]]*/) {
        value = trimmed
        sub(/^run:[[:space:]]*/, "", value)
        value = trim(value)
        sub(/[[:space:]]+#.*/, "", value)
        value = strip_quotes(value)
        if (line_looks_like_rust_tooling(value)) {
          note_rust_use("run: " value)
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
        finalize_steps_block()
      }
      exit fail
    }
  ' "$workflow" || fail=1
done

# Optional local tooling pin: mise.toml
#
# `mise.toml` is not required in all environments (Rust is primarily managed via rustup), but when
# present it should stay aligned with the repo's pinned toolchain.
mise_file="mise.toml"
if [ -f "${mise_file}" ]; then
  mise_rust_line="$(grep -E '^[[:space:]]*rust[[:space:]]*=' "${mise_file}" | head -n 1 || true)"
  if [ -n "${mise_rust_line}" ]; then
    mise_rust="${mise_rust_line#*=}"
    mise_rust="${mise_rust%%#*}"
    mise_rust="${mise_rust#"${mise_rust%%[![:space:]]*}"}"
    mise_rust="${mise_rust%"${mise_rust##*[![:space:]]}"}"
    if [[ "${mise_rust}" == \"*\" ]]; then
      mise_rust="${mise_rust#\"}"
      mise_rust="${mise_rust%\"}"
    elif [[ "${mise_rust}" == \'*\' ]]; then
      mise_rust="${mise_rust#\'}"
      mise_rust="${mise_rust%\'}"
    fi
    mise_rust="${mise_rust#v}"
    mise_rust="${mise_rust#V}"

    if ! [[ "${mise_rust}" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
      echo "mise.toml rust tool version must be patch-pinned (X.Y.Z); found rust=${mise_rust}" >&2
      fail=1
    elif [ "${mise_rust}" != "${channel}" ]; then
      echo "Rust toolchain pin mismatch between rust-toolchain.toml and mise.toml:" >&2
      echo "  rust-toolchain.toml: channel=${channel}" >&2
      echo "  mise.toml:           rust=${mise_rust}" >&2
      echo "" >&2
      echo "Fix: update mise.toml rust pin to ${channel} (or update rust-toolchain.toml)." >&2
      fail=1
    fi
  fi
fi

if [ "$fail" -ne 0 ]; then
  exit 1
fi

echo "Rust toolchain pins match rust-toolchain.toml (channel=${channel})."
