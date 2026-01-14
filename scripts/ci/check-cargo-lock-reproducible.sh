#!/usr/bin/env bash
set -euo pipefail

# Ensure tagged desktop releases are built with exactly the dependency set committed to the repo.
#
# This script is intended to run in CI (preflight) and fail fast if Cargo would update Cargo.lock.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

# `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
# globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift
# for CI/local preflight runs.
if [[ -n "${RUSTUP_TOOLCHAIN:-}" && -f "${repo_root}/rust-toolchain.toml" ]]; then
  unset RUSTUP_TOOLCHAIN
fi

desktop_manifest="apps/desktop/src-tauri/Cargo.toml"

mktemp_file() {
  # `mktemp` behavior differs between GNU and BSD (macOS). GNU mktemp supports
  # being called without a template; BSD mktemp requires one.
  local tmp
  if tmp="$(mktemp 2>/dev/null)"; then
    echo "$tmp"
    return 0
  fi
  mktemp -t formula-cargo-lock.XXXXXX
}

run_metadata() {
  local context="$1"
  shift

  local err
  err="$(mktemp_file)"
  if cargo metadata "$@" >/dev/null 2>"$err"; then
    rm -f "$err"
    return 0
  fi

  # Always print Cargo's actual error output for debugging. If the failure is due
  # to a stale lockfile, add a short remediation hint on top.
  cat "$err" >&2 || true

  if grep -Eqi "needs to be updated but --locked was passed|but --locked was passed to prevent this" "$err"; then
    echo "::error::Cargo.lock is out of date (${context}). Run 'cargo generate-lockfile' (or build locally) and commit the updated Cargo.lock."
  else
    echo "::error::cargo metadata failed (${context}). See logs above."
  fi

  rm -f "$err"
  return 1
}

desktop_targets=(
  x86_64-unknown-linux-gnu
  aarch64-unknown-linux-gnu
  x86_64-pc-windows-msvc
  aarch64-pc-windows-msvc
  x86_64-apple-darwin
  aarch64-apple-darwin
)

for target in "${desktop_targets[@]}"; do
  echo "Checking Cargo.lock for target: ${target}"
  run_metadata "desktop release target ${target}" \
    --locked \
    --format-version=1 \
    --manifest-path "${desktop_manifest}" \
    --features desktop \
    --filter-platform "${target}"
done

# The desktop app also bundles a WASM engine (`crates/formula-wasm`), built via wasm-pack. Validate
# the workspace lockfile for wasm32 as well so we fail early if `Cargo.lock` is out of date for the
# engine dependency graph.
echo "Checking Cargo.lock for target: wasm32-unknown-unknown"
run_metadata "wasm32-unknown-unknown" \
  --locked \
  --format-version=1 \
  --filter-platform wasm32-unknown-unknown

echo "Cargo.lock is consistent with the committed dependency graph."
