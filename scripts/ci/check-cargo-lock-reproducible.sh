#!/usr/bin/env bash
set -euo pipefail

# Ensure tagged desktop releases are built with exactly the dependency set committed to the repo.
#
# This script is intended to run in CI (preflight) and fail fast if Cargo would update Cargo.lock.

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

desktop_manifest="apps/desktop/src-tauri/Cargo.toml"

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
  if ! cargo metadata \
    --locked \
    --format-version=1 \
    --manifest-path "${desktop_manifest}" \
    --features desktop \
    --filter-platform "${target}" \
    >/dev/null; then
    echo "::error::Cargo.lock is out of date for desktop release target ${target}. Run 'cargo generate-lockfile' (or build locally) and commit the updated Cargo.lock."
    exit 1
  fi
done

# The desktop app also bundles a WASM engine (`crates/formula-wasm`), built via wasm-pack. Validate
# the workspace lockfile for wasm32 as well so we fail early if `Cargo.lock` is out of date for the
# engine dependency graph.
echo "Checking Cargo.lock for target: wasm32-unknown-unknown"
if ! cargo metadata \
  --locked \
  --format-version=1 \
  --filter-platform wasm32-unknown-unknown \
  >/dev/null; then
  echo "::error::Cargo.lock is out of date for wasm32-unknown-unknown. Run 'cargo generate-lockfile' (or build locally) and commit the updated Cargo.lock."
  exit 1
fi

echo "Cargo.lock is consistent with the committed dependency graph."
