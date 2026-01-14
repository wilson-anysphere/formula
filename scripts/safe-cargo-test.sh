#!/bin/bash
# Memory-aware wrapper for cargo test
# Usage: ./scripts/safe-cargo-test.sh [cargo test arguments...]
#
# Examples:
#   ./scripts/safe-cargo-test.sh
#   ./scripts/safe-cargo-test.sh --release
#   ./scripts/safe-cargo-test.sh -p formula-engine -- --test-threads=4

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml`. Some environments set it
# globally (often to `stable`), which would bypass the pinned toolchain and reintroduce
# "whatever stable is today" drift.
#
# Clear it so these wrappers reliably use the pinned toolchain.
if [ -n "${RUSTUP_TOOLCHAIN:-}" ] && [ -f "${REPO_ROOT}/rust-toolchain.toml" ]; then
  unset RUSTUP_TOOLCHAIN
fi

# Use a repo-local cargo home by default to avoid lock contention on ~/.cargo
# when many agents build in parallel. Preserve any user/CI override.
#
# Note: some agent runners pre-set `CARGO_HOME=$HOME/.cargo`. Treat that value as
# "unset" for our purposes so we still get per-repo isolation by default.
# In CI we respect `CARGO_HOME` even if it points at `$HOME/.cargo` so CI can use
# shared caching.
# To explicitly keep `CARGO_HOME=$HOME/.cargo` in local runs, set
# `FORMULA_ALLOW_GLOBAL_CARGO_HOME=1`.
DEFAULT_GLOBAL_CARGO_HOME="${HOME:-/root}/.cargo"
CARGO_HOME_NORM="${CARGO_HOME:-}"
CARGO_HOME_NORM="${CARGO_HOME_NORM%/}"
DEFAULT_GLOBAL_CARGO_HOME_NORM="${DEFAULT_GLOBAL_CARGO_HOME%/}"
if [ -z "${CARGO_HOME:-}" ] || {
  [ -z "${CI:-}" ] &&
    [ -z "${FORMULA_ALLOW_GLOBAL_CARGO_HOME:-}" ] &&
    [ "${CARGO_HOME_NORM}" = "${DEFAULT_GLOBAL_CARGO_HOME_NORM}" ];
}; then
  export CARGO_HOME="$REPO_ROOT/target/cargo-home"
fi
mkdir -p "$CARGO_HOME"
mkdir -p "$CARGO_HOME/bin"
case ":$PATH:" in
  *":$CARGO_HOME/bin:"*) ;;
  *) export PATH="$CARGO_HOME/bin:$PATH" ;;
esac

# Get smart job count
# Job count:
# - Respect explicit caller overrides first (`FORMULA_CARGO_JOBS` / `CARGO_BUILD_JOBS`).
# - Fall back to the adaptive helper when no explicit job count is configured.
if [ -n "${FORMULA_CARGO_JOBS:-}" ]; then
  JOBS="${FORMULA_CARGO_JOBS}"
elif [ -n "${CARGO_BUILD_JOBS:-}" ]; then
  JOBS="${CARGO_BUILD_JOBS}"
elif [ -x "$SCRIPT_DIR/smart-jobs.sh" ]; then
  JOBS=$("$SCRIPT_DIR/smart-jobs.sh")
else
  JOBS=4
fi

# Detect nproc (for RUST_TEST_THREADS default).
nproc_val=""
if command -v nproc >/dev/null 2>&1; then
  nproc_val="$(nproc 2>/dev/null || true)"
fi
if ! [[ "${nproc_val}" =~ ^[0-9]+$ ]] || [[ "${nproc_val}" -lt 1 ]]; then
  nproc_val="$(getconf _NPROCESSORS_ONLN 2>/dev/null || true)"
fi
if ! [[ "${nproc_val}" =~ ^[0-9]+$ ]] || [[ "${nproc_val}" -lt 1 ]]; then
  nproc_val=4
fi

export CARGO_BUILD_JOBS="${JOBS}"
export MAKEFLAGS="${MAKEFLAGS:--j${JOBS}}"

# Reduce rustc's internal codegen parallelism during tests to avoid spawning large numbers of
# helper threads on high-core-count hosts.
export CARGO_PROFILE_DEV_CODEGEN_UNITS="${CARGO_PROFILE_DEV_CODEGEN_UNITS:-1}"
export CARGO_PROFILE_TEST_CODEGEN_UNITS="${CARGO_PROFILE_TEST_CODEGEN_UNITS:-1}"
export CARGO_PROFILE_RELEASE_CODEGEN_UNITS="${CARGO_PROFILE_RELEASE_CODEGEN_UNITS:-1}"

# Rayon defaults to spawning one worker per core; cap it unless callers explicitly override it.
export RAYON_NUM_THREADS="${RAYON_NUM_THREADS:-${FORMULA_RAYON_NUM_THREADS:-${JOBS}}}"

# Rust's test harness defaults to running one thread per core. Cap it for stability unless callers
# explicitly override it.
if [ -z "${RUST_TEST_THREADS:-}" ]; then
  if [ "${nproc_val}" -lt 16 ]; then
    export RUST_TEST_THREADS="${nproc_val}"
  else
    export RUST_TEST_THREADS=16
  fi
fi

echo "🧪 Testing with -j${JOBS} (based on available memory)..."

limit_as="${FORMULA_CARGO_LIMIT_AS:-14G}"
cargo_cmd=(cargo test -j"$JOBS" "$@")
if [ -x "$SCRIPT_DIR/run_limited.sh" ] && [ -n "${limit_as}" ] && [ "${limit_as}" != "0" ] && [ "${limit_as}" != "off" ] && [ "${limit_as}" != "unlimited" ]; then
  bash "$SCRIPT_DIR/run_limited.sh" --as "${limit_as}" -- "${cargo_cmd[@]}"
else
  "${cargo_cmd[@]}"
fi
