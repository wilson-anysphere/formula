#!/bin/bash
# Initialize agent environment with safe memory defaults
# Source this at the start of each agent session: . scripts/agent-init.sh
# (bash/zsh also support: source scripts/agent-init.sh)

# Note: if you run this script directly (`bash scripts/agent-init.sh`), it cannot modify the
# environment of your current shell. We still run, but emit a warning so it's harder to miss.
if [ -n "${BASH_VERSION:-}" ]; then
  if [ "${BASH_SOURCE[0]}" = "${0}" ]; then
    echo "warning: scripts/agent-init.sh is meant to be sourced ('. scripts/agent-init.sh'); executing it won't update your current shell" >&2
  fi
elif [ "${0##*/}" = "agent-init.sh" ]; then
  echo "warning: scripts/agent-init.sh is meant to be sourced ('. scripts/agent-init.sh'); executing it won't update your current shell" >&2
fi

_formula_errexit_was_set=0
case $- in
  *e*) _formula_errexit_was_set=1 ;;
esac
set -e

# If wrapper-specific config variables are set without being exported (common in interactive shells),
# export them so subprocesses (including `scripts/cargo_agent.sh`) can observe the intended
# overrides.
if [ -n "${FORMULA_ALLOW_GLOBAL_CARGO_HOME:-}" ]; then
  export FORMULA_ALLOW_GLOBAL_CARGO_HOME
fi
if [ -n "${FORMULA_CARGO_JOBS:-}" ]; then
  export FORMULA_CARGO_JOBS
fi
if [ -n "${FORMULA_CARGO_TEST_JOBS:-}" ]; then
  export FORMULA_CARGO_TEST_JOBS
fi
if [ -n "${FORMULA_CARGO_LIMIT_AS:-}" ]; then
  export FORMULA_CARGO_LIMIT_AS
fi
if [ -n "${FORMULA_RUST_TEST_THREADS:-}" ]; then
  export FORMULA_RUST_TEST_THREADS
fi
if [ -n "${FORMULA_RAYON_NUM_THREADS:-}" ]; then
  export FORMULA_RAYON_NUM_THREADS
fi
if [ -n "${FORMULA_OPENSSL_VENDOR:-}" ]; then
  export FORMULA_OPENSSL_VENDOR
fi
if [ -n "${FORMULA_CARGO_RETRY_ATTEMPTS:-}" ]; then
  export FORMULA_CARGO_RETRY_ATTEMPTS
fi
if [ -n "${FORMULA_LLD_THREADS:-}" ]; then
  export FORMULA_LLD_THREADS
fi

# `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
# globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift.
#
# Clear it when we're inside this repo so subsequent `cargo` invocations (including ones that don't
# use wrapper scripts) respect `rust-toolchain.toml`.
if [ -n "${RUSTUP_TOOLCHAIN:-}" ] && command -v git >/dev/null 2>&1; then
  _formula_repo_root="$(git rev-parse --show-toplevel 2>/dev/null || true)"
  if [ -n "${_formula_repo_root}" ] && [ -f "${_formula_repo_root}/rust-toolchain.toml" ]; then
    unset RUSTUP_TOOLCHAIN
  fi
  unset _formula_repo_root
fi

# ============================================================================
# Memory Limits (CRITICAL)
# ============================================================================

# Node.js: 3GB heap limit (leaves room for other processes).
#
# Preserve any existing NODE_OPTIONS flags (some environments set global flags) and only inject the
# heap cap when callers haven't set one already.
if [ -z "${NODE_OPTIONS:-}" ]; then
  NODE_OPTIONS="--max-old-space-size=3072"
else
  case " ${NODE_OPTIONS} " in
    *" --max-old-space-size="*) ;;
    *) NODE_OPTIONS="--max-old-space-size=3072 ${NODE_OPTIONS}" ;;
  esac
fi
export NODE_OPTIONS

# Rust: Limit parallel compilation jobs (each can use 1-2GB).
#
# Prefer the wrapper-specific override (`FORMULA_CARGO_JOBS`) when provided. Otherwise respect any
# explicit `CARGO_BUILD_JOBS` override (useful for CI/global setups).
#
# When neither is set, default to conservative parallelism. On very high core-count machines, lld
# (and other toolchain components) can spawn many threads per invocation; combining that with
# Cargo-level parallelism can exceed sandbox process/thread limits and cause flaky
# "Resource temporarily unavailable" failures.
_formula_cpu_count=""
if command -v nproc >/dev/null 2>&1; then
  _formula_cpu_count="$(nproc 2>/dev/null || true)"
fi
if [ -z "${_formula_cpu_count}" ]; then
  _formula_cpu_count="$(getconf _NPROCESSORS_ONLN 2>/dev/null || true)"
fi
if [ -z "${_formula_cpu_count}" ] && command -v sysctl >/dev/null 2>&1; then
  _formula_cpu_count="$(sysctl -n hw.logicalcpu 2>/dev/null || sysctl -n hw.ncpu 2>/dev/null || true)"
fi
case "${_formula_cpu_count}" in
  ''|*[!0-9]*) _formula_cpu_count=0 ;;
esac
_formula_default_cargo_jobs=4
if [ "${_formula_cpu_count}" -ge 64 ]; then
  _formula_default_cargo_jobs=2
fi
unset _formula_cpu_count

_formula_prev_cargo_build_jobs="${CARGO_BUILD_JOBS:-}"
if [ -n "${FORMULA_CARGO_JOBS:-}" ]; then
  export CARGO_BUILD_JOBS="${FORMULA_CARGO_JOBS}"
elif [ -z "${CARGO_BUILD_JOBS:-}" ]; then
  export CARGO_BUILD_JOBS="${_formula_default_cargo_jobs}"
fi
export CARGO_BUILD_JOBS
unset _formula_default_cargo_jobs

# Make: Limit parallel jobs
#
# Keep MAKEFLAGS aligned with the chosen Cargo job count when MAKEFLAGS was previously derived from
# `CARGO_BUILD_JOBS` (common when re-sourcing this script after changing `FORMULA_CARGO_JOBS`).
if [ -z "${MAKEFLAGS:-}" ]; then
  export MAKEFLAGS="-j${CARGO_BUILD_JOBS}"
elif [ -n "${_formula_prev_cargo_build_jobs}" ] && [ "${MAKEFLAGS}" = "-j${_formula_prev_cargo_build_jobs}" ]; then
  export MAKEFLAGS="-j${CARGO_BUILD_JOBS}"
fi
export MAKEFLAGS

# Rayon: Limit thread pool size (defaults to one thread per core otherwise).
#
# Keep RAYON_NUM_THREADS aligned with the chosen Cargo job count when it was previously derived
# from the prior `CARGO_BUILD_JOBS` value (common when re-sourcing this script after changing
# `FORMULA_CARGO_JOBS`).
if [ -z "${RAYON_NUM_THREADS:-}" ] || {
  [ -n "${_formula_prev_cargo_build_jobs}" ] && [ "${RAYON_NUM_THREADS:-}" = "${_formula_prev_cargo_build_jobs}" ];
}; then
  export RAYON_NUM_THREADS="${FORMULA_RAYON_NUM_THREADS:-${CARGO_BUILD_JOBS}}"
fi
export RAYON_NUM_THREADS
unset _formula_prev_cargo_build_jobs

# Rust codegen units:
#
# Avoid setting `RUSTFLAGS=-C codegen-units=N` here. While it can reduce memory usage,
# it also overrides Cargo profile configuration and can defeat the safeguards in
# `scripts/cargo_agent.sh` that scale codegen parallelism to the chosen Cargo job count.
# Under high load, that mismatch can surface as flaky rustc ICEs like:
# "failed to spawn work thread: Resource temporarily unavailable".
#
# `scripts/cargo_agent.sh` already sets `CARGO_PROFILE_{DEV,TEST}_CODEGEN_UNITS` to a
# safe default automatically, and agents should always use that wrapper instead of
# invoking `cargo` directly.

# ============================================================================
# Cargo Home Isolation (CRITICAL - prevents cross-agent ~/.cargo lock contention)
# ============================================================================

# Cargo defaults to using ~/.cargo for registry/index/git caches. With many agents
# building in parallel this creates heavy lock contention under ~/.cargo and can
# make builds flaky/slow. Default to a repo-local CARGO_HOME to isolate agents.
# Note: some agent runners pre-set `CARGO_HOME=$HOME/.cargo`. Treat that value as
# "unset" for our purposes so we still get per-repo isolation by default.
# In CI we respect `CARGO_HOME` even if it points at `$HOME/.cargo` so CI can use
# shared caching.
# To explicitly keep `CARGO_HOME=$HOME/.cargo` in local runs, set
# `FORMULA_ALLOW_GLOBAL_CARGO_HOME=1` before sourcing this script.
_formula_default_global_cargo_home="${HOME:-/root}/.cargo"
_formula_cargo_home_norm="${CARGO_HOME:-}"
_formula_cargo_home_norm="${_formula_cargo_home_norm%/}"
_formula_default_global_cargo_home_norm="${_formula_default_global_cargo_home%/}"
if [ -z "${CARGO_HOME:-}" ] || {
  [ -z "${CI:-}" ] &&
    [ -z "${FORMULA_ALLOW_GLOBAL_CARGO_HOME:-}" ] &&
    [ "${_formula_cargo_home_norm}" = "${_formula_default_global_cargo_home_norm}" ];
}; then
  # Prefer `git rev-parse --show-toplevel` to locate the repo root. This works
  # even when sourced from `sh` (our agent runner shell), where bash-only
  # variables like `BASH_SOURCE` are unavailable.
  #
  # Fall back to `BASH_SOURCE` when running in bash (e.g. local dev), and finally
  # to `pwd` if neither is available.
  _formula_repo_root=""
  if command -v git >/dev/null 2>&1; then
    _formula_repo_root="$(git rev-parse --show-toplevel 2>/dev/null || true)"
  fi
  if [ -z "${_formula_repo_root}" ] && [ -n "${BASH_VERSION:-}" ]; then
    _formula_repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
  fi
  if [ -z "${_formula_repo_root}" ]; then
    _formula_repo_root="$(pwd)"
  fi
  export CARGO_HOME="${_formula_repo_root}/target/cargo-home"
fi
export CARGO_HOME
unset _formula_default_global_cargo_home
unset _formula_cargo_home_norm
unset _formula_default_global_cargo_home_norm
unset _formula_repo_root
mkdir -p "$CARGO_HOME"

# Ensure tools installed via `cargo install` under this CARGO_HOME are available.
mkdir -p "$CARGO_HOME/bin"
# Ensure `$CARGO_HOME/bin` is the *first* PATH entry, even if it already exists
# later in the PATH (e.g. from a login shell's profile).
_formula_path_without_cargo_bin=""
_formula_ifs_was_set=0
if [ "${IFS+x}" = "x" ]; then
  _formula_ifs_was_set=1
  _formula_old_ifs="${IFS}"
else
  _formula_old_ifs=""
fi
IFS=:
for _formula_entry in ${PATH:-}; do
  if [ "${_formula_entry}" = "${CARGO_HOME}/bin" ]; then
    continue
  fi
  if [ -z "${_formula_path_without_cargo_bin}" ]; then
    _formula_path_without_cargo_bin="${_formula_entry}"
  else
    _formula_path_without_cargo_bin="${_formula_path_without_cargo_bin}:${_formula_entry}"
  fi
done
if [ "${_formula_ifs_was_set}" -eq 1 ]; then
  IFS="${_formula_old_ifs}"
else
  unset IFS
fi
unset _formula_old_ifs
unset _formula_ifs_was_set
unset _formula_entry

if [ -n "${_formula_path_without_cargo_bin}" ]; then
  export PATH="${CARGO_HOME}/bin:${_formula_path_without_cargo_bin}"
else
  export PATH="${CARGO_HOME}/bin"
fi
unset _formula_path_without_cargo_bin

# ============================================================================
# Shared Caching (Optional - reduces disk duplication)
# ============================================================================

# Uncomment if sccache is installed and /shared/sccache exists
# export RUSTC_WRAPPER=sccache
# export SCCACHE_DIR=/shared/sccache
# export SCCACHE_CACHE_SIZE="50G"

# ============================================================================
# Headless Display Setup
# ============================================================================

setup_display() {
  if [ -z "${DISPLAY:-}" ]; then
    # Try to find an existing Xvfb
    if [ -e /tmp/.X99-lock ]; then
      export DISPLAY=:99
    elif [ -e /tmp/.X98-lock ]; then
      export DISPLAY=:98
    else
      # Start a new Xvfb on an available display
      for d in 99 98 97 96 95; do
        if [ ! -e "/tmp/.X${d}-lock" ]; then
          export DISPLAY=:$d
          Xvfb :$d -screen 0 1920x1080x24 >/dev/null 2>&1 &
          sleep 0.5
          break
        fi
      done
    fi
  fi
}

# Only setup display if Xvfb is available
if command -v Xvfb >/dev/null 2>&1; then
  setup_display
fi
if [ -n "${DISPLAY:-}" ]; then
  export DISPLAY
fi
unset -f setup_display

# ============================================================================
# Confirmation
# ============================================================================

echo "╔════════════════════════════════════════════════════════════════╗"
echo "║  Agent Environment Initialized                                  ║"
echo "╠════════════════════════════════════════════════════════════════╣"
echo "║  NODE_OPTIONS:      ${NODE_OPTIONS}"
echo "║  CARGO_BUILD_JOBS:  ${CARGO_BUILD_JOBS}"
echo "║  RAYON_NUM_THREADS: ${RAYON_NUM_THREADS}"
echo "║  CARGO_HOME:        ${CARGO_HOME}"
echo "║  MAKEFLAGS:         ${MAKEFLAGS}"
if [ -n "${DISPLAY:-}" ]; then
echo "║  DISPLAY:           ${DISPLAY}"
fi
echo "╚════════════════════════════════════════════════════════════════╝"

if [ "${_formula_errexit_was_set}" -eq 0 ]; then
  set +e
fi
unset _formula_errexit_was_set
