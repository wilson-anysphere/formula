#!/usr/bin/env bash
set -euo pipefail

# High-throughput cargo wrapper for multi-agent hosts.
#
# Goals:
# - Avoid cargo/rustc stampedes when many agents run commands concurrently
# - Enforce per-command RAM ceiling via RLIMIT_AS (no cgroups needed)
# - Sensible defaults for test thread count
#
# Usage:
#   scripts/cargo_agent.sh build --release
#   scripts/cargo_agent.sh test --lib
#   scripts/cargo_agent.sh test -p formula-engine
#   scripts/cargo_agent.sh check
#
# Environment:
#   FORMULA_CARGO_JOBS       cargo build jobs (default: 4)
#   FORMULA_CARGO_LIMIT_AS   Address-space cap (default: 12G)
#   FORMULA_RUST_TEST_THREADS  Default RUST_TEST_THREADS for cargo test (default: min(nproc, 16))
#
# Based on fastrender's cargo_agent.sh (simpler than cgroups/systemd-run).

usage() {
  cat <<'EOF'
usage: scripts/cargo_agent.sh <cargo-subcommand> [args...]

Examples:
  scripts/cargo_agent.sh check --quiet
  scripts/cargo_agent.sh build --release
  scripts/cargo_agent.sh test --lib
  scripts/cargo_agent.sh test -p formula-engine

Environment:
  FORMULA_CARGO_JOBS         cargo -j value (default: 4)
  FORMULA_CARGO_LIMIT_AS     Address-space cap (default: 12G)
  FORMULA_RUST_TEST_THREADS  RUST_TEST_THREADS for cargo test (default: min(nproc, 16))
EOF
}

if [[ $# -lt 1 ]]; then
  usage
  exit 2
fi

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

# Use repo-local cargo home by default to avoid lock contention
DEFAULT_GLOBAL_CARGO_HOME="${HOME:-/root}/.cargo"
if [[ -z "${CARGO_HOME:-}" ]] || {
  [[ -z "${CI:-}" ]] &&
  [[ -z "${FORMULA_ALLOW_GLOBAL_CARGO_HOME:-}" ]] &&
  [[ "${CARGO_HOME:-}" == "${DEFAULT_GLOBAL_CARGO_HOME}" ]];
}; then
  export CARGO_HOME="${repo_root}/target/cargo-home"
fi
mkdir -p "${CARGO_HOME}"
mkdir -p "${CARGO_HOME}/bin"
case ":${PATH}:" in
  *":${CARGO_HOME}/bin:"*) ;;
  *) export PATH="${CARGO_HOME}/bin:${PATH}" ;;
esac

# Detect nproc
nproc_val=""
if command -v nproc >/dev/null 2>&1; then
  nproc_val="$(nproc 2>/dev/null || true)"
fi
if ! [[ "${nproc_val}" =~ ^[0-9]+$ ]] || [[ "${nproc_val}" -lt 1 ]]; then
  nproc_val="$(getconf _NPROCESSORS_ONLN 2>/dev/null || true)"
fi
if ! [[ "${nproc_val}" =~ ^[0-9]+$ ]] || [[ "${nproc_val}" -lt 1 ]]; then
  if command -v sysctl >/dev/null 2>&1; then
    nproc_val="$(sysctl -n hw.logicalcpu 2>/dev/null || sysctl -n hw.ncpu 2>/dev/null || true)"
  fi
fi
if ! [[ "${nproc_val}" =~ ^[0-9]+$ ]] || [[ "${nproc_val}" -lt 1 ]]; then
  nproc_val=4
fi

# Defaults
jobs="${FORMULA_CARGO_JOBS:-4}"
limit_as="${FORMULA_CARGO_LIMIT_AS:-12G}"

# Subcommand
subcommand="$1"
shift

# For test runs, cap RUST_TEST_THREADS to avoid spawning hundreds of threads
if [[ "${subcommand}" == "test" && -z "${RUST_TEST_THREADS:-}" ]]; then
  rust_test_threads="${FORMULA_RUST_TEST_THREADS:-}"
  if [[ -z "${rust_test_threads}" ]]; then
    rust_test_threads=$(( nproc_val < 16 ? nproc_val : 16 ))
  fi
  export RUST_TEST_THREADS="${rust_test_threads}"
fi

cargo_cmd=(cargo "${subcommand}" -j "${jobs}" "$@")

echo "cargo_agent: jobs=${jobs} as=${limit_as} test_threads=${RUST_TEST_THREADS:-auto}" >&2

# Run through run_limited.sh if it exists and limit_as is set
if [[ -n "${limit_as}" && "${limit_as}" != "0" && "${limit_as}" != "off" && "${limit_as}" != "unlimited" ]]; then
  if [[ -x "${repo_root}/scripts/run_limited.sh" ]]; then
    exec bash "${repo_root}/scripts/run_limited.sh" --as "${limit_as}" -- "${cargo_cmd[@]}"
  fi
fi

exec "${cargo_cmd[@]}"
