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
#   scripts/cargo_agent.sh fmt -- --check
#
# Environment:
#   FORMULA_CARGO_JOBS       cargo build jobs (default: 4)
#   FORMULA_CARGO_LIMIT_AS   Address-space cap (default: 12G)
#   FORMULA_RUST_TEST_THREADS  Default RUST_TEST_THREADS for cargo test (default: min(nproc, 16))
#   FORMULA_RAYON_NUM_THREADS  Default RAYON_NUM_THREADS (default: FORMULA_CARGO_JOBS)
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
  scripts/cargo_agent.sh fmt -- --check

Environment:
  FORMULA_CARGO_JOBS         cargo -j value (default: 4)
  FORMULA_CARGO_LIMIT_AS     Address-space cap (default: 12G)
  FORMULA_RUST_TEST_THREADS  RUST_TEST_THREADS for cargo test (default: min(nproc, 16))
  FORMULA_RAYON_NUM_THREADS  RAYON_NUM_THREADS (default: FORMULA_CARGO_JOBS)
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
#
# Prefer the wrapper-specific override, but fall back to the standard Cargo env var so
# `source scripts/agent-init.sh` (which sets `CARGO_BUILD_JOBS`) influences the wrapper too.
jobs="${FORMULA_CARGO_JOBS:-${CARGO_BUILD_JOBS:-4}}"
limit_as="${FORMULA_CARGO_LIMIT_AS:-14G}"

# Rustc can spawn many internal worker threads during codegen (roughly proportional to
# `codegen-units`). On multi-agent machines this can hit OS thread limits and manifest as
# rustc ICEs like "failed to spawn work thread: Resource temporarily unavailable".
#
# To keep per-process thread counts proportional to our chosen Cargo job parallelism, default
# dev/test codegen units to `jobs` unless the caller has explicitly overridden them.
if [[ -z "${CARGO_PROFILE_DEV_CODEGEN_UNITS:-}" ]]; then
  export CARGO_PROFILE_DEV_CODEGEN_UNITS="${jobs}"
fi
if [[ -z "${CARGO_PROFILE_TEST_CODEGEN_UNITS:-}" ]]; then
  export CARGO_PROFILE_TEST_CODEGEN_UNITS="${jobs}"
fi

# Cargo also supports configuring the default parallelism via `CARGO_BUILD_JOBS`.
# Export it so commands that *don't* accept `-j` (e.g. `cargo fmt`, `cargo clean`)
# still inherit our safe default.
export CARGO_BUILD_JOBS="${jobs}"

# Rayon: default to a small thread pool on high-core agent hosts.
#
# Rayon defaults to spawning one worker per core. On our multi-agent machines this can
# lead to very large per-test-process thread pools which are both wasteful and can
# fail to initialize (e.g. "Resource temporarily unavailable") under load.
#
# Allow callers to override via either RAYON_NUM_THREADS (Rayon's standard env var),
# or the wrapper-specific FORMULA_RAYON_NUM_THREADS.
if [[ -z "${RAYON_NUM_THREADS:-}" ]]; then
  rayon_threads="${FORMULA_RAYON_NUM_THREADS:-}"
  if [[ -z "${rayon_threads}" ]]; then
    rayon_threads="${jobs}"
  fi
  export RAYON_NUM_THREADS="${rayon_threads}"
fi

# OpenSSL: avoid expensive vendored builds when a system OpenSSL is available.
#
# Some crates enable the `openssl/vendored` feature for portability, but building OpenSSL from
# source is slow and can fail under process/thread pressure on multi-agent hosts.
#
# When we're on Linux and `pkg-config` can locate OpenSSL, force `openssl-sys` to use the system
# installation even if `vendored` is enabled.
#
# Override:
# - set `OPENSSL_NO_VENDOR` explicitly to control openssl-sys directly, or
# - set `FORMULA_OPENSSL_VENDOR=1` to prevent this wrapper from setting `OPENSSL_NO_VENDOR`.
if [[ -z "${CI:-}" && -z "${OPENSSL_NO_VENDOR:-}" && -z "${FORMULA_OPENSSL_VENDOR:-}" ]]; then
  if command -v uname >/dev/null 2>&1 && [[ "$(uname -s)" == "Linux" ]]; then
    if command -v pkg-config >/dev/null 2>&1 && pkg-config --exists openssl; then
      export OPENSSL_NO_VENDOR=1
    fi
  fi
fi

# Subcommand
subcommand="$1"
shift

# Backwards-compat: the desktop Tauri crate has historically used both
# `formula-desktop-tauri` and `desktop` as the Cargo *package* name. Some scripts still hardcode
# one or the other. Detect the current package name from `apps/desktop/src-tauri/Cargo.toml` and
# rewrite `-p/--package` accordingly.
desktop_pkg_name=""
if [[ -f "${repo_root}/apps/desktop/src-tauri/Cargo.toml" ]]; then
  # Extract `name = "..."` from the `[package]` section.
  desktop_pkg_name="$(
    awk '
      BEGIN { in_pkg=0 }
      /^[[:space:]]*\[package\][[:space:]]*$/ { in_pkg=1; next }
      in_pkg && /^[[:space:]]*\[/ { in_pkg=0 }
      in_pkg {
        if ($0 ~ /^[[:space:]]*name[[:space:]]*=[[:space:]]*"/) {
          line = $0
          sub(/^[[:space:]]*name[[:space:]]*=[[:space:]]*"/, "", line)
          sub(/".*$/, "", line)
          print line
          exit
        }
      }
    ' "${repo_root}/apps/desktop/src-tauri/Cargo.toml" 2>/dev/null || true
  )"
fi

remap_from=""
remap_to=""
if [[ "${desktop_pkg_name}" == "desktop" ]]; then
  remap_from="formula-desktop-tauri"
  remap_to="desktop"
elif [[ "${desktop_pkg_name}" == "formula-desktop-tauri" ]]; then
  remap_from="desktop"
  remap_to="formula-desktop-tauri"
fi
remapped_args=()
expect_pkg_name=false
for arg in "$@"; do
  if [[ "${expect_pkg_name}" == "true" ]]; then
    if [[ -n "${remap_from}" && "${arg}" == "${remap_from}" ]]; then
      remapped_args+=("${remap_to}")
    else
      remapped_args+=("${arg}")
    fi
    expect_pkg_name=false
    continue
  fi

  case "${arg}" in
    -p|--package)
      remapped_args+=("${arg}")
      expect_pkg_name=true
      ;;
    -p=${remap_from})
      remapped_args+=("-p=${remap_to}")
      ;;
    --package=${remap_from})
      remapped_args+=("--package=${remap_to}")
      ;;
    *)
      remapped_args+=("${arg}")
      ;;
  esac
done
set -- "${remapped_args[@]}"

# For test runs, cap RUST_TEST_THREADS to avoid spawning hundreds of threads
if [[ "${subcommand}" == "test" && -z "${RUST_TEST_THREADS:-}" ]]; then
  rust_test_threads="${FORMULA_RUST_TEST_THREADS:-}"
  if [[ -z "${rust_test_threads}" ]]; then
    rust_test_threads=$(( nproc_val < 16 ? nproc_val : 16 ))
  fi
  export RUST_TEST_THREADS="${rust_test_threads}"
fi

# Limit rustc's internal codegen parallelism on multi-agent hosts.
#
# Even with `cargo -j` and Rayon's thread pool capped, rustc/LLVM may still spawn helper threads
# and/or one worker thread per codegen unit. Under high system load this can fail with
# "Resource temporarily unavailable" (EAGAIN) and crash compilation.
#
# Default to a single codegen unit for tests to reduce thread usage. Callers can override via the
# standard Cargo env vars (`CARGO_PROFILE_{dev,test}_CODEGEN_UNITS`).
if [[ "${subcommand}" == "test" ]]; then
  if [[ -z "${CARGO_PROFILE_DEV_CODEGEN_UNITS:-}" ]]; then
    export CARGO_PROFILE_DEV_CODEGEN_UNITS="1"
  fi
  if [[ -z "${CARGO_PROFILE_TEST_CODEGEN_UNITS:-}" ]]; then
    export CARGO_PROFILE_TEST_CODEGEN_UNITS="1"
  fi
fi

# Only some Cargo subcommands accept `-j/--jobs`. `cargo fmt`, `cargo clean`, etc
# reject it, and we want `cargo_agent.sh` to be usable for *any* cargo invocation.
case "${subcommand}" in
  bench|build|check|clippy|doc|install|run|rustc|test)
    cargo_cmd=(cargo "${subcommand}" -j "${jobs}" "$@")
    ;;
  *)
    cargo_cmd=(cargo "${subcommand}" "$@")
    ;;
esac

echo "cargo_agent: jobs=${jobs} as=${limit_as} test_threads=${RUST_TEST_THREADS:-auto}" >&2

# Run through run_limited.sh if it exists and limit_as is set
if [[ -n "${limit_as}" && "${limit_as}" != "0" && "${limit_as}" != "off" && "${limit_as}" != "unlimited" ]]; then
  if [[ -x "${repo_root}/scripts/run_limited.sh" ]]; then
    exec bash "${repo_root}/scripts/run_limited.sh" --as "${limit_as}" -- "${cargo_cmd[@]}"
  fi
fi

exec "${cargo_cmd[@]}"
