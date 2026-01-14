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
#   FORMULA_CARGO_TEST_JOBS  cargo build jobs for `cargo test` (default: 1 unless FORMULA_CARGO_JOBS is set)
#   FORMULA_CARGO_LIMIT_AS   Address-space cap (default: 14G)
#   FORMULA_RUST_TEST_THREADS  Default RUST_TEST_THREADS for cargo test
#                             (default: min(nproc, 16, jobs * 4))
#   FORMULA_RAYON_NUM_THREADS  Default RAYON_NUM_THREADS (default: FORMULA_CARGO_JOBS)
#   FORMULA_OPENSSL_VENDOR   Set to 1 to disable auto-setting OPENSSL_NO_VENDOR (allow vendored OpenSSL)
#   FORMULA_CARGO_RETRY_ATTEMPTS  Retry count for transient rustc EAGAIN panics (default: 5)
#   FORMULA_LLD_THREADS      When linking with lld via a cc driver, pass `--threads=<n>` (default: 1 on Linux)
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
  FORMULA_CARGO_TEST_JOBS    cargo -j value for `cargo test` (default: 1 unless FORMULA_CARGO_JOBS is set)
  FORMULA_CARGO_LIMIT_AS     Address-space cap (default: 14G)
  FORMULA_RUST_TEST_THREADS  RUST_TEST_THREADS for cargo test (default: min(nproc, 16, jobs * 4))
  FORMULA_RAYON_NUM_THREADS  RAYON_NUM_THREADS (default: FORMULA_CARGO_JOBS)
  FORMULA_OPENSSL_VENDOR     Set to 1 to disable auto-setting OPENSSL_NO_VENDOR (allow vendored OpenSSL)
  FORMULA_CARGO_RETRY_ATTEMPTS  Retry count for transient rustc EAGAIN panics (default: 5)
  FORMULA_LLD_THREADS        lld thread pool size for link steps (default: 1 on Linux)
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

# `RUSTUP_TOOLCHAIN` overrides `rust-toolchain.toml` (it has higher precedence than the
# repo-local toolchain file). Some environments set it globally (often to "stable"),
# which would reintroduce "whatever stable is today" drift for this repo.
#
# Clear it so cargo/rustc will respect the pinned toolchain declared in
# `rust-toolchain.toml`.
if [[ -n "${RUSTUP_TOOLCHAIN:-}" && -f "${repo_root}/rust-toolchain.toml" ]]; then
  unset RUSTUP_TOOLCHAIN
fi

# Use repo-local cargo home by default to avoid lock contention
DEFAULT_GLOBAL_CARGO_HOME="${HOME:-/root}/.cargo"

# Some agent runners pre-set `CARGO_HOME` to `$HOME/.cargo` (sometimes with a trailing slash).
# Treat that as "unset" for our purposes so we still get per-repo isolation by default.
CARGO_HOME_NORM="${CARGO_HOME:-}"
CARGO_HOME_NORM="${CARGO_HOME_NORM%/}"
DEFAULT_GLOBAL_CARGO_HOME_NORM="${DEFAULT_GLOBAL_CARGO_HOME%/}"
if [[ -z "${CARGO_HOME_NORM}" ]] || {
  [[ -z "${CI:-}" ]] &&
  [[ -z "${FORMULA_ALLOW_GLOBAL_CARGO_HOME:-}" ]] &&
  [[ "${CARGO_HOME_NORM}" == "${DEFAULT_GLOBAL_CARGO_HOME_NORM}" ]];
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
# `. scripts/agent-init.sh` (which sets `CARGO_BUILD_JOBS`) influences the wrapper too.
caller_jobs_env="${FORMULA_CARGO_JOBS:-${FORMULA_CARGO_TEST_JOBS:-}}"
jobs="${FORMULA_CARGO_JOBS:-${CARGO_BUILD_JOBS:-4}}"
# Fail fast when callers explicitly configure an invalid job count. (We keep a more forgiving
# fallback for generic env vars like `CARGO_BUILD_JOBS`, since those are sometimes set globally.)
if [[ -n "${FORMULA_CARGO_JOBS:-}" ]]; then
  if ! [[ "${FORMULA_CARGO_JOBS}" =~ ^[0-9]+$ ]] || [[ "${FORMULA_CARGO_JOBS}" -lt 1 ]]; then
    echo "cargo_agent: invalid FORMULA_CARGO_JOBS=${FORMULA_CARGO_JOBS} (expected integer >= 1)" >&2
    exit 2
  fi
fi
# Guardrails: reject/avoid dangerous values from generic env vars.
#
# - Cargo rejects 0 jobs (e.g. `-j0` or `CARGO_BUILD_JOBS=0`) with "jobs may not be 0".
#   Treat that as invalid/unset and fall back to a safe default.
# - Extremely large `CARGO_BUILD_JOBS` values (often set globally) can cause rustc stampedes and
#   `EAGAIN` thread/process spawn failures even under RLIMIT_AS.
#
# Allow opting into higher parallelism explicitly via `FORMULA_CARGO_JOBS`; otherwise clamp to a
# conservative upper bound.
if ! [[ "${jobs}" =~ ^[0-9]+$ ]] || [[ "${jobs}" -lt 1 ]]; then
  echo "cargo_agent: invalid CARGO_BUILD_JOBS/FORMULA_CARGO_JOBS value (${jobs}); defaulting to 4" >&2
  jobs="4"
fi
if [[ -z "${FORMULA_CARGO_JOBS:-}" && "${jobs}" -gt 8 ]]; then
  echo "cargo_agent: CARGO_BUILD_JOBS=${jobs} is high; clamping to 8 (set FORMULA_CARGO_JOBS to override)" >&2
  jobs="8"
fi
limit_as="${FORMULA_CARGO_LIMIT_AS:-14G}"
if [[ -n "${FORMULA_CARGO_LIMIT_AS:-}" ]]; then
  if [[ "${limit_as}" != "0" && "${limit_as}" != "off" && "${limit_as}" != "unlimited" ]]; then
    limit_as_norm="${limit_as//[[:space:]]/}"
    limit_as_norm="$(printf '%s' "${limit_as_norm}" | tr '[:upper:]' '[:lower:]')"
    limit_as_norm="${limit_as_norm%ib}"
    limit_as_norm="${limit_as_norm%b}"
    if ! [[ "${limit_as_norm}" =~ ^[0-9]+$ ]] && ! [[ "${limit_as_norm}" =~ ^[0-9]+[kmgt]$ ]]; then
      echo "cargo_agent: invalid FORMULA_CARGO_LIMIT_AS=${limit_as} (expected e.g. 14G, 4096M, 8GiB, 0, off, or unlimited)" >&2
      exit 2
    fi
  fi
fi

# Note: For `cargo test`, this wrapper may override the above `jobs` default to `1` (unless callers
# explicitly configure `FORMULA_CARGO_JOBS` / `FORMULA_CARGO_TEST_JOBS`). This avoids sporadic rustc
# thread spawn failures under high system load on multi-agent hosts.

# Record whether the caller explicitly configured Rayon thread counts before we set any defaults.
orig_rayon_num_threads="${RAYON_NUM_THREADS:-}"
orig_formula_rayon_num_threads="${FORMULA_RAYON_NUM_THREADS:-}"

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
  if ! [[ "${rayon_threads}" =~ ^[0-9]+$ ]] || [[ "${rayon_threads}" -lt 1 ]]; then
    echo "cargo_agent: invalid FORMULA_RAYON_NUM_THREADS=${rayon_threads} (expected integer >= 1)" >&2
    exit 2
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
# Note: if the caller is cross-compiling (explicit `--target` or `CARGO_BUILD_TARGET`), don't force
# openssl-sys to use the host OpenSSL install; the correct target OpenSSL may not be present.
if [[ -z "${CI:-}" && -z "${OPENSSL_NO_VENDOR:-}" && -z "${FORMULA_OPENSSL_VENDOR:-}" ]]; then
  allow_system_openssl=1
  if [[ -n "${CARGO_BUILD_TARGET:-}" ]]; then
    allow_system_openssl=0
  else
    for arg in "$@"; do
      if [[ "${arg}" == "--" ]]; then
        break
      fi
      case "${arg}" in
        --target|--target=*)
          allow_system_openssl=0
          break
          ;;
      esac
    done
  fi

  if [[ "${allow_system_openssl}" == "1" ]]; then
    if command -v uname >/dev/null 2>&1 && [[ "$(uname -s)" == "Linux" ]]; then
      if command -v pkg-config >/dev/null 2>&1 && pkg-config --exists openssl; then
        export OPENSSL_NO_VENDOR=1
      fi
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

supports_jobs=0
case "${subcommand}" in
  bench|build|check|clippy|doc|install|run|rustc|test)
    supports_jobs=1
    ;;
esac

explicit_jobs=""
remapped_args=()
expect_pkg_name=false
expect_jobs=false
while [[ $# -gt 0 ]]; do
  arg="$1"
  shift

  # Cargo treats `--` as "end of cargo flags", so stop rewriting there.
  if [[ "${arg}" == "--" ]]; then
    remapped_args+=("${arg}")
    if [[ $# -gt 0 ]]; then
      remapped_args+=("$@")
    fi
    break
  fi

  if [[ "${expect_pkg_name}" == "true" ]]; then
    if [[ -n "${remap_from}" && "${arg}" == "${remap_from}" ]]; then
      remapped_args+=("${remap_to}")
    else
      remapped_args+=("${arg}")
    fi
    expect_pkg_name=false
    continue
  fi
  if [[ "${expect_jobs}" == "true" ]]; then
    explicit_jobs="${arg}"
    expect_jobs=false
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
    -j|--jobs)
      if [[ "${supports_jobs}" == "1" ]]; then
        expect_jobs=true
      else
        remapped_args+=("${arg}")
      fi
      ;;
    -j*)
      if [[ "${supports_jobs}" == "1" ]]; then
        explicit_jobs="${arg#-j}"
      else
        remapped_args+=("${arg}")
      fi
      ;;
    --jobs=*)
      if [[ "${supports_jobs}" == "1" ]]; then
        explicit_jobs="${arg#--jobs=}"
      else
        remapped_args+=("${arg}")
      fi
      ;;
    *)
      remapped_args+=("${arg}")
      ;;
  esac
done
if [[ "${expect_pkg_name}" == "true" ]]; then
  echo "cargo_agent: missing value for -p/--package" >&2
  exit 2
fi
if [[ "${expect_jobs}" == "true" ]]; then
  echo "cargo_agent: missing value for -j/--jobs" >&2
  exit 2
fi
set -- "${remapped_args[@]}"

# If the caller explicitly passed `-j/--jobs`, treat that as authoritative. Also mark it as
# "explicitly configured" so we don't override it for `cargo test` runs.
if [[ -n "${explicit_jobs}" ]]; then
  if ! [[ "${explicit_jobs}" =~ ^[0-9]+$ ]] || [[ "${explicit_jobs}" -lt 1 ]]; then
    echo "cargo_agent: invalid -j/--jobs value: ${explicit_jobs}" >&2
    exit 2
  fi
  jobs="${explicit_jobs}"
  caller_jobs_env="${jobs}"
  export CARGO_BUILD_JOBS="${jobs}"
  if [[ -z "${orig_rayon_num_threads}" && -z "${orig_formula_rayon_num_threads}" ]]; then
    export RAYON_NUM_THREADS="${jobs}"
  fi
  # Keep Make/CMake build scripts in sync with the chosen -j level when MAKEFLAGS is not explicitly
  # configured (common case: agent-init defaulted it to just `-j4`).
  if [[ -z "${MAKEFLAGS:-}" || "${MAKEFLAGS}" =~ ^-j[0-9]+$ ]]; then
    export MAKEFLAGS="-j${jobs}"
  fi
fi

# Further reduce concurrency for test runs when callers haven't explicitly opted into
# higher parallelism. This avoids sporadic rustc panics like:
# "failed to spawn helper thread: Resource temporarily unavailable" (EAGAIN)
if [[ "${subcommand}" == "test" ]]; then
  # Record the pre-clamp job count so we can keep Rayon aligned when it was already set to track
  # the prior `jobs` value (common when environments default `RAYON_NUM_THREADS=$CARGO_BUILD_JOBS`).
  jobs_before_test_override="${jobs}"

  # If the caller explicitly passed `-j/--jobs`, respect it even if
  # `FORMULA_CARGO_TEST_JOBS` is set.
  if [[ -z "${explicit_jobs}" ]]; then
    if [[ -n "${FORMULA_CARGO_TEST_JOBS:-}" ]]; then
      if ! [[ "${FORMULA_CARGO_TEST_JOBS}" =~ ^[0-9]+$ ]] || [[ "${FORMULA_CARGO_TEST_JOBS}" -lt 1 ]]; then
        echo "cargo_agent: invalid FORMULA_CARGO_TEST_JOBS=${FORMULA_CARGO_TEST_JOBS} (expected integer >= 1)" >&2
        exit 2
      fi
      jobs="${FORMULA_CARGO_TEST_JOBS}"
    elif [[ -z "${caller_jobs_env}" ]]; then
      jobs="1"
    fi
  fi

  export CARGO_BUILD_JOBS="${jobs}"
  if [[ -z "${orig_formula_rayon_num_threads}" ]]; then
    # If Rayon threads were not explicitly configured (or were set to track the previous `jobs`
    # value), keep it aligned with the chosen `jobs` to avoid excessive thread creation under load.
    if [[ -z "${orig_rayon_num_threads}" || "${orig_rayon_num_threads}" == "${jobs_before_test_override}" ]]; then
      export RAYON_NUM_THREADS="${jobs}"
    fi
  fi
  if [[ -z "${MAKEFLAGS:-}" || "${MAKEFLAGS}" =~ ^-j[0-9]+$ ]]; then
    export MAKEFLAGS="-j${jobs}"
  fi
fi

# For test runs, cap RUST_TEST_THREADS to avoid spawning hundreds of threads.
#
# We set this *after* potentially reducing `jobs` for `cargo test` so the default thread count
# matches the chosen parallelism level. This reduces the chance of EAGAIN ("Resource temporarily
# unavailable") failures on multi-agent hosts when multiple test binaries run concurrently.
if [[ "${subcommand}" == "test" && -z "${RUST_TEST_THREADS:-}" ]]; then
  rust_test_threads="${FORMULA_RUST_TEST_THREADS:-}"
  if [[ -n "${rust_test_threads}" ]]; then
    if ! [[ "${rust_test_threads}" =~ ^[0-9]+$ ]] || [[ "${rust_test_threads}" -lt 1 ]]; then
      echo "cargo_agent: invalid FORMULA_RUST_TEST_THREADS=${rust_test_threads} (expected integer >= 1)" >&2
      exit 2
    fi
  fi
  if [[ -z "${rust_test_threads}" ]]; then
    rust_test_threads=$(( nproc_val < 16 ? nproc_val : 16 ))

    # Clamp based on the chosen cargo `-j` level (default: 4 * jobs).
    max_by_jobs=$(( jobs * 4 ))
    if [[ "${max_by_jobs}" -lt 1 ]]; then
      max_by_jobs=1
    fi
    if [[ "${rust_test_threads}" -gt "${max_by_jobs}" ]]; then
      rust_test_threads="${max_by_jobs}"
    fi
  fi
  export RUST_TEST_THREADS="${rust_test_threads}"
fi

# Make: default to matching the chosen Cargo parallelism.
#
# Some Rust crates invoke Make/CMake/etc from their build scripts. When Make is allowed
# to spawn one worker per core it can contribute to process/thread spawn failures on
# multi-agent hosts. Default to `-j<jobs>` unless the caller explicitly set MAKEFLAGS.
if [[ -z "${MAKEFLAGS:-}" ]]; then
  export MAKEFLAGS="-j${jobs}"
fi

# Limit rustc's internal codegen parallelism on multi-agent hosts.
#
# Even with `cargo -j` and Rayon's thread pool capped, rustc/LLVM may still spawn helper threads
# and/or one worker thread per codegen unit. Under high system load this can fail with
# "Resource temporarily unavailable" (EAGAIN) and crash compilation.
#
# Default to a single codegen unit for tests to reduce thread usage. Callers can override via the
# standard Cargo env vars (`CARGO_PROFILE_{dev,test,release,bench}_CODEGEN_UNITS`).
default_codegen_units="${jobs}"
if [[ "${subcommand}" == "test" ]]; then
  default_codegen_units="1"
fi
if [[ -z "${CARGO_PROFILE_DEV_CODEGEN_UNITS:-}" ]]; then
  export CARGO_PROFILE_DEV_CODEGEN_UNITS="${default_codegen_units}"
fi
if [[ -z "${CARGO_PROFILE_TEST_CODEGEN_UNITS:-}" ]]; then
  export CARGO_PROFILE_TEST_CODEGEN_UNITS="${default_codegen_units}"
fi
if [[ -z "${CARGO_PROFILE_RELEASE_CODEGEN_UNITS:-}" ]]; then
  export CARGO_PROFILE_RELEASE_CODEGEN_UNITS="${default_codegen_units}"
fi
if [[ -z "${CARGO_PROFILE_BENCH_CODEGEN_UNITS:-}" ]]; then
  export CARGO_PROFILE_BENCH_CODEGEN_UNITS="${default_codegen_units}"
fi

# Callers (or their shell environment) may set `RUSTFLAGS` with `-C codegen-units=...`.
# Cargo appends `RUSTFLAGS` at the end of the rustc command line, which can override
# `CARGO_PROFILE_*_CODEGEN_UNITS` (including our default of 1 for tests).
#
# To ensure our chosen test codegen-units value actually takes effect, append it to RUSTFLAGS so it
# wins as the final `-C codegen-units=...` flag.
if [[ "${subcommand}" == "test" ]]; then
  codegen_units_rustflags="${CARGO_PROFILE_TEST_CODEGEN_UNITS:-}"
  if [[ -n "${codegen_units_rustflags}" ]]; then
    if [[ -z "${RUSTFLAGS:-}" ]]; then
      export RUSTFLAGS="-C codegen-units=${codegen_units_rustflags}"
    else
      export RUSTFLAGS="${RUSTFLAGS} -C codegen-units=${codegen_units_rustflags}"
    fi
  fi
fi

# Limit lld's internal thread pool on multi-agent hosts.
#
# When the Rust toolchain is configured to link via `-fuse-ld=lld`, lld will spawn a thread pool by
# default. Under heavy system load / strict process limits this can fail with errors like:
# - `terminate called after throwing an instance of 'std::system_error' what(): Resource temporarily unavailable`
# - `ThreadPoolExecutor::ThreadPoolExecutor` in the crash backtrace
#
# Force single-threaded linking by default on Linux by passing `--threads=1` through the cc driver's
# `-Wl,` passthrough.
#
# Note: We only do this for the default (host) target, since `-Wl,` isn't understood by all linker
# invocation paths (e.g. some `--target wasm32-*` toolchains).
cargo_target=""
if [[ -n "${CARGO_BUILD_TARGET:-}" ]]; then
  cargo_target="${CARGO_BUILD_TARGET}"
else
  expect_target_value=false
  for arg in "$@"; do
    if [[ "${arg}" == "--" ]]; then
      break
    fi
    if [[ "${expect_target_value}" == "true" ]]; then
      cargo_target="${arg}"
      expect_target_value=false
      continue
    fi
    case "${arg}" in
      --target)
        expect_target_value=true
        ;;
      --target=*)
        cargo_target="${arg#--target=}"
        ;;
    esac
  done
fi

# Validate `FORMULA_LLD_THREADS` even when we won't apply it.
#
# The `--threads=<n>` linker flag is only injected for host builds below (since `-Wl,` isn't
# understood by all cross-toolchains). Still, if callers explicitly configure `FORMULA_LLD_THREADS`,
# fail fast when it's invalid rather than silently ignoring the typo.
lld_threads_env="${FORMULA_LLD_THREADS:-}"
if [[ -n "${lld_threads_env}" && "${lld_threads_env}" != "0" && "${lld_threads_env}" != "off" && "${lld_threads_env}" != "unlimited" ]]; then
  if ! [[ "${lld_threads_env}" =~ ^[0-9]+$ ]]; then
    echo "cargo_agent: invalid FORMULA_LLD_THREADS=${lld_threads_env} (expected integer, 0, off, or unlimited)" >&2
    exit 2
  fi
fi

if [[ -z "${cargo_target}" ]]; then
  lld_threads="${FORMULA_LLD_THREADS:-}"
  if [[ -z "${lld_threads}" ]]; then
    uname_s="$(uname -s 2>/dev/null || echo "")"
    if [[ "${uname_s}" == "Linux" ]]; then
      lld_threads="1"
    fi
  fi

  if [[ -n "${lld_threads}" && "${lld_threads}" != "0" && "${lld_threads}" != "off" && "${lld_threads}" != "unlimited" ]]; then
    if ! [[ "${lld_threads}" =~ ^[0-9]+$ ]]; then
      echo "cargo_agent: invalid FORMULA_LLD_THREADS=${lld_threads} (expected integer, 0, off, or unlimited)" >&2
      exit 2
    fi
    if [[ "${RUSTFLAGS:-}" != *"--threads="* ]]; then
      if [[ -z "${RUSTFLAGS:-}" ]]; then
        export RUSTFLAGS="-C link-arg=-Wl,--threads=${lld_threads}"
      else
        export RUSTFLAGS="${RUSTFLAGS} -C link-arg=-Wl,--threads=${lld_threads}"
      fi
    fi
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

echo "cargo_agent: jobs=${jobs} as=${limit_as} test_threads=${RUST_TEST_THREADS:-auto} rayon=${RAYON_NUM_THREADS:-auto}" >&2

run_cargo_cmd=("${cargo_cmd[@]}")
if [[ -n "${limit_as}" && "${limit_as}" != "0" && "${limit_as}" != "off" && "${limit_as}" != "unlimited" ]]; then
  if [[ -x "${repo_root}/scripts/run_limited.sh" ]]; then
    run_cargo_cmd=(bash "${repo_root}/scripts/run_limited.sh" --as "${limit_as}" -- "${cargo_cmd[@]}")
  fi
fi

# Retry on transient EAGAIN failures that are common on multi-agent hosts.
#
## Common symptoms
# - rustc panic: "failed to spawn work/helper thread: Resource temporarily unavailable"
# - linker (rust-lld/lld) abort: "std::system_error ... Resource temporarily unavailable"
# - test execution failure: "could not execute process ... Resource temporarily unavailable (os error 11)"
# - rustc execution failure: "could not execute process `rustc ...` (never executed)"
#
# Some agent hosts also run with `RUSTC_WRAPPER=sccache`. In practice the sccache daemon can die
# under heavy load, causing builds to fail with:
# - `sccache: error: failed to execute compile`
# - `Failed to send data to or receive data from server`
# In that case, retry once more with `RUSTC_WRAPPER` disabled so builds can proceed without the
# cache rather than failing spuriously.
#
# Retrying after a short backoff usually succeeds once other agents finish their compile/test bursts.
max_attempts="${FORMULA_CARGO_RETRY_ATTEMPTS:-5}"
if ! [[ "${max_attempts}" =~ ^[0-9]+$ ]] || [[ "${max_attempts}" -lt 1 ]]; then
  echo "cargo_agent: invalid FORMULA_CARGO_RETRY_ATTEMPTS=${max_attempts} (expected integer >= 1)" >&2
  exit 2
fi
attempt=1
disabled_rustc_wrapper_for_retry=false
while true; do
  # Capture output for error pattern detection while still streaming it to the console.
  tmp_log="$(mktemp -t cargo_agent.XXXXXX)"
  set +e
  # Preserve stdout/stderr streams for callers that capture machine-readable output (e.g. JSON)
  # while still recording a combined log for retry detection.
  exec 3> >(tee -a "${tmp_log}")
  tee_stdout_pid=$!
  exec 4> >(tee -a "${tmp_log}" >&2)
  tee_stderr_pid=$!
  "${run_cargo_cmd[@]}" 1>&3 2>&4
  status=$?
  exec 3>&- 4>&-
  wait "${tee_stdout_pid}" "${tee_stderr_pid}" 2>/dev/null || true
  set -e

  if [[ "${status}" -eq 0 ]]; then
    rm -f "${tmp_log}"
    exit 0
  fi

  retryable=false
  sccache_failed=false
  if grep -Eq "(Resource temporarily unavailable|os error 11)" "${tmp_log}"; then
    retryable=true
  elif grep -Eq "(rust-lld|ld\\.lld)" "${tmp_log}" \
    && { grep -q "__throw_system_error" "${tmp_log}" || grep -q "ThreadPoolExecutor" "${tmp_log}"; }; then
    retryable=true
  elif grep -Eq "(^sccache: error:|Failed to send data to or receive data from server)" "${tmp_log}"; then
    retryable=true
    sccache_failed=true
  fi

  if [[ "${retryable}" == "true" ]]; then
    if [[ "${attempt}" -ge "${max_attempts}" ]]; then
      rm -f "${tmp_log}"
      exit "${status}"
    fi

    if [[ "${sccache_failed}" == "true" && "${disabled_rustc_wrapper_for_retry}" == "false" ]]; then
      echo "cargo_agent: sccache failure; retrying with sccache disabled" >&2
      # Override any global Cargo config (`build.rustc-wrapper = "sccache"`) by forcing a benign
      # wrapper (`env`) that simply executes the underlying rustc.
      #
      # Note: Setting these vars to the empty string does *not* reliably override a configured
      # rustc wrapper; Cargo treats empty values as "unset" and can fall back to the global config.
      env_wrapper="$(command -v env 2>/dev/null || true)"
      if [[ -z "${env_wrapper}" ]]; then
        env_wrapper="env"
      fi
      export RUSTC_WRAPPER="${env_wrapper}"
      export RUSTC_WORKSPACE_WRAPPER="${env_wrapper}"
      export CARGO_BUILD_RUSTC_WRAPPER="${env_wrapper}"
      export CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER="${env_wrapper}"
      disabled_rustc_wrapper_for_retry=true
    fi

    # Exponential-ish backoff with jitter.
    base=$(( 1 << (attempt - 1) ))
    sleep_s=$(( base + RANDOM % (base + 1) ))
    echo "cargo_agent: transient resource exhaustion (EAGAIN); retrying in ${sleep_s}s (attempt ${attempt}/${max_attempts})" >&2
    rm -f "${tmp_log}"
    sleep "${sleep_s}"
    attempt=$(( attempt + 1 ))
    continue
  fi

  rm -f "${tmp_log}"
  exit "${status}"
done
