#!/usr/bin/env bash
set -euo pipefail

# Run any command under OS-enforced resource limits (RLIMIT_AS).
#
# This is the simplest reliable way to cap memory - no cgroups, no systemd, works everywhere.
# When a process exceeds its address space limit, the kernel kills it automatically.
#
# Usage:
#   scripts/run_limited.sh --as 8G -- cargo build --release
#   scripts/run_limited.sh --as 4G -- node scripts/build.mjs
#   LIMIT_AS=8G scripts/run_limited.sh -- cargo test
#
# Based on fastrender's approach (simpler than cgroups/systemd-run).

usage() {
  cat <<'EOF'
usage: scripts/run_limited.sh [--as <size>] [--cpu <secs>] -- <command...>

Limits:
  --as <size>     Address-space (virtual memory) limit. Example: 8G, 4096M.
  --cpu <secs>    CPU time limit (seconds).

Environment defaults (optional):
  LIMIT_AS (default: 8G)
  LIMIT_CPU

Size suffixes: K, M, G, T (case-insensitive). Plain numbers = MiB.
Use "unlimited" or "off" to disable a limit.

Examples:
  scripts/run_limited.sh --as 8G -- cargo build -j4
  scripts/run_limited.sh --as 4G -- npm run build
  LIMIT_AS=16G scripts/run_limited.sh -- cargo test -p formula-engine
EOF
}

to_kib() {
  local raw="${1:-}"
  raw="${raw//[[:space:]]/}"
  raw="$(printf '%s' "${raw}" | tr '[:upper:]' '[:lower:]')"
  raw="${raw%ib}"
  raw="${raw%b}"

  if [[ "${raw}" =~ ^[0-9]+$ ]]; then
    echo $((raw * 1024))  # Plain number = MiB
    return 0
  fi

  if [[ "${raw}" =~ ^([0-9]+)([kmgt])$ ]]; then
    local n="${BASH_REMATCH[1]}"
    local unit="${BASH_REMATCH[2]}"
    case "${unit}" in
      k) echo $((n)) ;;
      m) echo $((n * 1024)) ;;
      g) echo $((n * 1024 * 1024)) ;;
      t) echo $((n * 1024 * 1024 * 1024)) ;;
    esac
    return 0
  fi
  return 1
}

to_bytes() {
  local kib
  kib="$(to_kib "${1:-}")" || return 1
  echo $((kib * 1024))
}

# Defaults
AS="${LIMIT_AS:-8G}"
CPU="${LIMIT_CPU:-}"

while [[ $# -gt 0 ]]; do
  case "$1" in
    -h|--help) usage; exit 0 ;;
    --as) AS="${2:-}"; shift 2 ;;
    --cpu) CPU="${2:-}"; shift 2 ;;
    --no-as) AS=""; shift ;;
    --no-cpu) CPU=""; shift ;;
    --) shift; break ;;
    *) break ;;
  esac
done

if [[ $# -lt 1 ]]; then
  usage
  exit 2
fi

cmd=("$@")

# `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
# globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift for
# `cargo` invocations that use this wrapper directly.
#
# Clear it so `cargo` respects the repo's pinned toolchain by default (consistent with
# `scripts/cargo_agent.sh`).
if [[ "${cmd[0]}" == "cargo" && -n "${RUSTUP_TOOLCHAIN:-}" ]]; then
  repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
  if [[ -f "${repo_root}/rust-toolchain.toml" ]]; then
    unset RUSTUP_TOOLCHAIN
  fi
  unset repo_root
fi

# macOS/Darwin: ulimit -v doesn't work reliably. Skip limits.
# Windows (Git Bash): Also doesn't support ulimit properly.
uname_s="$(uname -s 2>/dev/null || echo "")"
case "${uname_s}" in
  Darwin|MINGW*|MSYS*|CYGWIN*)
    exec "${cmd[@]}"
    ;;
esac

# Check if any limit is actually set
any_limit=false
if [[ -n "${AS}" && "${AS}" != "0" && "${AS}" != "unlimited" && "${AS}" != "off" ]]; then any_limit=true; fi
if [[ -n "${CPU}" && "${CPU}" != "0" && "${CPU}" != "unlimited" && "${CPU}" != "off" ]]; then any_limit=true; fi

if [[ "${any_limit}" == "false" ]]; then
  exec "${cmd[@]}"
fi

# Resolve cargo through rustup shim to avoid address space issues with the shim itself
if [[ -n "${AS}" && "${AS}" != "0" && "${AS}" != "unlimited" && "${AS}" != "off" ]] \
  && [[ "${cmd[0]}" == "cargo" ]] \
  && command -v rustup >/dev/null 2>&1
then
  cargo_shim="$(command -v cargo || true)"
  if [[ -n "${cargo_shim}" ]]; then
    cargo_target="${cargo_shim}"
    if [[ -L "${cargo_shim}" ]]; then
      cargo_target="$(readlink "${cargo_shim}" 2>/dev/null || echo "${cargo_shim}")"
    fi

    if [[ "${cargo_target}" == "rustup" || "${cargo_target}" == */rustup || "${cargo_shim}" == */.cargo/bin/cargo ]]; then
      toolchain=""
      if [[ ${#cmd[@]} -gt 1 && "${cmd[1]}" == +* ]]; then
        toolchain="${cmd[1]#+}"
        cmd=("${cmd[0]}" "${cmd[@]:2}")
      fi

      if [[ -n "${toolchain}" ]]; then
        resolved="$(rustup which --toolchain "${toolchain}" cargo 2>/dev/null || true)"
      else
        resolved="$(rustup which cargo 2>/dev/null || true)"
      fi

      if [[ -n "${resolved}" ]]; then
        cmd[0]="${resolved}"
        toolchain_bin="$(dirname "${resolved}")"
        export PATH="${toolchain_bin}:${PATH}"
      fi
    fi
  fi
fi

# Try prlimit first (more reliable hard limits)
prlimit_ok=0
if command -v prlimit >/dev/null 2>&1; then
  if prlimit --as=67108864 --cpu=1 -- true >/dev/null 2>&1; then
    prlimit_ok=1
  fi
fi

if [[ "${prlimit_ok}" -eq 1 ]]; then
  pl=(prlimit --pid $$)
  if [[ -n "${AS}" && "${AS}" != "0" && "${AS}" != "unlimited" && "${AS}" != "off" ]]; then
    as_bytes="$(to_bytes "${AS}")" || { echo "invalid --as size: ${AS}" >&2; exit 2; }
    pl+=(--as="${as_bytes}")
  fi
  if [[ -n "${CPU}" && "${CPU}" != "0" && "${CPU}" != "unlimited" && "${CPU}" != "off" ]]; then
    if ! [[ "${CPU}" =~ ^[0-9]+$ ]]; then
      echo "invalid --cpu seconds: ${CPU}" >&2
      exit 2
    fi
    pl+=(--cpu="${CPU}")
  fi

  if "${pl[@]}" >/dev/null 2>&1; then
    exec "${cmd[@]}"
  fi
fi

# Fallback: ulimit
if [[ -n "${AS}" && "${AS}" != "0" && "${AS}" != "unlimited" && "${AS}" != "off" ]]; then
  as_kib="$(to_kib "${AS}")" || { echo "invalid --as size: ${AS}" >&2; exit 2; }
  ulimit -v "${as_kib}" 2>/dev/null || true
fi
if [[ -n "${CPU}" && "${CPU}" != "0" && "${CPU}" != "unlimited" && "${CPU}" != "off" ]]; then
  if [[ "${CPU}" =~ ^[0-9]+$ ]]; then
    ulimit -t "${CPU}" 2>/dev/null || true
  fi
fi

exec "${cmd[@]}"
