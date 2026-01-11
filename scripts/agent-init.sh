#!/bin/bash
# Initialize agent environment with safe memory defaults
# Source this at the start of each agent session: source scripts/agent-init.sh

set -e

# ============================================================================
# Memory Limits (CRITICAL)
# ============================================================================

# Node.js: 3GB heap limit (leaves room for other processes)
export NODE_OPTIONS="--max-old-space-size=3072"

# Rust: Limit parallel compilation jobs (each can use 1-2GB)
export CARGO_BUILD_JOBS="${CARGO_BUILD_JOBS:-4}"

# Make: Limit parallel jobs
export MAKEFLAGS="${MAKEFLAGS:--j4}"

# Rust codegen: Fewer units = less memory, slightly slower
export RUSTFLAGS="${RUSTFLAGS:--C codegen-units=4}"

# ============================================================================
# Cargo Home Isolation (CRITICAL - prevents cross-agent ~/.cargo lock contention)
# ============================================================================

# Cargo defaults to using ~/.cargo for registry/index/git caches. With many agents
# building in parallel this creates heavy lock contention under ~/.cargo and can
# make builds flaky/slow. Default to a repo-local CARGO_HOME to isolate agents,
# but preserve any user/CI override.
if [ -z "${CARGO_HOME:-}" ]; then
  export CARGO_HOME="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)/target/cargo-home"
fi
mkdir -p "$CARGO_HOME"

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
  if [ -z "$DISPLAY" ]; then
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
if command -v Xvfb &> /dev/null; then
  setup_display
fi

# ============================================================================
# Confirmation
# ============================================================================

echo "╔════════════════════════════════════════════════════════════════╗"
echo "║  Agent Environment Initialized                                  ║"
echo "╠════════════════════════════════════════════════════════════════╣"
echo "║  NODE_OPTIONS:      ${NODE_OPTIONS}"
echo "║  CARGO_BUILD_JOBS:  ${CARGO_BUILD_JOBS}"
echo "║  CARGO_HOME:        ${CARGO_HOME}"
echo "║  MAKEFLAGS:         ${MAKEFLAGS}"
if [ -n "$DISPLAY" ]; then
echo "║  DISPLAY:           ${DISPLAY}"
fi
echo "╚════════════════════════════════════════════════════════════════╝"
