#!/bin/bash
# Run a command with a virtual display, handling concurrent Xvfb instances
# Usage: ./scripts/xvfb-run-safe.sh <command> [args...]
#
# Examples:
#   ./scripts/xvfb-run-safe.sh npm run test:e2e
#   ./scripts/xvfb-run-safe.sh cargo tauri dev

set -e

if [ $# -eq 0 ]; then
  echo "Usage: $0 <command> [args...]"
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# Use a repo-local cargo home by default to avoid lock contention on ~/.cargo
# when many agents build concurrently. Preserve any user/CI override.
#
# Note: some agent runners pre-set `CARGO_HOME=$HOME/.cargo`. Treat that value as
# "unset" for our purposes so we still get per-repo isolation by default.
# In CI we respect `CARGO_HOME` even if it points at `$HOME/.cargo` so CI can use
# shared caching.
# To explicitly keep `CARGO_HOME=$HOME/.cargo` in local runs, set
# `FORMULA_ALLOW_GLOBAL_CARGO_HOME=1`.
DEFAULT_GLOBAL_CARGO_HOME="${HOME:-/root}/.cargo"
if [ -z "${CARGO_HOME:-}" ] || {
  [ -z "${CI:-}" ] &&
    [ -z "${FORMULA_ALLOW_GLOBAL_CARGO_HOME:-}" ] &&
    [ "${CARGO_HOME}" = "${DEFAULT_GLOBAL_CARGO_HOME}" ];
}; then
  export CARGO_HOME="$REPO_ROOT/target/cargo-home"
fi
mkdir -p "$CARGO_HOME"
mkdir -p "$CARGO_HOME/bin"
case ":$PATH:" in
  *":$CARGO_HOME/bin:"*) ;;
  *) export PATH="$CARGO_HOME/bin:$PATH" ;;
esac

# Find an available display number
find_display() {
  for d in $(seq 99 199); do
    if [ ! -e "/tmp/.X${d}-lock" ]; then
      echo $d
      return 0
    fi
  done
  echo "99"  # Fallback
}

# Check if we already have a display
if [ -n "$DISPLAY" ] && xdpyinfo >/dev/null 2>&1; then
  # Display already available, just run the command
  exec "$@"
fi

# Find available display
DISPLAY_NUM=$(find_display)
export DISPLAY=:$DISPLAY_NUM

# Start Xvfb
Xvfb :$DISPLAY_NUM -screen 0 1920x1080x24 >/dev/null 2>&1 &
XVFB_PID=$!

# Wait for Xvfb to start
sleep 0.5

# Cleanup function
cleanup() {
  if [ -n "$XVFB_PID" ] && kill -0 $XVFB_PID 2>/dev/null; then
    kill $XVFB_PID 2>/dev/null || true
  fi
}
trap cleanup EXIT

# Run the command
"$@"
