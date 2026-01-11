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
