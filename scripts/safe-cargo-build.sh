#!/bin/bash
# Memory-aware wrapper for cargo build
# Usage: ./scripts/safe-cargo-build.sh [cargo build arguments...]
#
# Examples:
#   ./scripts/safe-cargo-build.sh --release
#   ./scripts/safe-cargo-build.sh -p formula-engine
#   ./scripts/safe-cargo-build.sh --all-features

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Get smart job count
if [ -x "$SCRIPT_DIR/smart-jobs.sh" ]; then
  JOBS=$("$SCRIPT_DIR/smart-jobs.sh")
else
  JOBS=${CARGO_BUILD_JOBS:-4}
fi

echo "🦀 Building with -j${JOBS} (based on available memory)..."
cargo build -j"$JOBS" "$@"
