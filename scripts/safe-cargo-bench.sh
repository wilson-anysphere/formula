#!/bin/bash
# Memory-aware wrapper for cargo bench
# Usage: ./scripts/safe-cargo-bench.sh [cargo bench arguments...]
#
# Examples:
#   ./scripts/safe-cargo-bench.sh -p formula-engine
#   ./scripts/safe-cargo-bench.sh --bench engine -- --help

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# Use a repo-local cargo home by default to avoid lock contention on ~/.cargo
# when many agents build in parallel. Preserve any user/CI override.
if [ -z "${CARGO_HOME:-}" ]; then
  export CARGO_HOME="$REPO_ROOT/target/cargo-home"
fi
mkdir -p "$CARGO_HOME"
mkdir -p "$CARGO_HOME/bin"
case ":$PATH:" in
  *":$CARGO_HOME/bin:"*) ;;
  *) export PATH="$CARGO_HOME/bin:$PATH" ;;
esac

# Get smart job count
if [ -x "$SCRIPT_DIR/smart-jobs.sh" ]; then
  JOBS=$("$SCRIPT_DIR/smart-jobs.sh")
else
  JOBS=${CARGO_BUILD_JOBS:-4}
fi

echo "📊 Benchmarking with -j${JOBS} (based on available memory)..."
cargo bench -j"$JOBS" "$@"

