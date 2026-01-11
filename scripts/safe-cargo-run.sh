#!/bin/bash
# Memory-aware wrapper for cargo run
# Usage: ./scripts/safe-cargo-run.sh [cargo run arguments...]
#
# Examples:
#   ./scripts/safe-cargo-run.sh -p formula-engine --bin perf_bench --release
#   ./scripts/safe-cargo-run.sh --release -- --help

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# Use a repo-local cargo home by default to avoid lock contention on ~/.cargo
# when many agents build in parallel. Preserve any user/CI override.
if [ -z "${CARGO_HOME:-}" ]; then
  export CARGO_HOME="$REPO_ROOT/target/cargo-home"
fi
mkdir -p "$CARGO_HOME"

# Get smart job count
if [ -x "$SCRIPT_DIR/smart-jobs.sh" ]; then
  JOBS=$("$SCRIPT_DIR/smart-jobs.sh")
else
  JOBS=${CARGO_BUILD_JOBS:-4}
fi

# Print to stderr so stdout remains usable for program output (e.g. JSON).
echo "🏃 Running with -j${JOBS} (based on available memory)..." >&2
cargo run -j"$JOBS" "$@"

