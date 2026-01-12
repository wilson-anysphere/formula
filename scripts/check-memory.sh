#!/bin/bash
# Show current memory status and recommendations
# Usage: ./scripts/check-memory.sh

set -e

echo "╔════════════════════════════════════════════════════════════════╗"
echo "║  System Memory Status                                           ║"
echo "╚════════════════════════════════════════════════════════════════╝"
echo ""

# Memory overview
echo "=== Memory Overview ==="
free -h
echo ""

# Available memory in GB for calculations
AVAIL_GB=$(free -g | awk '/^Mem:/{print $7}')
TOTAL_GB=$(free -g | awk '/^Mem:/{print $2}')
USED_GB=$(free -g | awk '/^Mem:/{print $3}')
USED_PCT=$((USED_GB * 100 / TOTAL_GB))

echo "=== Top 15 Memory Consumers ==="
ps aux --sort=-%mem | head -16 | awk '{printf "%-10s %-8s %-6s %s\n", $1, $4"%", $6/1024"MB", $11}'
echo ""

echo "=== Rust/Node Processes ==="
echo "Rust compilations:"
pgrep -c rustc 2>/dev/null || echo "0"
echo "Node processes:"
pgrep -c node 2>/dev/null || echo "0"
echo ""

echo "=== Recommendation ==="
if [ "$AVAIL_GB" -lt 20 ]; then
  echo "🚨 CRITICAL: Only ${AVAIL_GB}GB available (${USED_PCT}% used)"
  echo "   → Pause new builds, wait for current ones to finish"
  echo "   → Use -j1 or -j2 if you must build"
  echo "   → Cancel your own builds with Ctrl+C if needed"
elif [ "$AVAIL_GB" -lt 50 ]; then
  echo "⚠️  LOW MEMORY: ${AVAIL_GB}GB available (${USED_PCT}% used)"
  echo "   → Use -j2 for Rust builds"
  echo "   → Avoid starting new heavy operations"
elif [ "$AVAIL_GB" -lt 200 ]; then
  echo "⚡ MODERATE: ${AVAIL_GB}GB available (${USED_PCT}% used)"
  echo "   → Use -j4 for builds (default)"
  echo "   → Normal operations are fine"
elif [ "$AVAIL_GB" -lt 500 ]; then
  echo "✅ GOOD: ${AVAIL_GB}GB available (${USED_PCT}% used)"
  echo "   → Use -j8 for builds"
  echo "   → System is healthy"
else
  echo "🚀 EXCELLENT: ${AVAIL_GB}GB available (${USED_PCT}% used)"
  echo "   → Use -j8 to -j16 for builds"
  echo "   → Plenty of headroom"
fi

echo ""

# Prefer explicitly configured job counts when present, otherwise fall back to a best-effort
# recommendation based on current memory pressure.
if [ -n "${FORMULA_CARGO_JOBS:-}" ]; then
  echo "Suggested -j value: ${FORMULA_CARGO_JOBS} (from FORMULA_CARGO_JOBS)"
elif [ -n "${CARGO_BUILD_JOBS:-}" ]; then
  echo "Suggested -j value: ${CARGO_BUILD_JOBS} (from CARGO_BUILD_JOBS)"
else
  echo "Suggested -j value: $(./scripts/smart-jobs.sh 2>/dev/null || echo 4)"
fi
