#!/bin/bash
# Returns an appropriate `-j` value based on current memory pressure.
# 
# Usage: 
#   cargo build -j$(./scripts/smart-jobs.sh)
#   make -j$(./scripts/smart-jobs.sh)
#
# The script considers:
#   - Available memory (not total)
#   - Our multi-agent host has many concurrent builds; default to conservative values
#     rather than scaling linearly with free RAM.

MIN_JOBS=1
MAX_JOBS=8

# Get available memory in GB
if command -v free &> /dev/null; then
  AVAIL_GB=$(free -g | awk '/^Mem:/{print $7}')
elif [ -f /proc/meminfo ]; then
  AVAIL_KB=$(grep MemAvailable /proc/meminfo | awk '{print $2}')
  AVAIL_GB=$((AVAIL_KB / 1024 / 1024))
else
  # Fallback: assume moderate availability
  AVAIL_GB=32
fi

# Choose a conservative job count from memory pressure bands.
#
# Note: this intentionally tops out at 8 so a single agent doesn't monopolize the host
# even when global free RAM is very high.
if [ "$AVAIL_GB" -lt 20 ]; then
  JOBS=1
elif [ "$AVAIL_GB" -lt 50 ]; then
  JOBS=2
elif [ "$AVAIL_GB" -lt 200 ]; then
  JOBS=4
else
  JOBS=8
fi

# Clamp to bounds
if [ "$JOBS" -lt "$MIN_JOBS" ]; then
  JOBS=$MIN_JOBS
elif [ "$JOBS" -gt "$MAX_JOBS" ]; then
  JOBS=$MAX_JOBS
fi

echo $JOBS
