#!/bin/bash
# Returns appropriate -j value based on current memory pressure
# 
# Usage: 
#   cargo build -j$(./scripts/smart-jobs.sh)
#   make -j$(./scripts/smart-jobs.sh)
#
# The script considers:
#   - Available memory (not total)
#   - Assumes each compile job needs ~1.5GB headroom
#   - Clamps between MIN_JOBS and MAX_JOBS

MIN_JOBS=2
MAX_JOBS=16
GB_PER_JOB=2  # Conservative: 2GB per parallel job

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

# Calculate jobs based on available memory
JOBS=$((AVAIL_GB / GB_PER_JOB))

# Clamp to bounds
if [ "$JOBS" -lt "$MIN_JOBS" ]; then
  JOBS=$MIN_JOBS
elif [ "$JOBS" -gt "$MAX_JOBS" ]; then
  JOBS=$MAX_JOBS
fi

echo $JOBS
