# Agent Development Environment Guide

## Overview

This guide is for **coding agents developing Formula**, not end-users. The development environment is:

- **Machine**: 192 vCPU, 1.5TB RAM, 110TB disk (Ubuntu Linux x64, headless, no GPU)
- **Concurrency**: Up to ~200 agents, each with their own repo copy
- **Critical constraint**: Memory. Disk/CPU are abundant; RAM exhaustion kills the machine.

---

## Memory Management (CRITICAL)

### The Math

```
Total RAM:     1,500 GB
System/Buffer:   100 GB (reserved)
Available:     1,400 GB
Agents:          ~200
Per-agent:        ~7 GB soft target
```

**7GB per agent is tight** for a full-stack TypeScript/Rust project. We must be disciplined.

### Memory Limits by Operation


| Operation           | Expected Peak | Limit | Notes                            |
| ------------------- | ------------- | ----- | -------------------------------- |
| Node.js process     | 512MB-2GB     | 4GB   | Use `--max-old-space-size`       |
| Rust compilation    | 2-8GB         | 8GB   | Depends on parallelism           |
| TypeScript check    | 500MB-2GB     | 2GB   | Can spike with large projects    |
| npm install         | 500MB-1GB     | 2GB   | Transient                        |
| Tests (unit)        | 200MB-1GB     | 2GB   |                                  |
| Tests (e2e/browser) | 500MB-2GB     | 2GB   | Headless Chrome                  |
| Total concurrent    | -             | 10GB  | Hard ceiling per agent workspace |


### Required Environment Variables

**Every agent MUST set these** in their shell before running commands:

```bash
# Add to ~/.bashrc or run at start of each session
export NODE_OPTIONS="--max-old-space-size=3072"  # 3GB limit for Node
export CARGO_BUILD_JOBS=4                         # Limit Rust parallelism
export MAKEFLAGS="-j4"                            # Limit make parallelism
export RUSTFLAGS="-C codegen-units=4"             # Reduce Rust memory per crate
```

### Memory-Safe Command Patterns

#### npm/pnpm/yarn

```bash
# GOOD: Memory-limited
NODE_OPTIONS="--max-old-space-size=2048" npm install
NODE_OPTIONS="--max-old-space-size=3072" npm run build

# BAD: Unbounded
npm install  # Without NODE_OPTIONS set
```

#### Rust/Cargo

```bash
# GOOD: Limited parallelism (4 jobs uses ~4-6GB)
cargo build -j4
cargo test -j4

# ACCEPTABLE: Let environment variable handle it
cargo build  # If CARGO_BUILD_JOBS=4 is set

# BAD: Full parallelism (could use 50GB+ with -j192)
cargo build -j$(nproc)
cargo build --jobs 192
```

#### TypeScript

```bash
# GOOD: Bounded memory
NODE_OPTIONS="--max-old-space-size=2048" npx tsc

# For very large projects, run incrementally
npx tsc --incremental
```

---

## CPU/Parallelism Guidelines

The scheduler handles CPU contention well. The issue is that **parallelism correlates with memory**.

### Recommended `-j` Values


| Scenario               | Recommended      | Rationale                                  |
| ---------------------- | ---------------- | ------------------------------------------ |
| Rust compilation       | `-j4` to `-j8`   | Each rustc can use 1-2GB                   |
| Make/CMake             | `-j4` to `-j8`   | Safe default                               |
| npm scripts (parallel) | `-j4`            | `npm-run-all -p` or `concurrently`         |
| Jest/Vitest            | `--maxWorkers=4` | Test parallelism                           |
| Likely solo operation  | `-j8` to `-j16`  | When you're confident few others compiling |


### Adaptive Parallelism Script

Use this helper to pick parallelism based on current system load:

```bash
# Usage: ./scripts/smart-jobs.sh
# Returns a reasonable -j value based on current memory pressure

#!/bin/bash
# See scripts/smart-jobs.sh for full implementation
```

---

## Disk I/O

Disk is abundant (110TB). Don't worry about it, but be aware:

### Things That Are Fine

- Large node_modules per repo (each agent has own copy)
- Rust target directories (can be 5-20GB each)
- Generous caching
- Build artifacts

### Things to Avoid (cleanliness, not necessity)

```bash
# Periodically clean if disk somehow becomes an issue
cargo clean
rm -rf node_modules/.cache
rm -rf .next/cache  # if using Next.js
```

---

## Headless Development

### Confirmed Working (No Special Setup)


| Tool            | Status | Notes           |
| --------------- | ------ | --------------- |
| Node.js/npm     | ✅      | Works perfectly |
| Rust/Cargo      | ✅      | Works perfectly |
| TypeScript      | ✅      | Works perfectly |
| Jest/Vitest     | ✅      | Works perfectly |
| ESLint/Prettier | ✅      | Works perfectly |
| Git             | ✅      | Works perfectly |


### Requires Setup

#### Canvas/Graphics Testing

For testing Canvas rendering code:

```bash
# Install virtual framebuffer
sudo apt-get install -y xvfb

# Run tests with virtual display
xvfb-run --auto-servernum npm test

# Or for a persistent display
export DISPLAY=:99
Xvfb :99 -screen 0 1920x1080x24 &
npm test
```

#### Headless Browser Testing (Puppeteer/Playwright)

```bash
# Install dependencies
sudo apt-get install -y \
  libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 \
  libxkbcommon0 libxcomposite1 libxdamage1 libxfixes3 \
  libxrandr2 libgbm1 libpango-1.0-0 libcairo2 libasound2

# Playwright handles this automatically
npx playwright install-deps

# Always use headless mode
# In test config:
# browser: { headless: true }
```

#### Tauri Development

```bash
# Tauri can build without display, but needs libraries
sudo apt-get install -y \
  libgtk-3-dev libwebkit2gtk-4.0-dev libappindicator3-dev \
  librsvg2-dev patchelf

# Build works headless
cargo tauri build

# Dev server needs virtual display for window
xvfb-run cargo tauri dev
```

---

## Potential Blockers & Solutions

### Blocker 1: Canvas Unit Tests

**Problem**: `node-canvas` and similar require native graphics libraries.

**Solution**:

```bash
# Install Cairo and friends
sudo apt-get install -y \
  libcairo2-dev libjpeg-dev libpango1.0-dev \
  libgif-dev librsvg2-dev libpixman-1-dev

# Then npm install works
npm install canvas
```

### Blocker 2: Electron/Tauri GUI Tests

**Problem**: Need a display for integration tests.

**Solution**: Use Xvfb (see above), or mock at a lower level:

```bash
xvfb-run --auto-servernum npm run test:integration
```

### Blocker 3: WebGL

**Problem**: No GPU means no WebGL.

**Solution**: We use Canvas 2D, not WebGL. For any WebGL code:

- Use software rendering: `LIBGL_ALWAYS_SOFTWARE=1`
- Or mock WebGL context in tests
- Or skip WebGL tests in CI with `describe.skip`

### Blocker 4: Font Rendering Consistency

**Problem**: Different fonts than macOS/Windows affect text measurements.

**Solution**:

```bash
# Install common fonts
sudo apt-get install -y \
  fonts-liberation fonts-dejavu-core fonts-noto-core \
  fontconfig

# Clear font cache
fc-cache -f -v

# In tests, use explicitly installed fonts
```

### Blocker 5: File Watcher Limits

**Problem**: Too many watchers across 200 repos.

**Solution**:

```bash
# Increase inotify limits (system-wide, one-time setup)
echo "fs.inotify.max_user_watches=524288" | sudo tee -a /etc/sysctl.conf
echo "fs.inotify.max_user_instances=1024" | sudo tee -a /etc/sysctl.conf
sudo sysctl -p

# Alternatively, disable watchers in development
# Most build tools have a --no-watch or poll mode
```

### Blocker 6: Rust Build Cache Explosion

**Problem**: 200 repos × 10GB target dirs = 2TB.

**Solution**: Use sccache for shared compilation cache:

```bash
# Install sccache
cargo install sccache

# Configure (add to ~/.bashrc)
export RUSTC_WRAPPER=sccache
export SCCACHE_DIR=/shared/sccache  # Shared directory
export SCCACHE_CACHE_SIZE="50G"     # Shared cache limit
```

---

## Helper Scripts

### scripts/agent-init.sh

Run this at the start of each agent session:

```bash
#!/bin/bash
# Initialize agent environment with safe defaults

set -e

# Memory limits
export NODE_OPTIONS="--max-old-space-size=3072"
export CARGO_BUILD_JOBS=4
export MAKEFLAGS="-j4"
export RUSTFLAGS="-C codegen-units=4"

# Headless display (if not already set)
if [ -z "$DISPLAY" ]; then
  export DISPLAY=:99
  if ! pgrep -x Xvfb > /dev/null; then
    Xvfb :99 -screen 0 1920x1080x24 &
    sleep 1
  fi
fi

# Confirm settings
echo "Agent environment initialized:"
echo "  NODE_OPTIONS: $NODE_OPTIONS"
echo "  CARGO_BUILD_JOBS: $CARGO_BUILD_JOBS"
echo "  DISPLAY: $DISPLAY"
```

### scripts/smart-jobs.sh

Adaptive parallelism based on current memory:

```bash
#!/bin/bash
# Returns appropriate -j value based on memory pressure
# Usage: cargo build -j$(./scripts/smart-jobs.sh)

# Get available memory in GB
AVAIL_GB=$(free -g | awk '/^Mem:/{print $7}')

# Each compile job needs ~1.5GB headroom
MAX_JOBS=$((AVAIL_GB / 2))

# Clamp between 2 and 16
if [ $MAX_JOBS -lt 2 ]; then
  MAX_JOBS=2
elif [ $MAX_JOBS -gt 16 ]; then
  MAX_JOBS=16
fi

echo $MAX_JOBS
```

### scripts/check-memory.sh

Quick memory status:

```bash
#!/bin/bash
# Show current memory status

echo "=== Memory Status ==="
free -h
echo ""
echo "=== Top Memory Consumers ==="
ps aux --sort=-%mem | head -15
echo ""
echo "=== Recommendation ==="
AVAIL=$(free -g | awk '/^Mem:/{print $7}')
if [ $AVAIL -lt 50 ]; then
  echo "⚠️  LOW MEMORY: Only ${AVAIL}GB available. Use -j2 for builds."
elif [ $AVAIL -lt 200 ]; then
  echo "⚡ MODERATE: ${AVAIL}GB available. Use -j4 for builds."
else
  echo "✅ PLENTY: ${AVAIL}GB available. Use -j8 or higher."
fi
```

### scripts/safe-cargo-build.sh

Memory-aware Rust builds:

```bash
#!/bin/bash
# Wrapper for cargo build with memory awareness
# Usage: ./scripts/safe-cargo-build.sh [cargo args...]

JOBS=$(./scripts/smart-jobs.sh 2>/dev/null || echo 4)
echo "Building with -j$JOBS based on available memory..."
cargo build -j$JOBS "$@"
```

---

## Quick Reference Card

### Before Any Build

```bash
source scripts/agent-init.sh
```

### Safe Commands

```bash
# Node/TypeScript
npm install                           # OK with NODE_OPTIONS set
npm run build                         # OK with NODE_OPTIONS set
npm test                              # OK
npx tsc --incremental                 # Preferred for large projects

# Rust
cargo build -j4                       # Safe
cargo test -j4                        # Safe
./scripts/safe-cargo-build.sh        # Auto-detects

# Testing with display
xvfb-run npm run test:e2e            # For GUI tests
```

### Dangerous Commands (Avoid)

```bash
cargo build -j$(nproc)               # 192 parallel rustc = OOM
cargo build -j0                       # Unlimited = OOM
npm run build & npm run build        # Concurrent builds from same agent
```

### Emergency: System Under Memory Pressure

```bash
# Check what's using memory
./scripts/check-memory.sh

# Wait for your own processes to finish, or cancel them via Ctrl+C
# DO NOT use pkill - it kills processes across ALL agents!

# If your specific process is stuck, find its PID and kill only that:
ps aux | grep "your-specific-pattern"
kill <specific-pid>
```

---

## Summary


| Resource    | Constraint | Guidance                                                   |
| ----------- | ---------- | ---------------------------------------------------------- |
| **RAM**     | Critical   | Set NODE_OPTIONS, limit -j to 4-8, monitor usage           |
| **CPU**     | Abundant   | Let scheduler handle it; parallelism limited by RAM anyway |
| **Disk**    | Abundant   | Don't worry about it                                       |
| **GPU**     | None       | Use Xvfb, headless browsers, software rendering            |
| **Network** | Standard   | No special considerations                                  |


**Golden Rule**: When in doubt, use `-j4`. It's fast enough and won't cause problems.