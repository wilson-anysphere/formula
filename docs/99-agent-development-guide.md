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

### How We Enforce Memory Limits

We use **`RLIMIT_AS` (address space limit)** - a simple POSIX feature that:
- Requires **no cgroups, no systemd, no special privileges**
- Works on **any Linux** (and degrades gracefully on macOS)
- **Kernel enforces it automatically** - processes die when they exceed the limit

This is simpler and more robust than cgroups/systemd-run.

### Memory Limits by Operation


| Operation           | Expected Peak | Limit | Notes                            |
| ------------------- | ------------- | ----- | -------------------------------- |
| Node.js process     | 512MB-2GB     | 4GB   | Use `--max-old-space-size`       |
| Rust compilation    | 2-8GB         | 12GB  | Use `scripts/cargo_agent.sh`     |
| TypeScript check    | 500MB-2GB     | 2GB   | Can spike with large projects    |
| npm install         | 500MB-1GB     | 2GB   | Transient                        |
| Tests (unit)        | 200MB-1GB     | 2GB   |                                  |
| Tests (e2e/browser) | 500MB-2GB     | 2GB   | Headless Chrome                  |
| Total concurrent    | -             | 12GB  | Hard ceiling per agent workspace |


### Required: Always Use Wrapper Scripts for Cargo

**MANDATORY** - use `scripts/cargo_agent.sh` for all cargo commands:

```bash
# CORRECT:
bash scripts/cargo_agent.sh build --release
bash scripts/cargo_agent.sh test -p formula-engine
bash scripts/cargo_agent.sh check

# WRONG - can exhaust RAM:
cargo build
cargo test
```

The wrapper script:
1. Enforces a **12GB address space limit** via `RLIMIT_AS`
2. Limits parallelism to **-j4** by default
3. Caps **RUST_TEST_THREADS** to avoid spawning hundreds of threads
4. Uses a **repo-local CARGO_HOME** to avoid lock contention

### Environment Setup (Optional but Recommended)

```bash
# Initialize safe defaults (sets NODE_OPTIONS, CARGO_BUILD_JOBS, etc.)
source scripts/agent-init.sh  # or: . scripts/agent-init.sh
```

### Running Commands with Memory Limits

#### `scripts/run_limited.sh` - Universal Memory Cap

For **any** command that might use a lot of memory:

```bash
# Run any command with an 8GB address space limit
bash scripts/run_limited.sh --as 8G -- npm run build
bash scripts/run_limited.sh --as 4G -- node heavy-script.js

# Environment variable alternative
LIMIT_AS=8G bash scripts/run_limited.sh -- npm test
```

#### `scripts/cargo_agent.sh` - Cargo Wrapper (Preferred)

For **all cargo commands**, use the wrapper:

```bash
bash scripts/cargo_agent.sh build --release
bash scripts/cargo_agent.sh test -p formula-engine --lib
bash scripts/cargo_agent.sh check -p formula-xlsx
bash scripts/cargo_agent.sh bench --bench perf_regressions
```

Environment variables to tune behavior:
- `FORMULA_CARGO_JOBS` - parallelism (default: 4)
- `FORMULA_CARGO_LIMIT_AS` - address space limit (default: 12G)
- `FORMULA_RUST_TEST_THREADS` - test parallelism (default: min(nproc, 16))

#### npm/pnpm/yarn

```bash
# GOOD: Memory-limited (NODE_OPTIONS set by agent-init.sh)
source scripts/agent-init.sh
npm install
npm run build

# Or explicitly:
NODE_OPTIONS="--max-old-space-size=3072" npm run build

# Or with run_limited.sh:
bash scripts/run_limited.sh --as 4G -- npm run build
```

### Cargo Home Isolation

The wrapper scripts automatically use a **repo-local CARGO_HOME** at `target/cargo-home` to avoid
lock contention when ~200 agents build concurrently. This means:

- Each repo has its own registry cache (more disk, but disk is abundant)
- `cargo clean` or deleting `target/` also clears the local registry cache
- `cargo install` binaries go to `target/cargo-home/bin`

To use a shared cache (CI), set `CARGO_HOME` before running:

```bash
export CARGO_HOME="$HOME/.cargo"
export FORMULA_ALLOW_GLOBAL_CARGO_HOME=1
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
- Per-repo Cargo registry caches (when `CARGO_HOME=target/cargo-home`)
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

### Vitest + WASM builds

The root Vitest config (`vitest.config.ts`) runs a global setup step
(`scripts/vitest.global-setup.mjs`) that ensures a Node-compatible wasm-bindgen
build of `crates/formula-wasm` exists. This may trigger a Rust/wasm-pack build,
which is slow and memory intensive.

For test runs that don't touch the WASM engine (e.g. `packages/llm`,
`packages/ai-tools`), you can skip the setup step:

```bash
FORMULA_SKIP_WASM_BUILD=1 pnpm vitest run packages/llm/src/*.test.ts
```

If you need engine-backed tests, omit the variable (or run
`node scripts/build-formula-wasm-node.mjs` once).


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

| Script | Purpose |
|--------|---------|
| `scripts/agent-init.sh` | Source at session start - sets NODE_OPTIONS, CARGO_HOME, etc. |
| `scripts/cargo_agent.sh` | **Use for all cargo commands** - enforces memory limits |
| `scripts/run_limited.sh` | Run any command with RLIMIT_AS cap |
| `scripts/smart-jobs.sh` | Returns adaptive -j value based on free RAM |
| `scripts/check-memory.sh` | Show memory status and recommendations |

### How Memory Limits Work

`scripts/run_limited.sh` uses `RLIMIT_AS` (via `prlimit` or `ulimit -v`):

```bash
# This sets a hard 8GB address space limit
# If the process tries to allocate more, the kernel kills it
bash scripts/run_limited.sh --as 8G -- cargo build

# Equivalent to:
prlimit --as=$((8*1024*1024*1024)) --pid $$ && cargo build
# or (fallback):
ulimit -v $((8*1024*1024)) && cargo build
```

This is **simpler and more reliable** than cgroups/systemd-run:
- No special privileges needed
- Works on any Linux
- Kernel enforces it automatically

---

## Quick Reference Card

### Before Any Build

```bash
source scripts/agent-init.sh
```

### Safe Commands

```bash
# Rust (ALWAYS use cargo_agent.sh)
bash scripts/cargo_agent.sh build --release
bash scripts/cargo_agent.sh test -p formula-engine
bash scripts/cargo_agent.sh check
bash scripts/cargo_agent.sh bench --bench perf_regressions

# Node/TypeScript (with agent-init.sh sourced)
npm install
npm run build
npm test
pnpm run typecheck

# Heavy operations with explicit limit
bash scripts/run_limited.sh --as 4G -- npm run build
bash scripts/run_limited.sh --as 8G -- node heavy-script.js

# GUI tests
xvfb-run npm run test:e2e
```

### Dangerous Commands (NEVER USE)

```bash
# WRONG - can exhaust RAM:
cargo build                           # No memory limit!
cargo test                            # No memory limit!
cargo build -j$(nproc)               # 192 parallel rustc = OOM
cargo build -j0                       # Unlimited = OOM

# WRONG - kills OTHER agents' processes:
pkill cargo                           # NEVER USE pkill
killall rustc                         # NEVER USE killall
```

### Emergency: Your Process is Stuck

```bash
# Check memory status
./scripts/check-memory.sh

# Cancel via Ctrl+C (safest)

# If you must kill a specific process, find its PID first:
ps aux | grep "your-specific-pattern"
kill <specific-pid>                   # Only YOUR process
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

---

## UI/UX Design Guidelines

When implementing or modifying UI components, follow these guidelines strictly.

### Reference Materials

**ALWAYS check the mockups first:**

```bash
open mockups/spreadsheet-main.html    # Main app layout
open mockups/command-palette.html     # Cmd+K palette
open mockups/README.md                # Full design system
```

### Anti-"AI Slop" Rules

| ❌ DON'T | ✅ DO |
|----------|-------|
| Purple gradients | Electric blue accent (#3b82f6) |
| Inter, Roboto, system fonts | IBM Plex Sans/Mono |
| Rounded everything (16px radius) | Sharp corners (4-6px radius) |
| Excessive whitespace | Data-dense layouts |
| Playful/bubbly aesthetic | Professional, Bloomberg-like |
| Light gray backgrounds | Deep charcoal (#0f1114, #161a1e) |

### Color Tokens (MUST USE)

```css
/* Never hardcode colors - use these variables */
--bg-base: #0f1114;           /* App background */
--bg-surface: #161a1e;        /* Panels */
--bg-elevated: #1c2127;       /* Cards, inputs */
--accent: #3b82f6;            /* Primary actions */
--text-primary: #f1f3f5;      /* Main text */
--text-secondary: #8b939e;    /* Labels */
```

### Typography

```css
/* Sans for UI */
font-family: 'IBM Plex Sans', -apple-system, sans-serif;

/* Mono for data, code, formulas */
font-family: 'IBM Plex Mono', 'SF Mono', monospace;
```

### Component Sizes

| Component | Height | Padding |
|-----------|--------|---------|
| Input fields | 28px | 0 8px |
| Buttons | 28px | 0 12px |
| Grid cells | 24px | 0 8px |
| Panel headers | 44px | 0 16px |

### Implementation Checklist

Before submitting UI changes:

- [ ] Colors use CSS variables from design system
- [ ] Typography matches mockups (IBM Plex, not Inter)
- [ ] Spacing uses the scale (4, 8, 12, 16, 20, 24, 32px)
- [ ] No purple gradients or overly rounded corners
- [ ] Works in dark mode (light mode is future work)
- [ ] Tested at 1x and 2x pixel density
