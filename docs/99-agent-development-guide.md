# Agent Development Environment Guide

## Overview

This guide is for **coding agents developing Formula**, not end-users. The development environment is:

- **Machine**: 192 vCPU, 1.5TB RAM, 110TB disk (Ubuntu Linux x64, headless, no GPU)
- **Concurrency**: Up to ~200 agents, each with their own repo copy
- **Critical constraint**: Memory. Disk/CPU are abundant; RAM exhaustion kills the machine.

## Local-only agent files (`scratchpad.md`, `handoff.md`)

Some agent harnesses use `scratchpad.md` (working notes) and `handoff.md` (message to the planner)
in the repo root. These files are **gitignored** and should **never** be committed.

If these files become hard to read because patch/diff formatting was pasted into them by mistake,
clean them up locally by replacing the literal backslash-`n` sequences and stray diff markers with
real newlines / normal Markdown formatting.

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
| Rust compilation    | 2-8GB         | 14GB  | Use `scripts/cargo_agent.sh`     |
| TypeScript check    | 500MB-2GB     | 2GB   | Can spike with large projects    |
| npm install         | 500MB-1GB     | 2GB   | Transient                        |
| Tests (unit)        | 200MB-1GB     | 2GB   |                                  |
| Tests (e2e/browser) | 500MB-2GB     | 2GB   | Headless Chrome                  |
| Total concurrent    | -             | 14GB  | Hard ceiling per agent workspace |


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
1. Enforces a **14GB address space limit** via `RLIMIT_AS`
2. Limits parallelism to **-j4** by default (and defaults `cargo test` to `-j1` unless you explicitly set `FORMULA_CARGO_JOBS` / `FORMULA_CARGO_TEST_JOBS`)
3. Caps **RUST_TEST_THREADS** to avoid spawning hundreds of threads (default: `min(nproc, 16, jobs * 4)`)
4. Caps **RAYON_NUM_THREADS** (defaults to `FORMULA_CARGO_JOBS`) to avoid huge per-process Rayon thread pools on high-core agent hosts
5. Defaults `MAKEFLAGS=-j<jobs>` and `CARGO_PROFILE_*_CODEGEN_UNITS=<jobs>` to keep non-Rust build steps + rustc/LLVM parallelism aligned
6. Uses a **repo-local CARGO_HOME** to avoid lock contention
7. Preserves stdout/stderr streams (does not merge stderr into stdout), so commands that emit machine-readable output on stdout (e.g. JSON) can be safely wrapped

### Rust toolchain is pinned

This repo pins Rust via `rust-toolchain.toml` so CI and desktop release builds don't drift with
new stable releases. In environments with `rustup`, `cargo` will automatically download/use the
pinned version when run from the repo root.

Note: the `RUSTUP_TOOLCHAIN` environment variable has higher precedence than `rust-toolchain.toml`.
If it is set globally (often to `stable`), it can accidentally bypass the pin. The repo cargo wrapper
(`scripts/cargo_agent.sh`) clears `RUSTUP_TOOLCHAIN` so agent builds reliably use the pinned
toolchain.

### Environment Setup (Optional but Recommended)

```bash
# Optional: override default Cargo parallelism for the session
export FORMULA_CARGO_JOBS=8

# Initialize safe defaults (sets NODE_OPTIONS, CARGO_BUILD_JOBS, MAKEFLAGS, RAYON_NUM_THREADS, etc.)
. scripts/agent-init.sh  # bash/zsh: source scripts/agent-init.sh
```

### Rust formatting (avoid noisy diffs)

For small fixes, prefer formatting the specific Rust files you touched (e.g. `rustfmt path/to/file.rs`)
instead of running `cargo fmt` across the whole workspace, which can produce large unrelated diffs.

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
- `FORMULA_CARGO_TEST_JOBS` - `cargo test` parallelism (default: 1 unless `FORMULA_CARGO_JOBS` is set)
- `FORMULA_CARGO_LIMIT_AS` - address space limit (default: 14G)
- `FORMULA_RUST_TEST_THREADS` - test parallelism (default: min(nproc, 16, jobs * 4))
- `FORMULA_RAYON_NUM_THREADS` - Rayon thread pool size (`RAYON_NUM_THREADS`) (default: `FORMULA_CARGO_JOBS`)
- `FORMULA_OPENSSL_VENDOR=1` - disable the wrapper's auto-setting of `OPENSSL_NO_VENDOR` (useful if you need vendored OpenSSL)
- `FORMULA_CARGO_RETRY_ATTEMPTS` - retry count for transient OS resource exhaustion (default: 5)
- `FORMULA_LLD_THREADS` - lld thread pool size for link steps (default: 1 on Linux host builds)

#### npm/pnpm/yarn

```bash
# GOOD: Memory-limited (NODE_OPTIONS set by agent-init.sh)
. scripts/agent-init.sh
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
- `bash scripts/cargo_agent.sh clean` (or deleting `target/`) also clears the local registry cache
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
| `pnpm test:node`       | `FORMULA_NODE_TEST_CONCURRENCY=2` (default) | Avoids over-parallelizing heavyweight integration tests |
| Likely solo operation  | `-j8`            | Avoid >8 on multi-agent hosts              |

### Node test runner concurrency

`pnpm test:node` uses Node's built-in test runner (`node --test`). On high-core
machines, Node defaults to running many test files in parallel, which can cause
timeouts in heavier integration tests.

The repo test wrapper (`scripts/run-node-tests.mjs`) keeps the suite reliable by:

- capping overall test-file concurrency by default
- running `*.e2e.test.js` files serially (`--test-concurrency=1`)
- running sync-server-backed tests (files that reference `startSyncServer`) serially
- installing a TypeScript-aware ESM loader (when `typescript` is installed) so `.test.js`
  suites can import workspace `.ts` sources without a separate build step

To tune the default non-e2e/non-sync-server test-file concurrency explicitly:

```bash
FORMULA_NODE_TEST_CONCURRENCY=4 pnpm test:node
# (or: NODE_TEST_CONCURRENCY=4)
```

To run a smaller subset of node:test suites, pass one or more substring filters
(matched against test file paths):

```bash
pnpm test:node collab
pnpm test:node startSyncServer
pnpm test:node -- desktop  # `--` delimiter is stripped for pnpm compatibility
```

`pnpm test:node --help` prints usage without running the Cursor AI policy scan (so it stays fast).

### Vitest: run single files without accidentally running the full suite

When debugging a specific Vitest file, **avoid** passing a bare `--` before the
file path (Vitest treats a literal `--` as a pattern and will end up running the
entire suite).

Prefer:

```bash
# Run a single Vitest file (no `--` delimiter)
pnpm test:vitest apps/desktop/src/tauri/__tests__/eventPermissions.vitest.ts
```

(`pnpm test:vitest -- <file>` is also OK: the wrapper at `scripts/run-vitest.mjs` strips the
`--` delimiter pnpm forwards so Vitest doesn't accidentally treat it as a test pattern.)

The wrapper runs Vitest via `vitest --run ...` (run-once mode) and tolerates common muscle memory
like `pnpm test:vitest -- run <file>` (the leading `run` is ignored). It also prefers the
calling package’s local `node_modules/.bin/vitest` so workspace packages can pin different
Vitest versions.

When running from within a package directory (e.g. `pnpm -C packages/grid test ...`), it also
normalizes repo-rooted paths like `packages/grid/src/foo.vitest.ts` to `src/foo.vitest.ts`.

Or run Vitest directly:

```bash
# Skip the WASM build in Vitest global setup (often unnecessary for unit tests)
FORMULA_SKIP_WASM_BUILD=1 node_modules/.bin/vitest run apps/desktop/src/tauri/__tests__/eventPermissions.vitest.ts
```

### Playwright: avoid a literal `--` argument when filtering e2e tests

`pnpm` forwards script arguments without requiring a `--` delimiter. If you add a bare `--`,
it will be forwarded to the underlying command as a literal argument.

For Playwright, a literal `--` is particularly problematic because it terminates option parsing,
so flags like `-g/--grep` stop working (and `--` can accidentally match additional spec file names).

Note: `apps/desktop` and `apps/web` wrap Playwright via:
- `apps/desktop/scripts/run-playwright.mjs`
- `apps/web/scripts/run-playwright.mjs`

These wrappers strip a single `--` delimiter for compatibility. Other scripts may not.
They also resolve `playwright` via the package’s local `node_modules/.bin` and run with a stable
package-root `cwd`, so `node apps/*/scripts/run-playwright.mjs ...` works reliably outside pnpm.

Prefer:

```bash
# Run specific e2e specs (no leading `--` delimiter)
pnpm -C apps/desktop test:e2e tests/e2e/freeze-panes.spec.ts tests/e2e/pivot-builder.spec.ts

# Use --grep to filter within a spec file
pnpm -C apps/desktop exec playwright test tests/e2e/split-view.spec.ts -g "Ctrl/Cmd\\+S commits"
```


### Adaptive Parallelism Script

Use this helper to pick parallelism based on current system load:

```bash
# Usage: ./scripts/smart-jobs.sh
# Returns a reasonable -j value based on current memory pressure.
# On our multi-agent hosts this intentionally caps at -j8 to avoid stampedes.

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
bash scripts/cargo_agent.sh clean
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
`packages/ai-tools`), you can skip the setup step (set `FORMULA_SKIP_WASM_BUILD=1`
or `true`):

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
  libgtk-3-dev libwebkit2gtk-4.1-dev libayatana-appindicator3-dev \
  librsvg2-dev libssl-dev patchelf

# If you need to build the Linux release bundles locally (.AppImage + .deb + .rpm),
# install the additional bundler prerequisites:
sudo apt-get install -y squashfs-tools fakeroot rpm cpio

# `appimagetool` is distributed as an AppImage and requires the FUSE 2 runtime.
sudo apt-get install -y libfuse2 || sudo apt-get install -y libfuse2t64

# Note: some distros still use `libwebkit2gtk-4.0-dev` and/or `libappindicator3-dev`.

# Build works headless (ALWAYS use the cargo wrapper in agent environments)
(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)

# Dev server needs virtual display for window
(cd apps/desktop && xvfb-run --auto-servernum bash ../../scripts/cargo_agent.sh tauri dev)
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
. scripts/agent-init.sh
bash scripts/cargo_agent.sh install sccache

# Configure (add to ~/.bashrc)
export RUSTC_WRAPPER=sccache
export SCCACHE_DIR=/shared/sccache  # Shared directory
export SCCACHE_CACHE_SIZE="50G"     # Shared cache limit
```

If `sccache` is unavailable or misconfigured (for example, `RUSTC_WRAPPER` is set globally but the
binary or cache directory isn't present), you can temporarily disable it for a single command:

```bash
RUSTC_WRAPPER= bash scripts/cargo_agent.sh test -p formula-xlsx --tests --no-run --locked
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
. scripts/agent-init.sh
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
make -j                               # Unlimited parallelism = OOM (no job limit)

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

> **Excel functionality + Cursor polish**
>
> Full Excel feature set (ribbon, sheets, formulas) with modern, clean aesthetics.
> Light mode. Professional. Polished.

### ⛔ CRITICAL: FUNCTIONALITY ≠ STYLING ⛔

**These are TWO INDEPENDENT AXES. Do not confuse them.**

```
FUNCTIONALITY (# of Excel buttons/features)
        ↑
  LOW   │   HIGH
────────┼────────→ STYLING (clean vs bloated)
        │
```

**THE RULE:**
1. **ADD all Excel buttons** - Font, Bold, Italic, Borders, Fill, Merge, AutoSum, Conditional Formatting, ALL of them
2. **STYLE them cleanly** - No Microsoft bloat, no complex nested layouts, just clean Cursor-style buttons

**FAILURE MODES (both wrong):**
- ❌ **Microsoft Bureaucracy**: HIGH functionality + BLOATED styling (complex layouts, heavy borders)
- ❌ **Stripped Minimal**: LOW functionality + CLEAN styling (removed features to "look minimal")

**CORRECT MODE:**
- ✅ **Formula**: HIGH functionality + CLEAN styling

**"Minimal" means visual simplicity, NOT feature removal.**

See `mockups/README.md` for the full design system.

### Core Philosophy

1. **Excel functionality is non-negotiable** – ribbon, sheets, formula bar, ALL buttons
2. **Light mode** – professionals prefer it
3. **Cursor-level polish** – clean, subtle, modern (but with ALL features)
4. **AI as co-pilot** – integrated but not the primary interface

### Reference Materials

**Check the mockups for direction, not pixel-perfect specs:**

```bash
mockups/spreadsheet-main.html    # Main app – ribbon, grid, AI sidebar
mockups/ai-agent-mode.html       # Agent execution view
mockups/command-palette.html     # Quick command search
mockups/README.md                # Full design system
```

> ⚠️ **Mockups are directional, not literal.**
> - Use them for **vision, layout, and design language**
> - They are **crude prototypes** — missing features, rough edges, incomplete details
> - Apply judgment: add polish, fix inconsistencies, implement missing interactions
> - Follow the **principles** (density, Excel functionality, Cursor polish) over exact pixels

### Design Tokens

**Colors (light mode):**
```css
/* Backgrounds - layered with subtle depth */
--bg-app: #f5f5f5;           /* App background */
--bg-surface: #fafafa;       /* Panels, card backgrounds */
--bg-elevated: #ffffff;      /* Active content, inputs */
--bg-hover: #eeeeee;         /* Hover states */
--bg-inset: #f0f0f0;         /* Inset panels, code blocks */

/* Text - clear hierarchy */
--text-primary: #1a1a1a;     /* Primary text */
--text-secondary: #5c5c5c;   /* Secondary text */
--text-tertiary: #8a8a8a;    /* Hints */

/* Borders */
--border: #e0e0e0;           /* Standard border */
--border-strong: #c8c8c8;    /* Emphasized border */

/* Accent - professional blue */
--accent: #0969da;           /* Primary accent */
--link: #0969da;             /* Hyperlinks */
--accent-bg: #ddf4ff;        /* Accent background */
--accent-border: #54aeff;    /* Accent border */

/* Selection */
--selection-bg: rgba(9, 105, 218, 0.08);
--selection-border: #0969da;

/* Semantic */
--green: #1a7f37;            /* Positive */
--red: #cf222e;              /* Negative */
```

**Typography:**
```css
--font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
--font-mono: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;

/* Sizes: 10px, 11px, 12px, 13px, 14px */
```

**Sizes:**
```css
--radius: 4px;      /* Panels, buttons */
--radius-sm: 3px;   /* Tags, small elements */
--radius-xs: 2px;   /* Micro radius: tiny indicators, highlights */
--radius-pill: 999px; /* Pills, segmented controls */

/* Grid: 22px rows, 90px columns, 40px row headers */
```

### Design Rules

| ✅ DO | ❌ DON'T |
|-------|----------|
| Light mode | Dark mode by default |
| Full ribbon interface | Hamburger menus |
| Sheet tabs at bottom | Hide sheet navigation |
| Visible formula bar | Collapsed inputs |
| Status bar calculations | Minimal status |
| Monospace for cells | Sans for numbers |
| System fonts | Trendy fonts |
| Professional blue accent | Purple AI gradients |

### Keyboard Shortcuts

**Excel-compatible (MUST work):**
- `F2` – Edit cell
- `F4` – Toggle absolute/relative
- `Ctrl+D` – Fill down
- `Ctrl+;` – Insert date
- `Alt+=` – AutoSum
- All standard Ctrl+C/V/X/Z/Y

**App shortcuts:**
- `Cmd/Ctrl+Shift+P` – Command palette

**AI shortcuts:**
- `Cmd/Ctrl+K` – Inline AI edit
- `Cmd+I` (macOS) / `Ctrl+Shift+A` (Windows/Linux) – Toggle AI sidebar
- `Tab` – Accept suggestion
  
Platform note:
- **macOS:** `Cmd+I` is reserved for **AI Chat**. Use `Ctrl+I` for **Italic** (Excel-compatible).
- **Windows/Linux:** `Ctrl+I` is reserved for **Italic** (Excel-compatible). Use `Ctrl+Shift+A` to toggle the AI sidebar.

### AI Sidebar

**One unified panel** - no mode tabs. Just type what you want:
- Ask questions → AI answers
- Request changes → AI proposes diff
- Complex tasks → opens Agent view

Features:
- Context tags showing what AI sees
- Diff preview before apply
- Accept/reject per change
- Agent view for autonomous multi-step execution (full-height separate view)

### Implementation Checklist

- [ ] Full ribbon interface with tabs/groups
- [ ] Light mode colors
- [ ] Sheet tabs at bottom
- [ ] Formula bar always visible
- [ ] Status bar with Sum/Avg/Count
- [ ] Excel keyboard shortcuts work
- [ ] Monospace font for cells
- [ ] Tested with financial data
