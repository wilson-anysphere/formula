# Performance Targets & Optimization Strategy

## Overview

Performance is a feature. Users should never wait, never see jank, never hit limits. These targets represent the bar we must clear to win over power users.

---

## Performance Targets

### Startup

| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| Cold start to interactive | <1.0s | Time from launch to first input accepted |
| Warm start | <0.5s | Time from launch with cached data |
| Time to first render | <0.3s | Time from launch to grid visible |
| Frontend asset download size (compressed JS/CSS/WASM) | <10MB | Brotli-compressed total of `dist/assets/**/*.{js,css,wasm}` (default; gzip optional via `FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION=gzip`; see `node scripts/frontend_asset_size_report.mjs`) |
| Desktop installer artifact size (DMG/MSI/EXE/AppImage) | <50MB per artifact | `python scripts/desktop_bundle_size_report.py` on Tauri build output (`target/**/release/bundle`) |

### File Operations

| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| Open 1MB xlsx | <1s | File dialog close to grid rendered |
| Open 10MB xlsx | <3s | File dialog close to grid rendered |
| Open 100MB xlsx | <10s | File dialog close to grid rendered |
| Save 1MB xlsx | <0.5s | Save command to file written |
| Save 10MB xlsx | <2s | Save command to file written |
| Save 100MB xlsx | <5s | Save command to file written |

### Rendering

| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| Scroll FPS | 60fps | Chrome DevTools Performance |
| Scroll with 1M rows | 60fps | No degradation |
| Scroll with 100 columns | 60fps | No degradation |
| Selection render | <16ms | Time from click to highlight |
| Cell edit start | <50ms | Time from F2 to cursor |
| Cell edit commit | <50ms | Time from Enter to display |

### Calculation

| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| Simple formula | <1ms | Single SUM of 100 cells |
| 1K cell recalc | <10ms | Dependent chain of 1K cells |
| 10K cell recalc | <50ms | Dependent chain of 10K cells |
| 100K cell recalc | <100ms | Dependent chain of 100K cells |
| 1M cell recalc | <1s | Full recalc of 1M formula cells |
| VLOOKUP 10K rows | <10ms | Single lookup in 10K row table |
| VLOOKUP 100K rows | <50ms | Single lookup in 100K row table |

### Memory

| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| Empty workbook | <50MB | Heap snapshot |
| 1MB xlsx loaded | <100MB | Heap snapshot |
| 10MB xlsx loaded | <200MB | Heap snapshot |
| 100MB xlsx loaded | <500MB | Heap snapshot |
| Memory per cell | <100 bytes | Calculated from heap |
| Desktop idle RSS (empty workbook) | <100MB | CI benchmark `desktop.memory.idle_rss_mb.p95` (process RSS + child processes) |

### Collaboration

| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| Edit to sync | <100ms | Time for change to appear on other client |
| Presence update | <200ms | Time for cursor to update |
| Conflict resolution | <500ms | Time to merge concurrent edits |
| Offline duration | Unlimited | Time can work offline |

### AI
 
| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| Tab completion | <100ms | Time from keystroke to suggestion |
| Inline assist | <2s | Time from Cmd/Ctrl+K to result |
| Chat response | <5s | Time from Enter to first token |
| Formula explanation | <3s | Time from request to explanation |

#### Tab-completion latency guardrail

We enforce the tab-completion latency budget with a lightweight micro-benchmark for the JS
`TabCompletionEngine`:

```bash
pnpm bench:tab-completion
# or:
node packages/ai-completion/bench/tabCompletionEngine.bench.mjs --ci
```

This benchmark prints p50/p95 per scenario and exits non-zero in CI if p95 exceeds the budget
(default: 100ms). It is also run in CI as `tab-completion-latency-guard`.

---

## Benchmarking Infrastructure

### Automated Performance Tests

The repo's CI performance suite is run via:

```bash
pnpm benchmark
```

This executes `apps/desktop/tests/performance/run.ts`, which contains both JS/TS microbenchmarks and
Rust engine benchmarks (`formula-engine/src/bin/perf_bench.rs`).

### Desktop perf commands (startup, memory, size)

For contributor-friendly **desktop shell** measurements (Tauri binary + real WebView), use the
following commands from the repo root:

```bash
# Builds the desktop frontend + binary, then measures cold-start timings (p50/p95).
pnpm perf:desktop-startup

# Builds the desktop frontend + binary, then measures idle memory (RSS) after TTI.
pnpm perf:desktop-memory

# Reports size of apps/desktop/dist, the desktop binary, and (if present) Tauri bundle artifacts.
pnpm perf:desktop-size
```

To capture machine-readable output, forward `--json` args to the underlying runner:

```bash
pnpm perf:desktop-startup -- --json target/perf-artifacts/desktop-startup.json
pnpm perf:desktop-memory -- --json target/perf-artifacts/desktop-memory.json
```

These scripts are designed to be safe to run locally:

- they use a repo-local HOME directory (`target/perf-home`) so they don't touch your real user config/caches
- you can override it with `FORMULA_PERF_HOME=/path/to/dir`
- set `FORMULA_PERF_PRESERVE_HOME=1` to avoid clearing the perf HOME between runs
- safety: `pnpm perf:desktop-*` will only auto-delete perf homes under `target/` by default; set
  `FORMULA_PERF_ALLOW_UNSAFE_CLEAN=1` to allow clearing a custom path outside `target/`

#### What the metrics mean

- **`window_visible_ms`**: time from native process start until the main window is shown.
- **`webview_loaded_ms`**: time from native process start until the WebView reports the initial page load finished
  (native callback; useful for separating “shell/webview” cost from frontend JS work).
- **`first_render_ms`**: time from native process start until the grid becomes visible (first meaningful render).
- **`tti_ms`**: time-to-interactive (first input accepted); includes webview + frontend init.
- **Startup benchmark kinds**:
  - **Full app** (`desktop.startup.*`): end-to-end UI startup (requires `apps/desktop/dist`); includes `first_render_ms`.
  - **Shell-only** (`desktop.shell_startup.*`): minimal webview startup via `--startup-bench` (does **not** require `apps/desktop/dist`);
    does not emit `first_render_ms`.
- **Startup benchmark modes**:
  - **Cold** (`.cold.`): each run uses a fresh profile directory (true cold start), e.g.
    - `desktop.startup.cold.window_visible_ms.p95`
    - `desktop.startup.cold.tti_ms.p95`
  - **Warm** (`.warm.`): runs reuse the same profile directory after a warmup launch, e.g.
    - `desktop.startup.warm.window_visible_ms.p95`
    - `desktop.startup.warm.tti_ms.p95`
  - Legacy unscoped metric names (e.g. `desktop.startup.window_visible_ms.p95`) are kept as **aliases to cold** mode for backwards compatibility.
- **`idleRssMb`**: resident set size (RSS) of the desktop process *plus child processes*, sampled
  after TTI + a "settle" delay. (Useful for regression tracking; RSS can double-count shared pages.)
- **Size**:
  - `apps/desktop/dist` is the Vite-built frontend asset directory embedded/served by Tauri.
  - `target/**/formula-desktop` is the built desktop executable.
  - `pnpm perf:desktop-size` also runs `python3 scripts/desktop_binary_size_report.py` (cargo-bloat + llvm-size fallback) to show which Rust crates/symbols dominate the desktop binary size.
  - `target/**/release/bundle` contains installer artifacts when you run `cargo tauri build`.

#### CI gating / overrides

Startup benchmark (runner defaults shown):

- `FORMULA_DESKTOP_STARTUP_MODE=cold|warm` (default: `cold`)
  - `cold`: each measured run uses a fresh isolated profile directory under `FORMULA_PERF_HOME` (default: `target/perf-home`),
    so caches do not carry across iterations (true cold-start).
  - `warm`: uses a single profile directory; one warmup run initializes caches, then measured runs reuse that profile.
- `FORMULA_DESKTOP_STARTUP_BENCH_KIND=shell|full` (default: `full` locally, `shell` on CI)
  - `shell`: runs the desktop binary with `--startup-bench` (does not require `apps/desktop/dist`)
  - `full`: runs the full app (requires `apps/desktop/dist`)
- `FORMULA_DESKTOP_STARTUP_RUNS=20`
- `FORMULA_DESKTOP_STARTUP_TIMEOUT_MS=15000`
- Cold-start targets (defaults shown; legacy unscoped vars are still accepted as fallbacks):
  - `FORMULA_DESKTOP_COLD_WINDOW_VISIBLE_TARGET_MS=500`
  - `FORMULA_DESKTOP_COLD_FIRST_RENDER_TARGET_MS=500`
  - `FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS=800` (p95 budget for `webview_loaded_ms`)
  - `FORMULA_DESKTOP_COLD_TTI_TARGET_MS=1000`
- Warm-start targets (optional; default to the cold targets if unset):
  - `FORMULA_DESKTOP_WARM_WINDOW_VISIBLE_TARGET_MS`
  - `FORMULA_DESKTOP_WARM_FIRST_RENDER_TARGET_MS`
  - `FORMULA_DESKTOP_WARM_TTI_TARGET_MS`
- `FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=1` to fail when p95 exceeds targets
- `FORMULA_RUN_DESKTOP_STARTUP_BENCH=1` to allow running the runner in CI (it skips by default)

Memory benchmark:

- `FORMULA_DESKTOP_MEMORY_RUNS=10`
- `FORMULA_DESKTOP_MEMORY_SETTLE_MS=5000`
- `FORMULA_DESKTOP_MEMORY_TIMEOUT_MS=20000`
- `FORMULA_DESKTOP_IDLE_RSS_TARGET_MB=<budget>` (default: 100)
- `FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=1` to fail when p95 exceeds the budget
- `FORMULA_RUN_DESKTOP_MEMORY_BENCH=1` to allow running the runner in CI (it skips by default)

Desktop installer artifact size gating (used by the desktop release workflow):

- `FORMULA_BUNDLE_SIZE_LIMIT_MB=50` (default: 50MB per artifact)
- `FORMULA_ENFORCE_BUNDLE_SIZE=1` to fail when any artifact exceeds the limit
  (reported by `scripts/desktop_bundle_size_report.py`)

Also reported (installer artifacts) on Linux PRs/main (informational by default) via `.github/workflows/desktop-bundle-size.yml`
(workflow name: “Desktop installer artifact sizes”)
which builds the Linux desktop bundles and uploads a JSON size report artifact for debugging.

Lightweight PR size gating (desktop binary + `apps/desktop/dist`; disabled by default):

- `FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB=<budget>`
- `FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB=<budget>`
  (enforced by `scripts/desktop_size_report.py` when set; CI passes these via GitHub Actions Variables)

Rust desktop binary size breakdown (cargo-bloat; informational by default):

- `FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB=<budget>`
- `FORMULA_ENFORCE_DESKTOP_BINARY_SIZE=1` (or `true`/`yes`/`on`) to make the size breakdown step fail when the binary exceeds the budget
  (reported by `scripts/desktop_binary_size_report.py`)

Desktop binary/dist size budgets (optional; used by `pnpm benchmark` and `pnpm perf:desktop-size`):

- `FORMULA_DESKTOP_BINARY_SIZE_TARGET_MB=<budget>` (decimal MB)
- `FORMULA_DESKTOP_DIST_SIZE_TARGET_MB=<budget>` (decimal MB)
- `FORMULA_DESKTOP_DIST_GZIP_SIZE_TARGET_MB=<budget>` (decimal MB; computed via `tar -czf`)
Frontend asset download size gating (web/desktop Vite `dist/assets`):

- `FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB=10` (default: 10MB total)
- `FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION=brotli|gzip` (default: brotli)
- `FORMULA_ENFORCE_FRONTEND_ASSET_SIZE=1` to fail when the total exceeds the limit
  (reported by `scripts/frontend_asset_size_report.mjs`)

CI wiring note: `.github/workflows/ci.yml` runs this report for both `apps/web/dist` and
`apps/desktop/dist`, and passes these env vars via GitHub Actions **Variables** (unset by default).

#### Renderer guardrails (Node/JSDOM)

For rendering, we also run the real `@formula/grid` canvas renderer under Node (via JSDOM + a mocked
2D canvas context) to guard against regressions in per-frame CPU cost at large sheet sizes:

- `gridRenderer.firstFrame.p95` (attach + resize + first frame)
- `gridRenderer.scrollStep.p95` (scroll + frame, 1M+ rows)
- `gridRenderer.scrollStepHorizontal.p95` (horizontal scroll + frame)
- `gridRenderer.selectionChange.p95` (selection update + frame)

These benchmarks enforce a 16ms p95 budget (60fps target). They are timed using CPU time
(`process.cpuUsage()`) rather than wall time to reduce noise from OS scheduling on shared CI runners.

```typescript
// tests/performance/benchmark.ts
 
interface BenchmarkResult {
  name: string;
  iterations: number;
  mean: number;
  median: number;
  p95: number;
  p99: number;
  stdDev: number;
  passed: boolean;
  target: number;
}

async function runBenchmark(
  name: string,
  fn: () => Promise<void>,
  options: { iterations?: number; warmup?: number; target: number }
): Promise<BenchmarkResult> {
  const { iterations = 100, warmup = 10, target } = options;
  const results: number[] = [];
  
  // Warmup
  for (let i = 0; i < warmup; i++) {
    await fn();
  }
  
  // Measure
  for (let i = 0; i < iterations; i++) {
    const start = performance.now();
    await fn();
    results.push(performance.now() - start);
  }
  
  // Calculate statistics
  results.sort((a, b) => a - b);
  const mean = results.reduce((a, b) => a + b) / results.length;
  const median = results[Math.floor(results.length / 2)];
  const p95 = results[Math.floor(results.length * 0.95)];
  const p99 = results[Math.floor(results.length * 0.99)];
  const variance = results.reduce((sum, x) => sum + Math.pow(x - mean, 2), 0) / results.length;
  const stdDev = Math.sqrt(variance);
  
  return {
    name,
    iterations,
    mean,
    median,
    p95,
    p99,
    stdDev,
    passed: p95 <= target,
    target
  };
}
```

### Continuous Performance Monitoring

```yaml
# .github/workflows/perf.yml
name: Performance

on:
  push:
    branches: [main]
  pull_request:

jobs:
  benchmark:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 22
      - run: npm ci
      - run: npm run benchmark
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          tool: 'customSmallerIsBetter'
          output-file-path: benchmark-results.json
          github-token: ${{ secrets.GITHUB_TOKEN }}
          auto-push: true
          alert-threshold: '120%'
          comment-on-alert: true
          fail-on-alert: true
```

---

## Optimization Strategies

### Startup Optimization

```typescript
// 1. Code splitting - only load what's needed immediately
const CalcEngine = lazy(() => import('./calc-engine'));
const AIAssistant = lazy(() => import('./ai-assistant'));
const ChartRenderer = lazy(() => import('./chart-renderer'));

// 2. Pre-computed data
// Compile function signatures, help text at build time
import { FUNCTION_SIGNATURES } from './generated/function-signatures';

// 3. Service worker for instant load
// Cache shell HTML, core JS, CSS
self.addEventListener('fetch', (event) => {
  if (event.request.mode === 'navigate') {
    event.respondWith(caches.match('/shell.html'));
  }
});

// 4. Skeleton UI
// Show grid structure immediately, fill data async
function showSkeletonUI() {
  renderGridLines();
  renderHeaders();
  // Data comes later via streaming
}
```

### Calculation Optimization

```rust
// 1. SIMD for bulk operations
#[cfg(target_arch = "x86_64")]
use std::arch::x86_64::*;

fn sum_f64_simd(values: &[f64]) -> f64 {
    unsafe {
        let mut sum = _mm256_setzero_pd();
        
        for chunk in values.chunks_exact(4) {
            let v = _mm256_loadu_pd(chunk.as_ptr());
            sum = _mm256_add_pd(sum, v);
        }
        
        // Horizontal sum
        let low = _mm256_castpd256_pd128(sum);
        let high = _mm256_extractf128_pd(sum, 1);
        let sum128 = _mm_add_pd(low, high);
        let high64 = _mm_unpackhi_pd(sum128, sum128);
        _mm_cvtsd_f64(_mm_add_sd(sum128, high64))
    }
}

// 2. Parallel calculation with rayon
fn recalculate_parallel(cells: &mut [Cell], graph: &DependencyGraph) {
    let order = graph.topological_sort();
    let levels = graph.parallelize(order);
    
    for level in levels {
        level.par_iter().for_each(|cell_id| {
            // Safe because cells in same level are independent
            evaluate_cell(cell_id);
        });
    }
}

// 3. Incremental calculation
fn recalculate_incremental(dirty: &HashSet<CellId>) {
    let affected = graph.get_all_dependents(dirty);
    let order = graph.topological_sort_subset(&affected);
    
    for cell_id in order {
        evaluate_cell(cell_id);
    }
}
```

### Rendering Optimization

```typescript
// 1. Request idle callback for non-critical updates
function updateCellCache(cells: Cell[]) {
  const deadline = 16; // ms
  let index = 0;
  
  function processChunk(idleDeadline: IdleDeadline) {
    while (index < cells.length && idleDeadline.timeRemaining() > 0) {
      cacheCell(cells[index]);
      index++;
    }
    
    if (index < cells.length) {
      requestIdleCallback(processChunk);
    }
  }
  
  requestIdleCallback(processChunk);
}

// 2. Offscreen canvas for complex rendering
const offscreen = new OffscreenCanvas(width, height);
const offscreenCtx = offscreen.getContext('2d');

function renderComplex() {
  // Render to offscreen
  renderToCanvas(offscreenCtx);
  
  // Copy to main in single operation
  mainCtx.drawImage(offscreen, 0, 0);
}

// 3. GPU acceleration for transforms
canvas.style.willChange = 'transform';
canvas.style.transform = `translate3d(${x}px, ${y}px, 0)`;

// 4. Efficient text rendering
const textCache = new Map<string, ImageBitmap>();

function renderText(text: string, style: TextStyle): ImageBitmap {
  const key = `${text}|${styleKey(style)}`;
  
  if (!textCache.has(key)) {
    const canvas = new OffscreenCanvas(100, 20);
    const ctx = canvas.getContext('2d');
    ctx.font = fontString(style);
    ctx.fillText(text, 0, 16);
    textCache.set(key, canvas.transferToImageBitmap());
  }
  
  return textCache.get(key);
}
```

### Memory Optimization

```typescript
// 1. Sparse storage - only store non-empty cells
class SparseSheet {
  private cells = new Map<number, Cell>();
  
  private key(row: number, col: number): number {
    return (row << 16) | col;
  }
  
  get(row: number, col: number): Cell | undefined {
    return this.cells.get(this.key(row, col));
  }
  
  set(row: number, col: number, cell: Cell): void {
    if (cell.isEmpty()) {
      this.cells.delete(this.key(row, col));
    } else {
      this.cells.set(this.key(row, col), cell);
    }
  }
}

// 2. String deduplication
class StringPool {
  private pool = new Map<string, string>();
  
  intern(str: string): string {
    if (!this.pool.has(str)) {
      this.pool.set(str, str);
    }
    return this.pool.get(str)!;
  }
}

// 3. Lazy parsing - don't parse formulas until needed
class LazyCell {
  private _ast?: ASTNode;
  private _formula?: string;
  
  get ast(): ASTNode {
    if (!this._ast && this._formula) {
      this._ast = parse(this._formula);
    }
    return this._ast!;
  }
}

// 4. LRU cache for computed values
class LRUCache<K, V> {
  private cache = new Map<K, V>();
  private readonly maxSize: number;
  
  get(key: K): V | undefined {
    const value = this.cache.get(key);
    if (value !== undefined) {
      // Move to end (most recent)
      this.cache.delete(key);
      this.cache.set(key, value);
    }
    return value;
  }
  
  set(key: K, value: V): void {
    if (this.cache.size >= this.maxSize) {
      // Delete oldest (first)
      const firstKey = this.cache.keys().next().value;
      this.cache.delete(firstKey);
    }
    this.cache.set(key, value);
  }
}
```

---

## Profiling Tools

### Development Profiling

```typescript
// Custom performance markers
performance.mark('recalc-start');
engine.recalculate();
performance.mark('recalc-end');
performance.measure('recalculation', 'recalc-start', 'recalc-end');

// Memory profiling
console.log('Memory before:', process.memoryUsage());
// ... operation
console.log('Memory after:', process.memoryUsage());

// Flame graph generation
const profiler = require('v8-profiler-next');
profiler.startProfiling('MyProfile');
// ... operation
const profile = profiler.stopProfiling('MyProfile');
profile.export(function(error, result) {
  fs.writeFileSync('profile.cpuprofile', result);
  profile.delete();
});
```

### Production Monitoring

```typescript
// Real User Monitoring (RUM)
class PerformanceMonitor {
  private metrics: Map<string, number[]> = new Map();
  
  measure(name: string, fn: () => void): void {
    const start = performance.now();
    fn();
    const duration = performance.now() - start;
    
    if (!this.metrics.has(name)) {
      this.metrics.set(name, []);
    }
    this.metrics.get(name)!.push(duration);
    
    // Send to analytics
    if (this.shouldReport()) {
      this.report();
    }
  }
  
  private report(): void {
    const summary: Record<string, { p50: number; p95: number; p99: number }> = {};
    
    for (const [name, values] of this.metrics) {
      values.sort((a, b) => a - b);
      summary[name] = {
        p50: values[Math.floor(values.length * 0.5)],
        p95: values[Math.floor(values.length * 0.95)],
        p99: values[Math.floor(values.length * 0.99)]
      };
    }
    
    analytics.track('performance', summary);
    this.metrics.clear();
  }
}
```

---

## Performance Regression Prevention

### Pre-commit Hooks

```bash
#!/bin/bash
# .git/hooks/pre-commit

echo "Running performance tests..."
npm run perf:quick

if [ $? -ne 0 ]; then
  echo "Performance regression detected. Commit blocked."
  exit 1
fi
```

### PR Performance Checks

```typescript
// Compare PR performance against main
async function comparePerformance(): Promise<PerfComparison> {
  const mainResults = await getMainBranchResults();
  const prResults = await runBenchmarks();
  
  const regressions: string[] = [];
  
  for (const [name, prMetric] of Object.entries(prResults)) {
    const mainMetric = mainResults[name];
    if (!mainMetric) continue;
    
    const change = (prMetric.p95 - mainMetric.p95) / mainMetric.p95;
    
    if (change > 0.1) { // 10% regression threshold
      regressions.push(`${name}: ${(change * 100).toFixed(1)}% slower`);
    }
  }
  
  return {
    passed: regressions.length === 0,
    regressions,
    summary: generateSummary(mainResults, prResults)
  };
}
```

---

## Target Hardware

### Minimum Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| CPU | 2 cores, 2GHz | 4 cores, 3GHz |
| RAM | 4GB | 8GB |
| Storage | 500MB free | 1GB free |
| Display | 1280x720 | 1920x1080 |
| GPU | Integrated | Dedicated |

### Performance Scaling

| Workload | Min Hardware | Rec Hardware |
|----------|--------------|--------------|
| 10K cells | ✓ Fast | ✓ Instant |
| 100K cells | ✓ Acceptable | ✓ Fast |
| 1M cells | ⚠️ Slow | ✓ Acceptable |
| 10M cells | ❌ Not recommended | ⚠️ Slow |

---

## Performance Budget

### Network Budget

| Resource | Budget | Notes |
|----------|--------|-------|
| Initial HTML | 15KB | Shell only |
| Critical CSS | 10KB | Above-fold styles |
| Critical JS | 100KB | Core functionality |
| Total initial | 200KB | Under 3G target |
| Full app | 5MB | After lazy loading |

Note: the repo’s CI-enforced **compressed JS/CSS/WASM** budget is tracked separately as **Frontend asset download size** (<10MB brotli by default; see `node scripts/frontend_asset_size_report.mjs`).

### Runtime Budget

| Operation | Budget | Notes |
|-----------|--------|-------|
| Frame render | 16ms | 60fps target |
| Input response | 50ms | Feels instant |
| Animation | 100ms | Smooth transition |
| Content load | 1000ms | User waits |
| Heavy operation | 10000ms | Show progress |
