/**
 * Desktop startup benchmark (Tauri binary).
 *
 * Reproducibility + safety:
 * - The desktop process is spawned with *all* user-data directories redirected under
 *   `target/perf-home` so the benchmark cannot read/write the real user profile.
 * - This avoids polluting developer machines and reduces variance on CI where cached home
 *   directories can otherwise leak across runs.
 *
 * Environment isolation is implemented in `desktopStartupRunnerShared.ts`:
 * - All platforms: `HOME` + `USERPROFILE` => `target/perf-home`
 * - Linux: `XDG_CONFIG_HOME`, `XDG_CACHE_HOME`, `XDG_DATA_HOME` => `target/perf-home/xdg-*`
 * - Windows: `APPDATA`, `LOCALAPPDATA`, `TEMP`, `TMP` => `target/perf-home/*`
 * - macOS/Linux: `TMPDIR` => `target/perf-home/tmp`
 *
 * Reset behavior:
 * - Set `FORMULA_DESKTOP_BENCH_RESET_HOME=1` to delete `target/perf-home` before *each* iteration.
 */

import { existsSync } from 'node:fs';
import { resolve } from 'node:path';

import { type BenchmarkResult } from './benchmark.ts';
import {
  defaultDesktopBinPath,
  mean,
  median,
  percentile,
  runOnce,
  stdDev,
  type StartupMetrics,
} from './desktopStartupRunnerShared.ts';

// Benchmark environment knobs:
// - `FORMULA_DISABLE_STARTUP_UPDATE_CHECK=1` prevents the release updater from running a
//   background check/download on startup, which can add nondeterministic CPU/memory/network
//   activity and skew startup/idle-memory benchmarks.
// - `FORMULA_STARTUP_METRICS=1` enables the Rust-side one-line startup metrics log we parse.

function buildResult(name: string, values: number[], targetMs: number): BenchmarkResult {
  const sorted = [...values].sort((a, b) => a - b);
  const avg = mean(sorted);
  const med = median(sorted);
  const p95 = percentile(sorted, 0.95);
  const p99 = percentile(sorted, 0.99);
  const sd = stdDev(sorted, avg);

  return {
    name,
    iterations: values.length,
    warmup: 0,
    unit: 'ms',
    mean: avg,
    median: med,
    p95,
    p99,
    stdDev: sd,
    targetMs,
    passed: p95 <= targetMs,
  };
}
export async function runDesktopStartupBenchmarks(): Promise<BenchmarkResult[]> {
  if (process.env.FORMULA_RUN_DESKTOP_STARTUP_BENCH !== '1') {
    return [];
  }

  const runs = Math.max(1, Number(process.env.FORMULA_DESKTOP_STARTUP_RUNS ?? '20') || 20);
  const timeoutMs = Math.max(1, Number(process.env.FORMULA_DESKTOP_STARTUP_TIMEOUT_MS ?? '15000') || 15000);
  const binPath = process.env.FORMULA_DESKTOP_BIN ? resolve(process.env.FORMULA_DESKTOP_BIN) : defaultDesktopBinPath();

  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      `Desktop binary not found (bin=${String(binPath)}). Build it via (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build) and/or set FORMULA_DESKTOP_BIN.`,
    );
  }

  const metrics: StartupMetrics[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-startup] run ${i + 1}/${runs}...`);
    metrics.push(
      await runOnce({
        binPath,
        timeoutMs,
        envOverrides: { FORMULA_DISABLE_STARTUP_UPDATE_CHECK: '1' },
      }),
    );
  }

  const windowVisible = metrics.map((m) => m.windowVisibleMs);
  const tti = metrics.map((m) => m.ttiMs);
  const webviewLoaded = metrics
    .map((m) => m.webviewLoadedMs)
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v));

  const windowTarget = Number(process.env.FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS ?? '500') || 500;
  const ttiTarget = Number(process.env.FORMULA_DESKTOP_TTI_TARGET_MS ?? '1000') || 1000;
  const webviewLoadedTarget = Number(process.env.FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS ?? '800') || 800;

  const results: BenchmarkResult[] = [
    buildResult('desktop.startup.window_visible_ms.p95', windowVisible, windowTarget),
    buildResult('desktop.startup.tti_ms.p95', tti, ttiTarget),
  ];

  // `webview_loaded_ms` is best-effort and historically missing in some runs due to the
  // frontend->Rust IPC call racing `__TAURI__` initialization. To avoid failing the whole
  // benchmark suite while the instrumentation is still stabilizing, only report the metric
  // when we get a sufficiently large sample.
  //
  // Policy:
  // - If 0 runs report `webview_loaded_ms`, skip the metric entirely.
  // - If fewer than 80% of runs report it, skip the metric (avoid biased p95 on a tiny subset).
  // - Otherwise, compute p95 over the runs that reported a value and gate on the target.
  const minValidFraction = 0.8;
  const minValidRuns = Math.ceil(runs * minValidFraction);
  if (webviewLoaded.length === 0) {
    // eslint-disable-next-line no-console
    console.log('[desktop-startup] webview_loaded_ms unavailable (0 runs reported it); skipping metric');
  } else if (webviewLoaded.length < minValidRuns) {
    // eslint-disable-next-line no-console
    console.log(
      `[desktop-startup] webview_loaded_ms only available for ${webviewLoaded.length}/${runs} runs (<${Math.round(
        minValidFraction * 100,
      )}%); skipping metric`,
    );
  } else {
    results.push(
      buildResult('desktop.startup.webview_loaded_ms.p95', webviewLoaded, webviewLoadedTarget),
    );
  }

  return results;
}
