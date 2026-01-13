/**
 * Desktop startup benchmark (Tauri binary).
 *
 * Reproducibility + safety:
 * - The desktop process is spawned with *all* user-data directories redirected under
 *   `target/perf-home` so the benchmark cannot read/write the real user profile.
 * - This avoids polluting developer machines and reduces variance on CI where cached home
 *   directories can otherwise leak across runs.
 *
 * Environment isolation is implemented in `desktopStartupUtil.ts`:
 * - All platforms: `HOME` + `USERPROFILE` => `target/perf-home`
 * - Linux: `XDG_CONFIG_HOME`, `XDG_CACHE_HOME`, `XDG_DATA_HOME` => `target/perf-home/xdg-*`
 * - Windows: `APPDATA`, `LOCALAPPDATA`, `TEMP`, `TMP` => `target/perf-home/*`
 * - macOS/Linux: `TMPDIR` => `target/perf-home/tmp`
 *
 * Startup modes:
 * - `FORMULA_DESKTOP_STARTUP_MODE=cold` (default when enabled): reset `target/perf-home` before
 *   each iteration so every launch uses a fresh app/webview profile.
 * - `FORMULA_DESKTOP_STARTUP_MODE=warm`: reset `target/perf-home` once, then reuse it so subsequent
 *   launches benefit from persisted caches (first run is treated as warmup).
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
} from './desktopStartupUtil.ts';

// Benchmark environment knobs:
// - `FORMULA_DISABLE_STARTUP_UPDATE_CHECK=1` prevents the release updater from running a
//   background check/download on startup, which can add nondeterministic CPU/memory/network
//   activity and skew startup/idle-memory benchmarks.
// - `FORMULA_STARTUP_METRICS=1` enables the Rust-side one-line startup metrics log we parse.

type StartupMode = 'cold' | 'warm';

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

  const modeRaw = (process.env.FORMULA_DESKTOP_STARTUP_MODE ?? 'cold').trim().toLowerCase();
  if (modeRaw !== 'cold' && modeRaw !== 'warm') {
    throw new Error(
      `Invalid FORMULA_DESKTOP_STARTUP_MODE=${JSON.stringify(modeRaw)} (expected "cold" or "warm")`,
    );
  }
  const mode: StartupMode = modeRaw;

  const runs = Math.max(1, Number(process.env.FORMULA_DESKTOP_STARTUP_RUNS ?? '20') || 20);
  const timeoutMs = Math.max(1, Number(process.env.FORMULA_DESKTOP_STARTUP_TIMEOUT_MS ?? '15000') || 15000);
  const binPath = process.env.FORMULA_DESKTOP_BIN
    ? resolve(process.env.FORMULA_DESKTOP_BIN)
    : defaultDesktopBinPath();

  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      `Desktop binary not found (bin=${String(binPath)}). Build it via (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build) and/or set FORMULA_DESKTOP_BIN.`,
    );
  }

  const envOverrides: NodeJS.ProcessEnv = { FORMULA_DISABLE_STARTUP_UPDATE_CHECK: '1' };

  // `desktopStartupUtil.runOnce()` can optionally reset `target/perf-home` on each
  // invocation via this parent-process env var. Make startup mode deterministic by managing
  // it here (and restoring the previous value after the benchmark completes).
  const prevResetHome = process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;
  const setResetHome = (value: string | undefined) => {
    if (value === undefined) {
      delete process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;
    } else {
      process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = value;
    }
  };

  const metrics: StartupMetrics[] = [];
  try {
    if (mode === 'warm') {
      // Start from a clean profile, then allow subsequent launches to reuse caches.
      setResetHome('1');
      // eslint-disable-next-line no-console
      console.log('[desktop-startup] warmup run 1/1 (warm)...');
      await runOnce({ binPath, timeoutMs, envOverrides });

      setResetHome(undefined);
      for (let i = 0; i < runs; i += 1) {
        // eslint-disable-next-line no-console
        console.log(`[desktop-startup] run ${i + 1}/${runs} (warm)...`);
        metrics.push(await runOnce({ binPath, timeoutMs, envOverrides }));
      }
    } else {
      // Reset before *every* run to avoid mixing cold + warm starts.
      setResetHome('1');
      for (let i = 0; i < runs; i += 1) {
        // eslint-disable-next-line no-console
        console.log(`[desktop-startup] run ${i + 1}/${runs} (cold)...`);
        metrics.push(await runOnce({ binPath, timeoutMs, envOverrides }));
      }
    }
  } finally {
    setResetHome(prevResetHome);
  }

  const windowVisible = metrics.map((m) => m.windowVisibleMs);
  const firstRender = metrics
    .map((m) => m.firstRenderMs)
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
  const tti = metrics.map((m) => m.ttiMs);
  const webviewLoaded = metrics
    .map((m) => m.webviewLoadedMs)
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v));

  const coldWindowTarget =
    Number(
      process.env.FORMULA_DESKTOP_COLD_WINDOW_VISIBLE_TARGET_MS ??
        process.env.FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS ??
        '500',
    ) || 500;
  const coldFirstRenderTarget =
    Number(
      process.env.FORMULA_DESKTOP_COLD_FIRST_RENDER_TARGET_MS ??
        process.env.FORMULA_DESKTOP_FIRST_RENDER_TARGET_MS ??
        '500',
    ) || 500;
  const coldTtiTarget =
    Number(
      process.env.FORMULA_DESKTOP_COLD_TTI_TARGET_MS ??
        process.env.FORMULA_DESKTOP_TTI_TARGET_MS ??
        '1000',
    ) || 1000;

  const warmWindowTarget =
    Number(process.env.FORMULA_DESKTOP_WARM_WINDOW_VISIBLE_TARGET_MS ?? String(coldWindowTarget)) ||
    coldWindowTarget;
  const warmFirstRenderTarget =
    Number(process.env.FORMULA_DESKTOP_WARM_FIRST_RENDER_TARGET_MS ?? String(coldFirstRenderTarget)) ||
    coldFirstRenderTarget;
  const warmTtiTarget =
    Number(process.env.FORMULA_DESKTOP_WARM_TTI_TARGET_MS ?? String(coldTtiTarget)) || coldTtiTarget;

  const windowTarget = mode === 'warm' ? warmWindowTarget : coldWindowTarget;
  const firstRenderTarget = mode === 'warm' ? warmFirstRenderTarget : coldFirstRenderTarget;
  const ttiTarget = mode === 'warm' ? warmTtiTarget : coldTtiTarget;

  if (firstRender.length !== metrics.length) {
    throw new Error(
      'Desktop did not report first_render_ms. Ensure the frontend calls `report_startup_first_render` when the grid becomes visible.',
    );
  }

  const results: BenchmarkResult[] = [
    buildResult(`desktop.startup.${mode}.window_visible_ms.p95`, windowVisible, windowTarget),
    buildResult(`desktop.startup.${mode}.first_render_ms.p95`, firstRender, firstRenderTarget),
    buildResult(`desktop.startup.${mode}.tti_ms.p95`, tti, ttiTarget),
  ];

  // Backwards compatibility: keep the legacy metric names aliased to cold-start mode.
  if (mode === 'cold') {
    results.push(buildResult('desktop.startup.window_visible_ms.p95', windowVisible, windowTarget));
    results.push(buildResult('desktop.startup.first_render_ms.p95', firstRender, firstRenderTarget));
    results.push(buildResult('desktop.startup.tti_ms.p95', tti, ttiTarget));

    const webviewLoadedTarget = Number(process.env.FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS ?? '800') || 800;

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
      results.push(buildResult('desktop.startup.webview_loaded_ms.p95', webviewLoaded, webviewLoadedTarget));
    }
  }

  return results;
}
