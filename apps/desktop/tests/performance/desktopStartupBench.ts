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
 * Startup modes (profile reset behavior):
 * - `FORMULA_DESKTOP_STARTUP_MODE=cold` (default when enabled): each iteration uses a fresh
 *   profile directory under `target/perf-home` so every launch is a true cold start.
 * - `FORMULA_DESKTOP_STARTUP_MODE=warm`: a single profile directory is initialized once (warmup),
 *   then reused for the measured runs so persisted caches are reflected in the results.
 *
 * Benchmark kind (what we measure):
 * - `FORMULA_DESKTOP_STARTUP_BENCH_KIND=full` (default): launch the normal app (requires bundled frontend assets).
 * - `FORMULA_DESKTOP_STARTUP_BENCH_KIND=shell`: launch `--startup-bench` (measures shell/webview startup without
 *   requiring `apps/desktop/dist`).
 */
import { existsSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

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
type StartupBenchKind = 'full' | 'shell';

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

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

function parseStartupMode(): StartupMode {
  const modeRaw = (process.env.FORMULA_DESKTOP_STARTUP_MODE ?? 'cold').trim().toLowerCase();
  if (modeRaw !== 'cold' && modeRaw !== 'warm') {
    throw new Error(
      `Invalid FORMULA_DESKTOP_STARTUP_MODE=${JSON.stringify(modeRaw)} (expected "cold" or "warm")`,
    );
  }
  return modeRaw;
}

function parseBenchKind(): StartupBenchKind {
  const kindRaw = (process.env.FORMULA_DESKTOP_STARTUP_BENCH_KIND ?? '').trim().toLowerCase();
  if (!kindRaw) {
    // Prefer the lightweight shell benchmark in CI so we can measure startup without requiring
    // a Vite/wasm build of `apps/desktop/dist`. For full end-to-end measurements, set
    // `FORMULA_DESKTOP_STARTUP_BENCH_KIND=full` (or run the dedicated desktop perf workflows).
    return process.env.CI ? 'shell' : 'full';
  }
  if (kindRaw !== 'full' && kindRaw !== 'shell') {
    throw new Error(
      `Invalid FORMULA_DESKTOP_STARTUP_BENCH_KIND=${JSON.stringify(kindRaw)} (expected "full" or "shell")`,
    );
  }
  return kindRaw;
}

export async function runDesktopStartupBenchmarks(): Promise<BenchmarkResult[]> {
  if (process.env.FORMULA_RUN_DESKTOP_STARTUP_BENCH !== '1') {
    return [];
  }

  const startupMode = parseStartupMode();
  const benchKind = parseBenchKind();

  const runs = Math.max(1, Number(process.env.FORMULA_DESKTOP_STARTUP_RUNS ?? '20') || 20);
  const timeoutMs = Math.max(
    1,
    Number(process.env.FORMULA_DESKTOP_STARTUP_TIMEOUT_MS ?? '15000') || 15000,
  );
  const binPath = process.env.FORMULA_DESKTOP_BIN
    ? resolve(process.env.FORMULA_DESKTOP_BIN)
    : defaultDesktopBinPath();

  if (!binPath || !existsSync(binPath)) {
    const buildHint =
      benchKind === 'shell'
        ? 'cargo build -p formula-desktop-tauri --bin formula-desktop --features desktop --release'
        : '(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)';
    throw new Error(
      `Desktop binary not found (bin=${String(binPath)}). Build it via ${buildHint} and/or set FORMULA_DESKTOP_BIN.`,
    );
  }

  const envOverrides: NodeJS.ProcessEnv = { FORMULA_DISABLE_STARTUP_UPDATE_CHECK: '1' };
  const argv = benchKind === 'shell' ? ['--startup-bench'] : [];

  const perfHome =
    process.env.FORMULA_PERF_HOME && process.env.FORMULA_PERF_HOME.trim() !== ''
      ? resolve(repoRoot, process.env.FORMULA_PERF_HOME)
      : resolve(repoRoot, 'target', 'perf-home');

  const profileRoot = resolve(
    perfHome,
    `desktop-startup-${benchKind}-${startupMode}-${Date.now()}-${process.pid}`,
  );

  // `desktopStartupUtil.runOnce()` can optionally reset the profile directory (HOME/XDG/etc) on
  // each invocation via this parent-process env var. Make startup mode deterministic by managing
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
    if (startupMode === 'warm') {
      const profileDir = resolve(profileRoot, 'profile');
      // Start from a clean profile, then allow subsequent launches to reuse caches.
      setResetHome('1');
      // eslint-disable-next-line no-console
      console.log(`[desktop-${benchKind}-startup] warmup run 1/1 (warm, profile=${profileDir})...`);
      await runOnce({ binPath, timeoutMs, argv, envOverrides, profileDir });

      setResetHome(undefined);
      for (let i = 0; i < runs; i += 1) {
        // eslint-disable-next-line no-console
        console.log(`[desktop-${benchKind}-startup] run ${i + 1}/${runs} (warm, profile=${profileDir})...`);
        metrics.push(await runOnce({ binPath, timeoutMs, argv, envOverrides, profileDir }));
      }
    } else {
      // Reset before *every* run to avoid mixing cold + warm starts.
      setResetHome('1');
      for (let i = 0; i < runs; i += 1) {
        const profileDir = resolve(profileRoot, `run-${String(i + 1).padStart(2, '0')}`);
        // eslint-disable-next-line no-console
        console.log(`[desktop-${benchKind}-startup] run ${i + 1}/${runs} (cold, profile=${profileDir})...`);
        metrics.push(await runOnce({ binPath, timeoutMs, argv, envOverrides, profileDir }));
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

  // Targets:
  // - Full startup: preserve existing env var naming (cold/warm targets + legacy fallbacks).
  // - Shell startup: allow a parallel set of env vars (prefixed with FORMULA_DESKTOP_SHELL_*),
  //   but fall back to the full targets so CI can start using this without extra configuration.
  const fullColdWindowTarget =
    Number(
      process.env.FORMULA_DESKTOP_COLD_WINDOW_VISIBLE_TARGET_MS ??
        process.env.FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS ??
        '500',
    ) || 500;
  const fullColdFirstRenderTarget =
    Number(
      process.env.FORMULA_DESKTOP_COLD_FIRST_RENDER_TARGET_MS ??
        process.env.FORMULA_DESKTOP_FIRST_RENDER_TARGET_MS ??
        '500',
    ) || 500;
  const fullColdTtiTarget =
    Number(
      process.env.FORMULA_DESKTOP_COLD_TTI_TARGET_MS ??
        process.env.FORMULA_DESKTOP_TTI_TARGET_MS ??
        '1000',
    ) || 1000;

  const fullWarmWindowTarget =
    Number(process.env.FORMULA_DESKTOP_WARM_WINDOW_VISIBLE_TARGET_MS ?? String(fullColdWindowTarget)) ||
    fullColdWindowTarget;
  const fullWarmFirstRenderTarget =
    Number(process.env.FORMULA_DESKTOP_WARM_FIRST_RENDER_TARGET_MS ?? String(fullColdFirstRenderTarget)) ||
    fullColdFirstRenderTarget;
  const fullWarmTtiTarget =
    Number(process.env.FORMULA_DESKTOP_WARM_TTI_TARGET_MS ?? String(fullColdTtiTarget)) || fullColdTtiTarget;

  const shellColdWindowTarget =
    Number(
      process.env.FORMULA_DESKTOP_SHELL_COLD_WINDOW_VISIBLE_TARGET_MS ??
        process.env.FORMULA_DESKTOP_SHELL_WINDOW_VISIBLE_TARGET_MS ??
        String(fullColdWindowTarget),
    ) || fullColdWindowTarget;
  const shellColdTtiTarget =
    Number(
      process.env.FORMULA_DESKTOP_SHELL_COLD_TTI_TARGET_MS ??
        process.env.FORMULA_DESKTOP_SHELL_TTI_TARGET_MS ??
        String(fullColdTtiTarget),
    ) || fullColdTtiTarget;
  const shellWarmWindowTarget =
    Number(
      process.env.FORMULA_DESKTOP_SHELL_WARM_WINDOW_VISIBLE_TARGET_MS ??
        process.env.FORMULA_DESKTOP_SHELL_WINDOW_VISIBLE_TARGET_MS ??
        String(shellColdWindowTarget),
    ) || shellColdWindowTarget;
  const shellWarmTtiTarget =
    Number(
      process.env.FORMULA_DESKTOP_SHELL_WARM_TTI_TARGET_MS ??
        process.env.FORMULA_DESKTOP_SHELL_TTI_TARGET_MS ??
        String(shellColdTtiTarget),
    ) || shellColdTtiTarget;

  const windowTarget =
    benchKind === 'shell'
      ? startupMode === 'warm'
        ? shellWarmWindowTarget
        : shellColdWindowTarget
      : startupMode === 'warm'
        ? fullWarmWindowTarget
        : fullColdWindowTarget;
  const ttiTarget =
    benchKind === 'shell'
      ? startupMode === 'warm'
        ? shellWarmTtiTarget
        : shellColdTtiTarget
      : startupMode === 'warm'
        ? fullWarmTtiTarget
        : fullColdTtiTarget;
  const firstRenderTarget = startupMode === 'warm' ? fullWarmFirstRenderTarget : fullColdFirstRenderTarget;

  const metricPrefix = benchKind === 'shell' ? 'desktop.shell_startup' : 'desktop.startup';

  if (benchKind === 'full' && firstRender.length !== metrics.length) {
    throw new Error(
      'Desktop did not report first_render_ms. Ensure the frontend calls `report_startup_first_render` when the grid becomes visible.',
    );
  }

  const results: BenchmarkResult[] = [
    buildResult(`${metricPrefix}.${startupMode}.window_visible_ms.p95`, windowVisible, windowTarget),
  ];
  if (benchKind === 'full') {
    results.push(
      buildResult(`${metricPrefix}.${startupMode}.first_render_ms.p95`, firstRender, firstRenderTarget),
    );
  }
  results.push(buildResult(`${metricPrefix}.${startupMode}.tti_ms.p95`, tti, ttiTarget));

  // Convenience aliases:
  // - For cold-start runs, also expose a stable top-level metric name (no `.cold` suffix).
  // - For full startup, keep backwards-compatible legacy aliases.
  if (startupMode === 'cold') {
    if (benchKind === 'shell') {
      results.push(buildResult('desktop.shell_startup.window_visible_ms.p95', windowVisible, windowTarget));
      results.push(buildResult('desktop.shell_startup.tti_ms.p95', tti, ttiTarget));
    } else {
      results.push(buildResult('desktop.startup.window_visible_ms.p95', windowVisible, windowTarget));
      results.push(buildResult('desktop.startup.first_render_ms.p95', firstRender, firstRenderTarget));
      results.push(buildResult('desktop.startup.tti_ms.p95', tti, ttiTarget));
    }
  }

  // `webview_loaded_ms` is recorded by the Rust host (via a native page-load callback) and should
  // be available in all runs. Keep this best-effort skip policy anyway so the benchmark harness
  // can still run against older binaries and so a regression doesn't fail the entire suite on a
  // small/biased sample.
  //
  // Policy:
  // - If 0 runs report `webview_loaded_ms`, skip the metric entirely.
  // - If fewer than 80% of runs report it, skip the metric (avoid biased p95 on a tiny subset).
  // - Otherwise, compute p95 over the runs that reported a value and gate on the target.
  //
  // For now we only publish `webview_loaded_ms` as a cold-start metric (to keep the metric set
  // small and comparable over time), matching the historical behavior.
  if (startupMode === 'cold') {
    const webviewLoadedTarget =
      benchKind === 'shell'
        ? Number(
            process.env.FORMULA_DESKTOP_SHELL_WEBVIEW_LOADED_TARGET_MS ??
              process.env.FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS ??
              '800',
          ) || 800
        : Number(process.env.FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS ?? '800') || 800;
    const minValidFraction = 0.8;
    const minValidRuns = Math.ceil(runs * minValidFraction);
    if (webviewLoaded.length === 0) {
      // eslint-disable-next-line no-console
      console.log(
        `[desktop-${benchKind}-startup] webview_loaded_ms unavailable (0 runs reported it); skipping metric`,
      );
    } else if (webviewLoaded.length < minValidRuns) {
      // eslint-disable-next-line no-console
      console.log(
        `[desktop-${benchKind}-startup] webview_loaded_ms only available for ${webviewLoaded.length}/${runs} runs (<${Math.round(
          minValidFraction * 100,
        )}%); skipping metric`,
      );
    } else {
      const name =
        benchKind === 'shell'
          ? 'desktop.shell_startup.webview_loaded_ms.p95'
          : 'desktop.startup.webview_loaded_ms.p95';
      results.push(buildResult(name, webviewLoaded, webviewLoadedTarget));
    }
  }

  return results;
}
