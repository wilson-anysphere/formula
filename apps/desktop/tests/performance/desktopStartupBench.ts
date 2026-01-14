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
 * - All platforms: `HOME` + `USERPROFILE` => a per-run directory under `target/perf-home`
 *   (override the perf root via `FORMULA_PERF_HOME`).
 * - Linux: `XDG_CONFIG_HOME`, `XDG_CACHE_HOME`, `XDG_DATA_HOME` => `${HOME}/xdg-*`
 * - Windows: `APPDATA`, `LOCALAPPDATA`, `TEMP`, `TMP` => `${HOME}/*`
 * - macOS/Linux: `TMPDIR` => `${HOME}/tmp`
 *
 * Reset toggle:
 * - `FORMULA_DESKTOP_BENCH_RESET_HOME=1` deletes the per-run profile directory before spawning
 *   the desktop process. (This is also used internally by the cold/warm startup modes below.)
 *
 * Startup modes (profile reset behavior):
 * - `FORMULA_DESKTOP_STARTUP_MODE=cold` (default when enabled): each iteration uses a fresh
 *   profile directory under `target/perf-home` so every launch is a true cold start.
 * - `FORMULA_DESKTOP_STARTUP_MODE=warm`: a single profile directory is initialized once (warmup),
 *   then reused for the measured runs so persisted caches are reflected in the results.
 *
 * Benchmark kind (what we measure):
 * - `FORMULA_DESKTOP_STARTUP_BENCH_KIND=full` (default; requires built frontend assets): launch the normal app.
 * - `FORMULA_DESKTOP_STARTUP_BENCH_KIND=shell`: launch `--startup-bench` (measures shell/webview startup without
 *   requiring `apps/desktop/dist`).
 *
 * Optional idle RSS metric:
 * - Metric: `desktop.memory.<mode>.rss_mb.p95` (and `desktop.memory.rss_mb.p95` alias for cold mode)
 * - Target: `FORMULA_DESKTOP_RSS_TARGET_MB` (default: 100)
 * - Delay after capturing `[startup] ... tti_ms=...` before sampling:
 *   `FORMULA_DESKTOP_RSS_IDLE_DELAY_MS` (default: 1000)
 *
 * Platform support for RSS:
 * - Linux (CI primary): reads `/proc/<pid>/status` (VmRSS) for the desktop PID (resolving through
 *   any Xvfb wrapper processes).
 * - macOS: best-effort via `ps`.
 * - Windows: best-effort via PowerShell (`Get-Process ... WorkingSet64`).
 * - Other platforms: skipped.
 *
 * RSS measurement is best-effort; if we can't sample RSS, we skip the memory metric rather than
 * failing the timing benchmarks.
 */
import { spawnSync } from 'node:child_process';
import { existsSync, realpathSync } from 'node:fs';
import { readFile, readlink, readdir } from 'node:fs/promises';
import { resolve } from 'node:path';

import { type BenchmarkResult } from './benchmark.ts';
import {
  defaultDesktopBinPath,
  mean,
  median,
  percentile,
  buildDesktopStartupProfileRoot,
  runDesktopStartupIterations,
  resolveDesktopStartupArgv,
  resolveDesktopStartupBenchKind,
  resolveDesktopStartupMode,
  resolveDesktopStartupTargets,
  resolvePerfHome,
  stdDev,
  type StartupMetrics,
} from './desktopStartupUtil.ts';

// Benchmark environment knobs:
// - `FORMULA_DISABLE_STARTUP_UPDATE_CHECK=1` prevents the release updater from running a
//   background check/download on startup, which can add nondeterministic CPU/memory/network
//   activity and skew startup/idle-memory benchmarks.
// - `FORMULA_STARTUP_METRICS=1` enables the Rust-side one-line startup metrics log we parse.

function buildResult(
  name: string,
  values: number[],
  target: number,
  unit: BenchmarkResult['unit'],
): BenchmarkResult {
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
    unit,
    mean: avg,
    median: med,
    p95,
    p99,
    stdDev: sd,
    targetMs: target,
    passed: p95 <= target,
  };
}

async function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  await new Promise<void>((resolvePromise, rejectPromise) => {
    const timer = setTimeout(() => {
      cleanup();
      resolvePromise();
    }, ms);
    const cleanup = () => {
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
    };
    const onAbort = () => {
      cleanup();
      rejectPromise(new Error('aborted'));
    };
    if (signal) {
      if (signal.aborted) {
        onAbort();
        return;
      }
      signal.addEventListener('abort', onAbort);
    }
  });
}

function parseProcChildrenPids(content: string): number[] {
  const trimmed = content.trim();
  if (!trimmed) return [];
  return trimmed
    .split(/\s+/g)
    .map((x) => Number(x))
    .filter((n) => Number.isInteger(n) && n > 0);
}

function parseProcStatusVmRssKb(content: string): number | null {
  const match = content.match(/^VmRSS:\s+(\d+)\s+kB\s*$/m);
  if (!match) return null;
  const kb = Number(match[1]);
  if (!Number.isFinite(kb)) return null;
  return kb;
}

async function readUtf8(path: string): Promise<string | null> {
  try {
    return await readFile(path, 'utf8');
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT' || code === 'ESRCH' || code === 'EACCES') return null;
    throw err;
  }
}

async function readProcExeLinux(pid: number): Promise<string | null> {
  try {
    const target = await readlink(`/proc/${pid}/exe`);
    // If the binary was replaced/cleaned up mid-run, Linux appends " (deleted)".
    return target.replace(/ \(deleted\)$/, '');
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT' || code === 'ESRCH' || code === 'EACCES') return null;
    throw err;
  }
}

async function getChildPidsLinux(pid: number): Promise<number[]> {
  // `/proc/<pid>/task/<tid>/children` is per-thread (not per-process). Union children across
  // all tasks so we don't miss descendants spawned by worker threads.
  let tids: string[];
  try {
    tids = await readdir(`/proc/${pid}/task`);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT' || code === 'ESRCH' || code === 'EACCES') return [];
    throw err;
  }

  const out = new Set<number>();
  for (const tid of tids) {
    const content = await readUtf8(`/proc/${pid}/task/${tid}/children`);
    if (!content) continue;
    for (const child of parseProcChildrenPids(content)) {
      out.add(child);
    }
  }

  return [...out];
}

async function collectProcessTreePidsLinux(rootPid: number): Promise<number[]> {
  const seen = new Set<number>();
  const stack: number[] = [rootPid];
  while (stack.length > 0) {
    const pid = stack.pop()!;
    if (seen.has(pid)) continue;
    seen.add(pid);
    const children = await getChildPidsLinux(pid);
    for (const child of children) {
      if (!seen.has(child)) stack.push(child);
    }
  }
  return [...seen];
}

async function findPidForExecutableLinux(
  rootPid: number,
  binPath: string,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<number | null> {
  const binResolved = resolve(binPath);
  let binReal = binResolved;
  try {
    binReal = realpathSync(binResolved);
  } catch {
    // Best-effort; realpath can fail in some sandbox setups.
  }

  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (signal?.aborted) return null;
    const pids = await collectProcessTreePidsLinux(rootPid);
    for (const pid of pids) {
      const exe = await readProcExeLinux(pid);
      if (!exe) continue;
      if (exe === binReal || exe === binResolved) return pid;
    }
    await sleep(50, signal);
  }
  return null;
}

async function getProcessRssMbLinux(pid: number): Promise<number | null> {
  const status = await readUtf8(`/proc/${pid}/status`);
  if (!status) return null;
  const kb = parseProcStatusVmRssKb(status);
  if (kb == null) return null;
  return kb / 1024;
}

function getRssMbViaPs(pid: number): number | null {
  try {
    const proc = spawnSync('ps', ['-o', 'rss=', '-p', String(pid)], {
      encoding: 'utf8',
      timeout: 5000,
      maxBuffer: 1024 * 1024,
    });
    if (proc.status !== 0) return null;
    const kb = Number(proc.stdout.trim());
    if (!Number.isFinite(kb)) return null;
    return kb / 1024;
  } catch {
    return null;
  }
}

function getRssMbViaPowerShell(pid: number): number | null {
  try {
    const proc = spawnSync(
      'powershell.exe',
      ['-NoProfile', '-Command', `(Get-Process -Id ${pid}).WorkingSet64`],
      { encoding: 'utf8', timeout: 15000, maxBuffer: 1024 * 1024, windowsHide: true },
    );
    if (proc.status !== 0) return null;
    const bytes = Number(proc.stdout.trim());
    if (!Number.isFinite(bytes)) return null;
    return bytes / (1024 * 1024);
  } catch {
    return null;
  }
}

async function captureDesktopRssMb(
  childPid: number,
  binPath: string,
  idleDelayMs: number,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<number | null> {
  try {
    await sleep(idleDelayMs, signal);

    if (process.platform === 'linux') {
      // When running under Xvfb, `childPid` is the wrapper process group leader. Resolve the real
      // desktop PID by executable path.
      const desktopPid = await findPidForExecutableLinux(
        childPid,
        binPath,
        Math.min(2000, timeoutMs),
        signal,
      );
      if (!desktopPid) return null;
      return await getProcessRssMbLinux(desktopPid);
    }

    if (process.platform === 'darwin') {
      return getRssMbViaPs(childPid);
    }

    if (process.platform === 'win32') {
      return getRssMbViaPowerShell(childPid);
    }
  } catch {
    // Best-effort: RSS measurement failures should not fail the benchmark run.
    return null;
  }

  return null;
}

export async function runDesktopStartupBenchmarks(): Promise<BenchmarkResult[]> {
  if (process.env.FORMULA_RUN_DESKTOP_STARTUP_BENCH !== '1') {
    return [];
  }

  const startupMode = resolveDesktopStartupMode({ defaultMode: 'cold' });
  // Prefer the lightweight shell benchmark in CI so we can measure startup without requiring
  // built frontend assets. Locally, default to full end-to-end startup timings.
  const benchKind = resolveDesktopStartupBenchKind({ defaultKind: process.env.CI ? 'shell' : 'full' });

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
  const argv = resolveDesktopStartupArgv(benchKind);

  const rssIdleDelayMs = Math.max(
    0,
    Number(process.env.FORMULA_DESKTOP_RSS_IDLE_DELAY_MS ?? '1000') || 1000,
  );
  const rssTargetMb = Number(process.env.FORMULA_DESKTOP_RSS_TARGET_MB ?? '100') || 100;

  const perfHome = resolvePerfHome();
  const profileRoot = buildDesktopStartupProfileRoot({ perfHome, benchKind, mode: startupMode });

  const metrics: StartupMetrics[] = [];
  const rssSamples: number[] = [];

  metrics.push(
    ...(await runDesktopStartupIterations({
      mode: startupMode,
      runs,
      timeoutMs,
      binPath,
      argv,
      envOverrides,
      profileRoot,
      afterCapture: async (child, _metrics, signal) => {
        if (!child.pid) return;
        const rssMb = await captureDesktopRssMb(child.pid, binPath, rssIdleDelayMs, timeoutMs, signal);
        if (rssMb != null && Number.isFinite(rssMb)) rssSamples.push(rssMb);
      },
      afterCaptureTimeoutMs: rssIdleDelayMs + 4000,
      onProgress: ({ phase, mode, iteration, total, profileDir }) => {
        // eslint-disable-next-line no-console
        if (phase === 'warmup') {
          console.log(`[desktop-${benchKind}-startup] warmup run 1/1 (warm, profile=${profileDir})...`);
        } else {
          console.log(`[desktop-${benchKind}-startup] run ${iteration}/${total} (${mode}, profile=${profileDir})...`);
        }
      },
    })),
  );

  const windowVisible = metrics.map((m) => m.windowVisibleMs);
  const firstRender = metrics
    .map((m) => m.firstRenderMs)
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
  const tti = metrics.map((m) => m.ttiMs);
  const webviewLoaded = metrics
    .map((m) => m.webviewLoadedMs)
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v));

  const targets = resolveDesktopStartupTargets({ benchKind, mode: startupMode });
  const windowTarget = targets.windowVisibleTargetMs;
  const ttiTarget = targets.ttiTargetMs;
  const firstRenderTarget = targets.firstRenderTargetMs;

  const metricPrefix = benchKind === 'shell' ? 'desktop.shell_startup' : 'desktop.startup';
  const memoryPrefix = benchKind === 'shell' ? 'desktop.shell_memory' : 'desktop.memory';

  if (benchKind === 'full' && firstRender.length !== metrics.length) {
    throw new Error(
      'Desktop did not report first_render_ms. Ensure the frontend calls `report_startup_first_render` when the grid becomes visible.',
    );
  }

  const results: BenchmarkResult[] = [
    buildResult(`${metricPrefix}.${startupMode}.window_visible_ms.p95`, windowVisible, windowTarget, 'ms'),
  ];
  if (benchKind === 'full') {
    results.push(
      buildResult(`${metricPrefix}.${startupMode}.first_render_ms.p95`, firstRender, firstRenderTarget, 'ms'),
    );
  }
  results.push(buildResult(`${metricPrefix}.${startupMode}.tti_ms.p95`, tti, ttiTarget, 'ms'));

  const minRssValidFraction = 0.8;
  const minRssValidRuns = Math.ceil(runs * minRssValidFraction);
  if (rssSamples.length >= minRssValidRuns) {
    results.push(buildResult(`${memoryPrefix}.${startupMode}.rss_mb.p95`, rssSamples, rssTargetMb, 'mb'));
  } else if (rssSamples.length > 0) {
    // eslint-disable-next-line no-console
    console.log(
      `[desktop-${benchKind}-startup] rss_mb only available for ${rssSamples.length}/${runs} runs (<${Math.round(
        minRssValidFraction * 100,
      )}%); skipping memory metric`,
    );
  }

  // Convenience aliases:
  // - For cold-start runs, also expose a stable top-level metric name (no `.cold` suffix).
  // - For full startup, keep backwards-compatible legacy aliases.
  if (startupMode === 'cold') {
    if (benchKind === 'shell') {
      results.push(buildResult('desktop.shell_startup.window_visible_ms.p95', windowVisible, windowTarget, 'ms'));
      results.push(buildResult('desktop.shell_startup.tti_ms.p95', tti, ttiTarget, 'ms'));
    } else {
      results.push(buildResult('desktop.startup.window_visible_ms.p95', windowVisible, windowTarget, 'ms'));
      results.push(buildResult('desktop.startup.first_render_ms.p95', firstRender, firstRenderTarget, 'ms'));
      results.push(buildResult('desktop.startup.tti_ms.p95', tti, ttiTarget, 'ms'));
    }

    if (rssSamples.length >= minRssValidRuns) {
      results.push(buildResult(`${memoryPrefix}.rss_mb.p95`, rssSamples, rssTargetMb, 'mb'));
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
    const webviewLoadedTarget = targets.webviewLoadedTargetMs;
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
      results.push(buildResult(name, webviewLoaded, webviewLoadedTarget, 'ms'));
    }
  }

  return results;
}
