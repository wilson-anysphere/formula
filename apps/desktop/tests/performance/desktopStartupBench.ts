import { spawn } from 'node:child_process';
import { existsSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';

import { type BenchmarkResult } from './benchmark.ts';

type StartupMetrics = {
  windowVisibleMs: number;
  webviewLoadedMs: number | null;
  ttiMs: number;
};

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

function mean(values: number[]): number {
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0;
  const idx = Math.floor(sorted.length * p);
  return sorted[Math.min(idx, sorted.length - 1)]!;
}

function median(sorted: number[]): number {
  return sorted[Math.floor(sorted.length / 2)]!;
}

function stdDev(values: number[], avg: number): number {
  const variance = values.reduce((sum, x) => sum + Math.pow(x - avg, 2), 0) / values.length;
  return Math.sqrt(variance);
}

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

function defaultDesktopBinPath(): string | null {
  const exe = process.platform === 'win32' ? 'formula-desktop.exe' : 'formula-desktop';
  const candidates = [
    resolve(repoRoot, 'target', 'release', exe),
    resolve(repoRoot, 'target', 'debug', exe),
    resolve(repoRoot, 'apps/desktop/src-tauri/target', 'release', exe),
    resolve(repoRoot, 'apps/desktop/src-tauri/target', 'debug', exe),
  ];
  for (const p of candidates) {
    if (existsSync(p)) return p;
  }
  return null;
}

function shouldUseXvfb(): boolean {
  if (process.platform !== 'linux') return false;
  if (process.env.DISPLAY && process.env.DISPLAY.trim() !== '') return false;
  const xvfb = resolve(repoRoot, 'scripts/xvfb-run-safe.sh');
  return existsSync(xvfb);
}

function parseStartupLine(line: string): StartupMetrics | null {
  // Example:
  // [startup] window_visible_ms=123 webview_loaded_ms=234 tti_ms=456
  const match = line.match(
    /^\[startup\]\s+window_visible_ms=(\d+)\s+webview_loaded_ms=(\d+|n\/a)\s+tti_ms=(\d+)\s*$/,
  );
  if (!match) return null;
  const windowVisibleMs = Number(match[1]);
  const webviewLoadedRaw = match[2]!;
  const webviewLoadedMs = webviewLoadedRaw === 'n/a' ? null : Number(webviewLoadedRaw);
  const ttiMs = Number(match[3]);
  if (!Number.isFinite(windowVisibleMs) || !Number.isFinite(ttiMs)) return null;
  return { windowVisibleMs, webviewLoadedMs, ttiMs };
}

async function runOnce(binPath: string, timeoutMs: number): Promise<StartupMetrics> {
  const useXvfb = shouldUseXvfb();
  const command = useXvfb ? resolve(repoRoot, 'scripts/xvfb-run-safe.sh') : binPath;
  const args = useXvfb ? [binPath] : [];

  const perfHome = resolve(repoRoot, 'target', 'perf-home');
  mkdirSync(perfHome, { recursive: true });

  return await new Promise<StartupMetrics>((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: ['ignore', 'pipe', 'pipe'],
      // When running under the Xvfb wrapper we want to be able to reliably
      // terminate the whole process tree (wrapper + Xvfb + desktop binary).
      detached: useXvfb,
      env: {
        ...process.env,
        // Enable the Rust-side single-line log in release builds.
        FORMULA_STARTUP_METRICS: '1',
        // In case the app reads $HOME for config, keep per-run caches out of the real home dir.
        HOME: perfHome,
        USERPROFILE: perfHome,
      },
    });

    const killChild = (signal: NodeJS.Signals = 'SIGTERM') => {
      // On Linux headless runs we wrap the binary in xvfb-run-safe.sh, which
      // starts additional processes (Xvfb + the actual desktop binary). If we
      // only kill the wrapper, the child processes can be orphaned and linger
      // into subsequent iterations.
      if (useXvfb && child.pid) {
        try {
          process.kill(-child.pid, signal);
          return;
        } catch {
          // Fall back to killing the wrapper process itself.
        }
      }
      try {
        child.kill(signal);
      } catch {
        // ignore
      }
    };

    let done = false;
    let captured: StartupMetrics | null = null;
    let captureKillTimer: NodeJS.Timeout | null = null;
    let exitDeadline: NodeJS.Timeout | null = null;

    const deadline = setTimeout(() => {
      if (done) return;
      done = true;
      killChild();
      cleanup();
      rejectPromise(new Error(`Timed out after ${timeoutMs}ms waiting for startup metrics`));
    }, timeoutMs);

    const cleanup = () => {
      clearTimeout(deadline);
      if (captureKillTimer) clearTimeout(captureKillTimer);
      if (exitDeadline) clearTimeout(exitDeadline);
      rlOut.close();
      rlErr.close();
    };

    const onLine = (line: string) => {
      if (done || captured) return;
      const parsed = parseStartupLine(line.trim());
      if (!parsed) return;
      captured = parsed;
      // We got the data we came for; don't fail the run just because shutdown is slow.
      clearTimeout(deadline);

      // Stop the app after capturing the metrics so we can run multiple iterations.
      killChild();
      exitDeadline = setTimeout(() => {
        if (done) return;
        done = true;
        try {
          killChild('SIGKILL');
        } catch {
          killChild();
        }
        cleanup();
        rejectPromise(new Error('Timed out waiting for desktop process to exit after capturing metrics'));
      }, 5000);

      // If the process doesn't exit quickly, force-kill it so we don't accumulate
      // background GUI processes during a multi-run benchmark.
      captureKillTimer = setTimeout(() => {
        try {
          killChild('SIGKILL');
        } catch {
          killChild();
        }
      }, 2000);
    };

    const rlOut = createInterface({ input: child.stdout! });
    const rlErr = createInterface({ input: child.stderr! });
    rlOut.on('line', onLine);
    rlErr.on('line', onLine);

    child.on('error', (err) => {
      if (done) return;
      done = true;
      cleanup();
      rejectPromise(err);
    });

    child.on('exit', (code, signal) => {
      cleanup();
      if (done) return;
      done = true;
      if (captured) {
        resolvePromise(captured);
        return;
      }
      rejectPromise(
        new Error(`Desktop process exited before reporting metrics (code=${code}, signal=${signal})`),
      );
    });
  });
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
    metrics.push(await runOnce(binPath, timeoutMs));
  }

  const windowVisible = metrics.map((m) => m.windowVisibleMs);
  const tti = metrics.map((m) => m.ttiMs);

  const windowTarget = Number(process.env.FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS ?? '500') || 500;
  const ttiTarget = Number(process.env.FORMULA_DESKTOP_TTI_TARGET_MS ?? '1000') || 1000;

  return [
    buildResult('desktop.startup.window_visible_ms.p95', windowVisible, windowTarget),
    buildResult('desktop.startup.tti_ms.p95', tti, ttiTarget),
  ];
}
