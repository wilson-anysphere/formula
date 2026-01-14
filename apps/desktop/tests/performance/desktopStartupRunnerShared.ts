import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, isAbsolute, parse, resolve, relative } from 'node:path';
import { createInterface, type Interface } from 'node:readline';
import { fileURLToPath } from 'node:url';

import { terminateProcessTree, type TerminateProcessTreeMode } from './processTree.ts';

export type StartupMetrics = {
  windowVisibleMs: number;
  webviewLoadedMs: number | null;
  firstRenderMs: number | null;
  ttiMs: number;
};

// Ensure paths are rooted at repo root even when invoked from elsewhere.
export const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

function resolvePerfHome(): string {
  const fromEnv = process.env.FORMULA_PERF_HOME;
  if (fromEnv && fromEnv.trim() !== '') {
    // `resolve(repoRoot, ...)` safely handles both absolute and relative paths.
    return resolve(repoRoot, fromEnv);
  }
  return resolve(repoRoot, 'target', 'perf-home');
}

const perfHome = resolvePerfHome();

function isSubpath(parentDir: string, maybeChild: string): boolean {
  const rel = relative(parentDir, maybeChild);
  if (rel === '' || rel.startsWith('..')) return false;
  // `path.relative()` can return an absolute path on Windows when drives differ.
  if (isAbsolute(rel)) return false;
  return true;
}

function resolveProfileDirs(profileDir: string): {
  home: string;
  tmp: string;
  xdgConfig: string;
  xdgCache: string;
  xdgState: string;
  xdgData: string;
  appData: string;
  localAppData: string;
} {
  return {
    home: profileDir,
    tmp: resolve(profileDir, 'tmp'),
    xdgConfig: resolve(profileDir, 'xdg-config'),
    xdgCache: resolve(profileDir, 'xdg-cache'),
    xdgState: resolve(profileDir, 'xdg-state'),
    xdgData: resolve(profileDir, 'xdg-data'),
    appData: resolve(profileDir, 'AppData', 'Roaming'),
    localAppData: resolve(profileDir, 'AppData', 'Local'),
  };
}

export function defaultDesktopBinPath(): string | null {
  const exe = process.platform === 'win32' ? 'formula-desktop.exe' : 'formula-desktop';
  const candidates = [
    // Cargo workspace default target dir (most common).
    resolve(repoRoot, 'target', 'release', exe),
    resolve(repoRoot, 'target', 'debug', exe),
    // Fallbacks in case a caller built with a custom target dir rooted under the app.
    resolve(repoRoot, 'apps/desktop/src-tauri/target', 'release', exe),
    resolve(repoRoot, 'apps/desktop/src-tauri/target', 'debug', exe),
  ];
  for (const p of candidates) {
    if (existsSync(p)) return p;
  }
  return null;
}

export function shouldUseXvfb(): boolean {
  if (process.platform !== 'linux') return false;
  const xvfb = resolve(repoRoot, 'scripts/xvfb-run-safe.sh');
  if (!existsSync(xvfb)) return false;

  // In CI/headless environments it is common for $DISPLAY to be set but unusable.
  // `xvfb-run-safe.sh` is conservative about trusting DISPLAY, so prefer it in CI.
  if (process.env.CI) return true;

  // If DISPLAY is unset, we need Xvfb.
  if (!process.env.DISPLAY || process.env.DISPLAY.trim() === '') return true;
  return false;
}

export function mean(values: number[]): number {
  return values.reduce((a, b) => a + b, 0) / values.length;
}

/**
 * Percentile over a sorted array.
 *
 * Matches the implementation used by `apps/desktop/tests/performance/benchmark.ts`.
 */
export function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0;
  const idx = Math.floor(sorted.length * p);
  return sorted[Math.min(idx, sorted.length - 1)]!;
}

export function median(sorted: number[]): number {
  return sorted[Math.floor(sorted.length / 2)]!;
}

export function stdDev(values: number[], avg: number): number {
  const variance = values.reduce((sum, x) => sum + Math.pow(x - avg, 2), 0) / values.length;
  return Math.sqrt(variance);
}

export function parseStartupLine(line: string): StartupMetrics | null {
  // Example:
  // [startup] window_visible_ms=123 webview_loaded_ms=234 first_render_ms=345 tti_ms=456
  const trimmed = line.trim();
  const idx = trimmed.indexOf('[startup]');
  if (idx === -1) return null;

  const payload = trimmed.slice(idx + '[startup]'.length).trim();
  if (payload.length === 0) return null;

  const entries = payload.split(/\s+/);
  const kv: Record<string, string> = {};
  for (const entry of entries) {
    const eq = entry.indexOf('=');
    if (eq <= 0) continue;
    const k = entry.slice(0, eq);
    const v = entry.slice(eq + 1);
    if (!k || !v) continue;
    kv[k] = v;
  }

  const windowVisibleRaw = kv['window_visible_ms'];
  const ttiRaw = kv['tti_ms'];
  if (!windowVisibleRaw || !ttiRaw) return null;

  const windowVisibleMs = Number(windowVisibleRaw);
  const ttiMs = Number(ttiRaw);
  if (!Number.isFinite(windowVisibleMs) || !Number.isFinite(ttiMs)) return null;

  const webviewLoadedRaw = kv['webview_loaded_ms'];
  const webviewLoadedMs =
    !webviewLoadedRaw || webviewLoadedRaw === 'n/a' ? null : Number(webviewLoadedRaw);
  if (webviewLoadedMs !== null && !Number.isFinite(webviewLoadedMs)) return null;

  const firstRenderRaw = kv['first_render_ms'];
  const firstRenderMs =
    !firstRenderRaw || firstRenderRaw === 'n/a' ? null : Number(firstRenderRaw);
  if (firstRenderMs !== null && !Number.isFinite(firstRenderMs)) return null;

  return { windowVisibleMs, webviewLoadedMs, firstRenderMs, ttiMs };
}

type RunOnceOptions = {
  binPath: string;
  timeoutMs: number;
  /**
   * Extra CLI args to pass to the desktop binary.
   *
   * This is primarily used for special/CI modes like `--startup-bench` that should not
   * require bundled frontend assets.
   */
  argv?: string[];
  envOverrides?: NodeJS.ProcessEnv;
  /**
   * Root directory for all app/user-data state for this run (HOME, XDG dirs, temp dirs, etc).
   *
   * If omitted, defaults to `target/perf-home` (or `FORMULA_PERF_HOME` if set).
   */
  profileDir?: string;
  /**
   * Optional hook invoked after startup metrics are captured but before the process is terminated.
   *
   * This is used by benchmarks that need to take a final measurement (e.g. RSS) while the app is
   * still alive. The hook is best-effort: any error is ignored and shutdown proceeds.
   */
  afterCapture?: (
    child: ChildProcess,
    metrics: StartupMetrics,
    signal: AbortSignal,
  ) => void | Promise<void>;
  /**
   * Maximum time to wait for `afterCapture` before proceeding with shutdown.
   *
   * This prevents the benchmark harness from hanging indefinitely if the hook blocks.
   */
  afterCaptureTimeoutMs?: number;
};

function mergeEnvParts(parts: Array<NodeJS.ProcessEnv | undefined>): NodeJS.ProcessEnv {
  const out: NodeJS.ProcessEnv = {};
  for (const part of parts) {
    if (!part) continue;
    for (const [k, v] of Object.entries(part)) {
      if (v === undefined) {
        // Allow overrides to delete inherited vars (useful for ensuring isolation).
        delete out[k];
        continue;
      }
      out[k] = v;
    }
  }
  return out;
}

function closeReadline(rl: Interface | null): void {
  if (!rl) return;
  try {
    rl.close();
  } catch {
    // ignore
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

export async function runOnce({
  binPath,
  timeoutMs,
  argv,
  envOverrides,
  profileDir: profileDirRaw,
  afterCapture,
  afterCaptureTimeoutMs,
}: RunOnceOptions): Promise<StartupMetrics> {
  const profileDir = profileDirRaw ? resolve(repoRoot, profileDirRaw) : perfHome;
  const dirs = resolveProfileDirs(profileDir);
  // Best-effort isolation: keep the desktop app from mutating a developer's real home directory.
  // Optionally, force a clean state between iterations to avoid cache pollution.
  if (process.env.FORMULA_DESKTOP_BENCH_RESET_HOME === '1') {
    // Extra guardrails: if a caller misconfigures `FORMULA_PERF_HOME` / `profileDir`, avoid
    // `rm -rf /` style footguns.
    const rootDir = parse(profileDir).root;
    if (profileDir === rootDir || profileDir === repoRoot) {
      throw new Error(`Refusing to reset unsafe desktop benchmark profile dir: ${profileDir}`);
    }

    const safeRoot = perfHome;
    if (profileDir !== safeRoot && !isSubpath(safeRoot, profileDir)) {
      throw new Error(
        `Refusing to reset desktop benchmark profile dir outside ${safeRoot} (got ${profileDir})`,
      );
    }
    rmSync(profileDir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 });
  }

  mkdirSync(dirs.home, { recursive: true });
  mkdirSync(dirs.tmp, { recursive: true });
  mkdirSync(dirs.xdgConfig, { recursive: true });
  mkdirSync(dirs.xdgCache, { recursive: true });
  mkdirSync(dirs.xdgState, { recursive: true });
  mkdirSync(dirs.xdgData, { recursive: true });
  mkdirSync(dirs.appData, { recursive: true });
  mkdirSync(dirs.localAppData, { recursive: true });

  const useXvfb = shouldUseXvfb();
  const xvfbPath = resolve(repoRoot, 'scripts/xvfb-run-safe.sh');
  const command = useXvfb ? 'bash' : binPath;
  const desktopArgs = argv ?? [];
  const args = useXvfb ? [xvfbPath, binPath, ...desktopArgs] : desktopArgs;

  const env = mergeEnvParts([
    process.env,
    {
      // Keep perf benchmarks stable/quiet by disabling the automatic startup update check
      // (which can add nondeterministic network/CPU activity).
      FORMULA_DISABLE_STARTUP_UPDATE_CHECK: '1',
      // Enable the Rust-side single-line log in release builds.
      FORMULA_STARTUP_METRICS: '1',
      // In case the app reads $HOME / XDG dirs for config, keep per-run caches out of the real home dir.
      HOME: dirs.home,
      USERPROFILE: dirs.home,
      XDG_CONFIG_HOME: dirs.xdgConfig,
      XDG_CACHE_HOME: dirs.xdgCache,
      XDG_STATE_HOME: dirs.xdgState,
      XDG_DATA_HOME: dirs.xdgData,
      APPDATA: dirs.appData,
      LOCALAPPDATA: dirs.localAppData,
      TMPDIR: dirs.tmp,
      TEMP: dirs.tmp,
      TMP: dirs.tmp,
    },
    envOverrides,
  ]);

  return await new Promise<StartupMetrics>((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: ['ignore', 'pipe', 'pipe'],
      env,
      // On POSIX, start the app in its own process group so we can kill the whole tree
      // without relying on any kill-by-name pattern.
      detached: process.platform !== 'win32',
      windowsHide: true,
    });

    let rlOut: Interface | null = null;
    let rlErr: Interface | null = null;

    let settled = false;
    let captured: StartupMetrics | null = null;
    let startupTimeout: NodeJS.Timeout | null = null;
    let forceKillTimer: NodeJS.Timeout | null = null;
    let exitDeadline: NodeJS.Timeout | null = null;
    let timedOutWaitingForMetrics = false;

    const cleanup = () => {
      if (startupTimeout) clearTimeout(startupTimeout);
      if (forceKillTimer) clearTimeout(forceKillTimer);
      if (exitDeadline) clearTimeout(exitDeadline);
      closeReadline(rlOut);
      closeReadline(rlErr);
    };

    const settle = (kind: 'resolve' | 'reject', value: any) => {
      if (settled) return;
      settled = true;
      cleanup();
      if (kind === 'resolve') resolvePromise(value);
      else rejectPromise(value);
    };

    const beginShutdown = (reason: 'captured' | 'timeout') => {
      if (settled) return;
      if (exitDeadline) return;

      const initialMode: TerminateProcessTreeMode =
        process.platform === 'win32' || reason === 'timeout' ? 'force' : 'graceful';

      // Stop the app after capturing the metrics so we can run multiple iterations. On POSIX we
      // send SIGTERM to the process group; on Windows we use `taskkill /T /F` to ensure WebView2
      // child processes don't survive across runs.
      terminateProcessTree(child, initialMode);

      // If the process doesn't exit quickly, force-kill it so we don't accumulate
      // background GUI processes during a multi-run benchmark.
      forceKillTimer = setTimeout(() => terminateProcessTree(child, 'force'), 2000);

      exitDeadline = setTimeout(() => {
        terminateProcessTree(child, 'force');

        // Extremely defensive: don't hang the parent process even if kill fails.
        try {
          child.unref();
        } catch {
          // ignore
        }
        try {
          child.stdout?.destroy();
        } catch {
          // ignore
        }
        try {
          child.stderr?.destroy();
        } catch {
          // ignore
        }

        const msg =
          reason === 'captured'
            ? 'Timed out waiting for desktop process to exit after capturing metrics'
            : 'Timed out waiting for desktop process to exit after timing out waiting for metrics';
        settle('reject', new Error(msg));
      }, 5000);
    };

    const onLine = (line: string) => {
      if (captured || timedOutWaitingForMetrics) return;
      const parsed = parseStartupLine(line);
      if (!parsed) return;
      captured = parsed;
      if (startupTimeout) {
        clearTimeout(startupTimeout);
        startupTimeout = null;
      }

      const hook = afterCapture;
      if (!hook) {
        beginShutdown('captured');
        return;
      }

      const hookTimeoutMs = afterCaptureTimeoutMs ?? 5000;
      void (async () => {
        const controller = new AbortController();
        let timer: NodeJS.Timeout | null = null;
        try {
          await Promise.race([
            Promise.resolve().then(() => hook(child, parsed, controller.signal)),
            new Promise<void>((resolve) => {
              timer = setTimeout(() => {
                controller.abort();
                resolve();
              }, Math.max(0, hookTimeoutMs));
            }),
          ]);
        } catch {
          // Best-effort hook: ignore errors and proceed to shutdown.
        } finally {
          if (timer) clearTimeout(timer);
        }
        beginShutdown('captured');
      })();
    };

    if (child.stdout) {
      rlOut = createInterface({ input: child.stdout });
      rlOut.on('line', onLine);
    }
    if (child.stderr) {
      rlErr = createInterface({ input: child.stderr });
      rlErr.on('line', onLine);
    }

    startupTimeout = setTimeout(() => {
      timedOutWaitingForMetrics = true;
      beginShutdown('timeout');
    }, timeoutMs);

    child.on('error', (err) => {
      settle('reject', err);
    });

    // Use `close` (not `exit`) so stdout/stderr are fully drained before we decide whether we
    // captured the `[startup] ...` line. This matters for modes like `--startup-bench` that exit
    // quickly after logging.
    child.on('close', (code, signal) => {
      if (settled) return;

      if (timedOutWaitingForMetrics) {
        settle('reject', new Error(`Timed out after ${timeoutMs}ms waiting for startup metrics`));
        return;
      }

      // If the process exits before we initiated shutdown (e.g. an `afterCapture` hook was
      // running), still attempt to tear down the full process group. WebView helper processes can
      // outlive the parent process and leak across runs.
      if (!exitDeadline) {
        terminateProcessTree(child, 'force');
      }

      if (captured) {
        settle('resolve', captured);
        return;
      }

      settle(
        'reject',
        new Error(`Desktop process exited before reporting metrics (code=${code}, signal=${signal})`),
      );
    });
  });
}
