/**
 * Desktop idle RSS benchmark (Linux-only).
 *
 * Reproducibility + safety:
 * - The desktop process is spawned with user-data directories redirected under
 *   `target/perf-home` (by default; override via `FORMULA_PERF_HOME`) so the benchmark cannot
 *   read/write the real user profile.
 * - This keeps caches deterministic on CI runners (and avoids polluting developer machines).
 *
 * Environment isolation:
 * - All platforms: `HOME` + `USERPROFILE` => a per-run directory under `target/perf-home`
 * - Linux: `XDG_CONFIG_HOME`, `XDG_CACHE_HOME`, `XDG_DATA_HOME` => `${HOME}/xdg-*`
 * - Windows: `APPDATA`, `LOCALAPPDATA`, `TEMP`, `TMP` => `${HOME}/*`
 * - macOS/Linux: `TMPDIR` => `${HOME}/tmp`
 *
 * Reset behavior:
 * - Set `FORMULA_DESKTOP_BENCH_RESET_HOME=1` to delete the benchmark profile directory (HOME)
 *   before *each* iteration.
 */

import { spawn, type ChildProcess } from 'node:child_process';
import { existsSync, mkdirSync, realpathSync, rmSync } from 'node:fs';
import { readFile, readlink, readdir } from 'node:fs/promises';
import { parse, resolve } from 'node:path';
import { createInterface } from 'node:readline';

import { type BenchmarkResult } from './benchmark.ts';
import {
  defaultDesktopBinPath,
  mean,
  median,
  percentile,
  parseStartupLine as parseStartupMetricsLine,
  repoRoot,
  shouldUseXvfb,
  terminateProcessTree,
  stdDev,
} from './desktopStartupUtil.ts';

function resolvePerfHome(): string {
  const fromEnv = process.env.FORMULA_PERF_HOME;
  if (fromEnv && fromEnv.trim() !== '') {
    return resolve(repoRoot, fromEnv);
  }
  return resolve(repoRoot, 'target', 'perf-home');
}

// Best-effort isolation: keep the desktop app from mutating a developer's real home directory.
const perfHome = resolvePerfHome();

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

// Benchmark environment knobs:
// - `FORMULA_DISABLE_STARTUP_UPDATE_CHECK=1` prevents the release updater from running a
//   background check/download on startup, which can add nondeterministic CPU/memory/network
//   activity and skew idle-memory benchmarks.
// - `FORMULA_STARTUP_METRICS=1` enables the Rust-side one-line startup metrics log we parse.
export function parseProcChildrenPids(content: string): number[] {
  const trimmed = content.trim();
  if (!trimmed) return [];
  return trimmed
    .split(/\s+/g)
    .map((x) => Number(x))
    .filter((n) => Number.isInteger(n) && n > 0);
}

export function parseProcStatusVmRssKb(content: string): number | null {
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

async function readProcExe(pid: number): Promise<string | null> {
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
  // NOTE: `/proc/<pid>/task/<tid>/children` is per-thread, not per-process. A multi-threaded
  // process can fork from any thread, so union children across all tasks to avoid missing
  // descendants (e.g. WebKit WebView helper processes).
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

async function getProcessRssBytesLinux(pid: number): Promise<number> {
  const status = await readUtf8(`/proc/${pid}/status`);
  if (!status) return 0;
  const kb = parseProcStatusVmRssKb(status);
  if (!kb) return 0;
  return kb * 1024;
}

async function getProcessTreeRssBytesLinux(rootPid: number): Promise<number> {
  const pids = await collectProcessTreePidsLinux(rootPid);
  let total = 0;
  for (const pid of pids) {
    total += await getProcessRssBytesLinux(pid);
  }
  return total;
}

async function sleep(ms: number): Promise<void> {
  await new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

async function findPidForExecutableLinux(
  rootPid: number,
  binPath: string,
  timeoutMs: number,
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
    const pids = await collectProcessTreePidsLinux(rootPid);
    for (const pid of pids) {
      const exe = await readProcExe(pid);
      if (!exe) continue;
      if (exe === binReal || exe === binResolved) return pid;
    }
    await sleep(50);
  }
  return null;
}

async function killPids(pids: number[], signal: NodeJS.Signals): Promise<void> {
  for (const pid of pids) {
    try {
      process.kill(pid, signal);
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === 'ESRCH') continue;
      // Ignore permission issues (shouldn't happen for our own children, but be defensive).
      if (code === 'EPERM') continue;
      throw err;
    }
  }
}

async function isPidAlive(pid: number): Promise<boolean> {
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ESRCH') return false;
    // If we don't have permissions to signal it, assume it's alive.
    if (code === 'EPERM') return true;
    return false;
  }
}

async function killProcessTreeLinux(rootPid: number, timeoutMs: number): Promise<void> {
  const initial = await collectProcessTreePidsLinux(rootPid);
  // Terminate children first so the root can't respawn them during shutdown.
  const ordered = initial.filter((p) => p !== rootPid).concat(rootPid);
  await killPids(ordered, 'SIGTERM');

  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    let anyAlive = false;
    for (const pid of ordered) {
      if (await isPidAlive(pid)) {
        anyAlive = true;
        break;
      }
    }
    if (!anyAlive) return;
    await sleep(50);
  }

  const stillAlive: number[] = [];
  for (const pid of ordered) {
    if (await isPidAlive(pid)) stillAlive.push(pid);
  }
  await killPids(stillAlive, 'SIGKILL');

  if (stillAlive.length > 0) {
    const killDeadline = Date.now() + 2000;
    while (Date.now() < killDeadline) {
      let anyAlive = false;
      for (const pid of stillAlive) {
        if (await isPidAlive(pid)) {
          anyAlive = true;
          break;
        }
      }
      if (!anyAlive) break;
      await sleep(50);
    }
  }
}

async function waitForTti(child: ChildProcess, timeoutMs: number): Promise<number> {
  return await new Promise<number>((resolvePromise, rejectPromise) => {
    let done = false;

    const deadline = setTimeout(() => {
      if (done) return;
      done = true;
      cleanup();
      rejectPromise(new Error(`Timed out after ${timeoutMs}ms waiting for [startup] tti_ms log line`));
    }, timeoutMs);

    const onLine = (line: string) => {
      if (done) return;
      const parsed = parseStartupMetricsLine(line);
      if (!parsed) return;
      done = true;
      cleanup();
      resolvePromise(parsed.ttiMs);
    };

    const onClose = (code: number | null, signal: NodeJS.Signals | null) => {
      if (done) return;
      done = true;
      cleanup();
      rejectPromise(
        new Error(`Desktop process exited before reporting TTI (code=${code}, signal=${signal})`),
      );
    };

    const onError = (err: Error) => {
      if (done) return;
      done = true;
      cleanup();
      rejectPromise(err);
    };

    const rlOut = createInterface({ input: child.stdout! });
    const rlErr = createInterface({ input: child.stderr! });
    rlOut.on('line', onLine);
    rlErr.on('line', onLine);
    // Use `close` (not `exit`) so stdout/stderr are fully drained before we decide whether we
    // captured the `[startup] ...` line. This avoids false negatives when the process exits
    // quickly after logging.
    child.on('close', onClose);
    child.on('error', onError);

    const cleanup = () => {
      clearTimeout(deadline);
      rlOut.close();
      rlErr.close();
      child.off('close', onClose);
      child.off('error', onError);
    };
  });
}

async function waitForExit(child: ChildProcess, timeoutMs: number): Promise<void> {
  if (child.exitCode !== null || child.signalCode !== null) return;

  await new Promise<void>((resolvePromise, rejectPromise) => {
    const deadline = setTimeout(() => {
      cleanup();
      rejectPromise(new Error(`Timed out after ${timeoutMs}ms waiting for desktop process tree to exit`));
    }, timeoutMs);

    const onClose = () => {
      cleanup();
      resolvePromise();
    };

    const cleanup = () => {
      clearTimeout(deadline);
      child.off('close', onClose);
    };

    child.on('close', onClose);
  });
}

function defangChild(child: ChildProcess): void {
  // Best-effort: prevent the parent Node process from hanging if the child refuses to exit.
  // This is a last resort after kill attempts/timeouts.
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
}

async function runOnce(
  binPath: string,
  timeoutMs: number,
  settleMs: number,
  profileDir: string,
): Promise<number> {
  const useXvfb = shouldUseXvfb();
  const xvfbPath = resolve(repoRoot, 'scripts/xvfb-run-safe.sh');
  const command = useXvfb ? 'bash' : binPath;
  const args = useXvfb ? [xvfbPath, binPath] : [];
  const dirs = resolveProfileDirs(profileDir);

  // Optionally, force a clean state between iterations to avoid cache pollution.
  if (process.env.FORMULA_DESKTOP_BENCH_RESET_HOME === '1') {
    const rootDir = parse(perfHome).root;
    if (perfHome === rootDir || perfHome === repoRoot) {
      throw new Error(`Refusing to reset unsafe desktop benchmark home dir: ${perfHome}`);
    }
    rmSync(dirs.home, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 });
  }

  mkdirSync(dirs.home, { recursive: true });
  mkdirSync(dirs.tmp, { recursive: true });
  mkdirSync(dirs.xdgConfig, { recursive: true });
  mkdirSync(dirs.xdgCache, { recursive: true });
  mkdirSync(dirs.xdgState, { recursive: true });
  mkdirSync(dirs.xdgData, { recursive: true });
  mkdirSync(dirs.appData, { recursive: true });
  mkdirSync(dirs.localAppData, { recursive: true });

  const child = spawn(command, args, {
    cwd: repoRoot,
    stdio: ['ignore', 'pipe', 'pipe'],
    env: {
      ...process.env,
      // Keep perf benchmarks stable/quiet by disabling the automatic startup update check.
      FORMULA_DISABLE_STARTUP_UPDATE_CHECK: '1',
      // Enable the Rust-side single-line log in release builds.
      FORMULA_STARTUP_METRICS: '1',
      // In case the app reads $HOME for config, keep per-run caches out of the real home dir.
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
    // On POSIX, start the app in its own process group so we can signal the entire tree.
    // Even though this benchmark currently runs on Linux only, keeping this consistent
    // with other perf runners prevents copy/paste drift.
    detached: process.platform !== 'win32',
    windowsHide: true,
  });

  if (!child.pid) {
    child.kill();
    throw new Error('Failed to spawn desktop process (missing pid)');
  }

  try {
    await waitForTti(child, timeoutMs);

    const resolvedPid = await findPidForExecutableLinux(child.pid, binPath, Math.min(2000, timeoutMs));
    if (!resolvedPid) {
      if (useXvfb) {
        throw new Error('Failed to resolve desktop PID under Xvfb wrapper for RSS sampling');
      }
      throw new Error('Failed to resolve desktop PID for RSS sampling');
    }
    const rootPid = resolvedPid;

    if (settleMs > 0) {
      await sleep(settleMs);
    }

    const rssBytes = await getProcessTreeRssBytesLinux(rootPid);
    if (rssBytes <= 0) {
      throw new Error('Failed to sample desktop RSS (process may have exited)');
    }
    return rssBytes / (1024 * 1024);
  } finally {
    // Ensure we clean up even on timeouts / crashes.
    try {
      if (process.platform === 'linux') {
        // Even if the wrapper process already exited, lingering WebView/WebKit helper processes can
        // remain in the same process group. Best-effort: signal the group first, then fall back to
        // enumerating the /proc process tree.
        terminateProcessTree(child, 'graceful');
        await killProcessTreeLinux(child.pid, 5000);
        // Last resort: if any processes are still alive in the original process group, force-kill them.
        terminateProcessTree(child, 'force');
      } else {
        terminateProcessTree(child, 'force');
      }
    } finally {
      await waitForExit(child, 5000).catch(() => {
        // If we failed to cleanly terminate, try a last-resort kill + detach the handles
        // so the parent process cannot hang indefinitely.
        terminateProcessTree(child, 'force');
        defangChild(child);
        throw new Error('Timed out waiting for desktop process tree to exit after sampling memory');
      });
    }
  }
}

function buildResult(name: string, values: number[], targetMb: number): BenchmarkResult {
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
    unit: 'mb',
    mean: avg,
    median: med,
    p95,
    p99,
    stdDev: sd,
    targetMs: targetMb,
    passed: p95 <= targetMb,
  };
}

export async function runDesktopMemoryBenchmarks(): Promise<BenchmarkResult[]> {
  if (process.env.FORMULA_RUN_DESKTOP_MEMORY_BENCH !== '1') {
    return [];
  }

  if (process.platform !== 'linux') {
    return [];
  }

  const runs = Math.max(1, Number(process.env.FORMULA_DESKTOP_MEMORY_RUNS ?? '10') || 10);
  const settleMs = Math.max(0, Number(process.env.FORMULA_DESKTOP_MEMORY_SETTLE_MS ?? '5000') || 5000);
  const timeoutMs = Math.max(
    1,
    Number(process.env.FORMULA_DESKTOP_MEMORY_TIMEOUT_MS ?? '30000') || 30000,
  );

  const targetRaw =
    process.env.FORMULA_DESKTOP_IDLE_RSS_TARGET_MB ?? process.env.FORMULA_DESKTOP_MEMORY_TARGET_MB ?? '100';
  const targetMb = Number(targetRaw) || 100;

  const binPath = process.env.FORMULA_DESKTOP_BIN
    ? resolve(process.env.FORMULA_DESKTOP_BIN)
    : defaultDesktopBinPath();

  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      `Desktop binary not found (bin=${String(binPath)}). Build it via (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build) and/or set FORMULA_DESKTOP_BIN.`,
    );
  }

  const profileRoot = resolve(perfHome, `desktop-memory-${Date.now()}-${process.pid}`);

  // eslint-disable-next-line no-console
  console.log(
    `[desktop-memory] idle RSS benchmark: runs=${runs} settleMs=${settleMs} timeoutMs=${timeoutMs} targetMb=${targetMb} profile=${profileRoot}`,
  );

  const values: number[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory] run ${i + 1}/${runs}...`);
    const rssMb = await runOnce(binPath, timeoutMs, settleMs, profileRoot);
    values.push(rssMb);
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory]   idleRssMb=${rssMb.toFixed(1)}mb`);
  }

  return [buildResult('desktop.memory.idle_rss_mb.p95', values, targetMb)];
}
