/**
 * Desktop idle RSS benchmark (Linux-only).
 *
 * Reproducibility + safety:
 * - The desktop process is spawned with user-data directories redirected under
 *   `target/perf-home` so the benchmark cannot read/write the real user profile.
 * - This keeps caches deterministic on CI runners (and avoids polluting developer machines).
 *
 * Environment isolation:
 * - All platforms: `HOME` + `USERPROFILE` => `target/perf-home`
 * - Linux: `XDG_CONFIG_HOME`, `XDG_CACHE_HOME`, `XDG_DATA_HOME` => `target/perf-home/xdg-*`
 * - Windows: `APPDATA`, `LOCALAPPDATA`, `TEMP`, `TMP` => `target/perf-home/*`
 * - macOS/Linux: `TMPDIR` => `target/perf-home/tmp`
 *
 * Reset behavior:
 * - Set `FORMULA_DESKTOP_BENCH_RESET_HOME=1` to delete `target/perf-home` before *each* iteration.
 */

import { spawn, type ChildProcess } from 'node:child_process';
import { existsSync, mkdirSync, realpathSync, rmSync } from 'node:fs';
import { readFile, readlink, readdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';

import { type BenchmarkResult } from './benchmark.ts';
import {
  defaultDesktopBinPath,
  mean,
  median,
  percentile,
  parseStartupLine as parseStartupMetricsLine,
  shouldUseXvfb,
  stdDev,
} from './desktopStartupUtil.ts';

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

// Best-effort isolation: keep the desktop app from mutating a developer's real home directory.
const perfHome = resolve(repoRoot, 'target', 'perf-home');
const perfTmp = resolve(perfHome, 'tmp');
const perfXdgConfig = resolve(perfHome, 'xdg-config');
const perfXdgCache = resolve(perfHome, 'xdg-cache');
const perfXdgState = resolve(perfHome, 'xdg-state');
const perfXdgData = resolve(perfHome, 'xdg-data');
const perfAppData = resolve(perfHome, 'AppData', 'Roaming');
const perfLocalAppData = resolve(perfHome, 'AppData', 'Local');

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

    const onExit = (code: number | null, signal: NodeJS.Signals | null) => {
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
    child.on('exit', onExit);
    child.on('error', onError);

    const cleanup = () => {
      clearTimeout(deadline);
      rlOut.close();
      rlErr.close();
      child.off('exit', onExit);
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

    const onExit = () => {
      cleanup();
      resolvePromise();
    };

    const cleanup = () => {
      clearTimeout(deadline);
      child.off('exit', onExit);
    };

    child.on('exit', onExit);
  });
}

async function runOnce(binPath: string, timeoutMs: number, settleMs: number): Promise<number> {
  const useXvfb = shouldUseXvfb();
  const xvfbPath = resolve(repoRoot, 'scripts/xvfb-run-safe.sh');
  const command = useXvfb ? 'bash' : binPath;
  const args = useXvfb ? [xvfbPath, binPath] : [];

  // Optionally, force a clean state between iterations to avoid cache pollution.
  if (process.env.FORMULA_DESKTOP_BENCH_RESET_HOME === '1') {
    rmSync(perfHome, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 });
  }

  mkdirSync(perfHome, { recursive: true });
  mkdirSync(perfTmp, { recursive: true });
  mkdirSync(perfXdgConfig, { recursive: true });
  mkdirSync(perfXdgCache, { recursive: true });
  mkdirSync(perfXdgState, { recursive: true });
  mkdirSync(perfXdgData, { recursive: true });
  mkdirSync(perfAppData, { recursive: true });
  mkdirSync(perfLocalAppData, { recursive: true });

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
      HOME: perfHome,
      USERPROFILE: perfHome,
      XDG_CONFIG_HOME: perfXdgConfig,
      XDG_CACHE_HOME: perfXdgCache,
      XDG_STATE_HOME: perfXdgState,
      XDG_DATA_HOME: perfXdgData,
      APPDATA: perfAppData,
      LOCALAPPDATA: perfLocalAppData,
      TMPDIR: perfTmp,
      TEMP: perfTmp,
      TMP: perfTmp,
    },
  });

  if (!child.pid) {
    child.kill();
    throw new Error('Failed to spawn desktop process (missing pid)');
  }

  try {
    await waitForTti(child, timeoutMs);

    const rootPid =
      (await findPidForExecutableLinux(child.pid, binPath, Math.min(2000, timeoutMs))) ?? child.pid;

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
        await killProcessTreeLinux(child.pid, 5000);
      } else {
        child.kill();
      }
    } finally {
      await waitForExit(child, 5000).catch(() => {
        // If we failed to cleanly terminate, try a last-resort SIGKILL.
        try {
          child.kill('SIGKILL');
        } catch {
          child.kill();
        }
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
    Number(process.env.FORMULA_DESKTOP_MEMORY_TIMEOUT_MS ?? '20000') || 20000,
  );

  const targetMb = Number(process.env.FORMULA_DESKTOP_IDLE_RSS_TARGET_MB ?? '100') || 100;

  const binPath = process.env.FORMULA_DESKTOP_BIN
    ? resolve(process.env.FORMULA_DESKTOP_BIN)
    : defaultDesktopBinPath();

  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      `Desktop binary not found (bin=${String(binPath)}). Build it via (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build) and/or set FORMULA_DESKTOP_BIN.`,
    );
  }

  const values: number[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory] run ${i + 1}/${runs}...`);
    values.push(await runOnce(binPath, timeoutMs, settleMs));
  }

  return [buildResult('desktop.memory.idle_rss_mb.p95', values, targetMb)];
}
