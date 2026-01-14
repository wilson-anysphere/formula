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
 * - Linux: `XDG_CONFIG_HOME`, `XDG_CACHE_HOME`, `XDG_STATE_HOME`, `XDG_DATA_HOME` => `${HOME}/xdg-*`
 * - Windows: `APPDATA`, `LOCALAPPDATA`, `TEMP`, `TMP` => `${HOME}/*`
 * - macOS/Linux: `TMPDIR` => `${HOME}/tmp`
 *
 * Reset behavior:
 * - Set `FORMULA_DESKTOP_BENCH_RESET_HOME=1` to delete the benchmark profile directory (HOME)
 *   before *each* iteration.
 */

import { existsSync } from 'node:fs';
import { resolve } from 'node:path';

import { buildBenchmarkResultFromValues, type BenchmarkResult } from './benchmark.ts';
import {
  defaultDesktopBinPath,
  findPidForExecutableLinux,
  formatPerfPath,
  getProcessTreeRssBytesLinux,
  sleep,
  resolvePerfHome,
  runOnce as runDesktopOnce,
} from './desktopStartupUtil.ts';

async function sampleIdleRssMbLinux(options: {
  binPath: string;
  timeoutMs: number;
  settleMs: number;
  profileDir: string;
}): Promise<number> {
  const { binPath, timeoutMs, settleMs, profileDir } = options;

  let sampledRssMb: number | null = null;
  let sampleError: Error | null = null;

  await runDesktopOnce({
    binPath,
    timeoutMs,
    profileDir,
    afterCapture: async (child, _metrics, signal) => {
      try {
        if (settleMs > 0) {
          await sleep(settleMs, signal);
        }

        const wrapperPid = child.pid;
        if (!wrapperPid || wrapperPid <= 0) {
          throw new Error('Failed to spawn desktop process (missing pid)');
        }

        const resolvedPid = await findPidForExecutableLinux(
          wrapperPid,
          binPath,
          Math.min(2000, timeoutMs),
          signal,
        );
        if (!resolvedPid) {
          throw new Error('Failed to resolve desktop PID for RSS sampling');
        }

        const rssBytes = await getProcessTreeRssBytesLinux(resolvedPid);
        if (rssBytes <= 0) {
          throw new Error('Failed to sample desktop RSS (process may have exited)');
        }

        sampledRssMb = rssBytes / (1024 * 1024);
      } catch (err) {
        sampleError = err instanceof Error ? err : new Error(String(err));
      }
    },
    afterCaptureTimeoutMs: settleMs + 5000,
  });

  if (sampleError) throw sampleError;
  if (sampledRssMb == null) throw new Error('Failed to sample desktop RSS');
  return sampledRssMb;
}

export async function runDesktopMemoryBenchmarks(): Promise<BenchmarkResult[]> {
  if (process.env.FORMULA_RUN_DESKTOP_MEMORY_BENCH !== '1') {
    return [];
  }

  if (process.platform !== 'linux') {
    return [];
  }

  const runs = Math.max(1, Number(process.env.FORMULA_DESKTOP_MEMORY_RUNS ?? '10') || 10);
  // Allow explicitly setting `FORMULA_DESKTOP_MEMORY_SETTLE_MS=0` to sample immediately. Treat
  // unset/blank/invalid values as the default.
  const settleRaw = process.env.FORMULA_DESKTOP_MEMORY_SETTLE_MS;
  const settleParsed = settleRaw && settleRaw.trim() !== '' ? Number(settleRaw) : 5000;
  const settleMs = Number.isFinite(settleParsed) ? Math.max(0, settleParsed) : 5000;
  const timeoutMs = Math.max(
    1,
    Number(process.env.FORMULA_DESKTOP_MEMORY_TIMEOUT_MS ?? '20000') || 20000,
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

  // Best-effort isolation: keep the desktop app from mutating a developer's real home directory.
  const perfHome = resolvePerfHome();
  const profileRoot = resolve(perfHome, `desktop-memory-${Date.now()}-${process.pid}`);

  // eslint-disable-next-line no-console
  console.log(
    `[desktop-memory] idle RSS benchmark: runs=${runs} settleMs=${settleMs} timeoutMs=${timeoutMs} targetMb=${targetMb} profile=${formatPerfPath(profileRoot)}`,
  );

  const values: number[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory] run ${i + 1}/${runs}...`);
    const rssMb = await sampleIdleRssMbLinux({ binPath, timeoutMs, settleMs, profileDir: profileRoot });
    values.push(rssMb);
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory]   idleRssMb=${rssMb.toFixed(1)}mb`);
  }

  const p95 = buildBenchmarkResultFromValues('desktop.memory.idle_rss_mb.p95', values, targetMb, 'mb');
  const p50: BenchmarkResult = {
    ...p95,
    name: 'desktop.memory.idle_rss_mb.p50',
    // Informational: CI gating is based on the p95 budget.
    targetMs: undefined,
    passed: true,
  };
  return [p95, p50];
}
