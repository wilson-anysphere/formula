import { resolve } from 'node:path';

import { repoRoot } from './desktopStartupRunnerShared.ts';

export type DesktopMemoryBenchEnv = {
  runs: number;
  timeoutMs: number;
  settleMs: number;
  targetMb: number;
  enforce: boolean;
  /**
   * Resolved desktop binary path from `FORMULA_DESKTOP_BIN`, or null when unset.
   *
   * Callers typically fall back to `defaultDesktopBinPath()` when this is null.
   */
  binPath: string | null;
};

function parseNumber(raw: string | undefined): number | null {
  if (raw === undefined) return null;
  const trimmed = raw.trim();
  if (!trimmed) return null;
  const n = Number(trimmed);
  if (!Number.isFinite(n)) return null;
  return n;
}

function parsePositiveNumber(raw: string | undefined): number | null {
  const n = parseNumber(raw);
  if (n === null || n <= 0) return null;
  return n;
}

function parseNonNegativeNumber(raw: string | undefined): number | null {
  const n = parseNumber(raw);
  if (n === null || n < 0) return null;
  return n;
}

/**
 * Resolve environment variables for the desktop idle memory benchmark.
 *
 * Centralizing this parsing avoids drift between:
 * - `desktopMemoryBench.ts` (integrated benchmark runner)
 * - `desktop-memory-runner.ts` (standalone CLI runner)
 */
export function resolveDesktopMemoryBenchEnv(options: {
  env?: NodeJS.ProcessEnv;
} = {}): DesktopMemoryBenchEnv {
  const env = options.env ?? process.env;

  const runs = parsePositiveNumber(env.FORMULA_DESKTOP_MEMORY_RUNS) ?? 10;
  const timeoutMs = parsePositiveNumber(env.FORMULA_DESKTOP_MEMORY_TIMEOUT_MS) ?? 20_000;
  const settleMs = parseNonNegativeNumber(env.FORMULA_DESKTOP_MEMORY_SETTLE_MS) ?? 5_000;

  const targetMb =
    parsePositiveNumber(env.FORMULA_DESKTOP_IDLE_RSS_TARGET_MB) ??
    parsePositiveNumber(env.FORMULA_DESKTOP_MEMORY_TARGET_MB) ??
    100;

  const enforce = env.FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH === '1';

  const rawBin = env.FORMULA_DESKTOP_BIN;
  // Mirror the startup benchmark behavior: resolve relative paths from the repo root, even if the
  // caller invokes the benchmark from a different working directory.
  const binPath = rawBin && rawBin.trim() !== '' ? resolve(repoRoot, rawBin) : null;

  return { runs, timeoutMs, settleMs, targetMb, enforce, binPath };
}
