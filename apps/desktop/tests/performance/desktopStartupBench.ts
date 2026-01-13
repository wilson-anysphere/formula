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
    metrics.push(await runOnce({ binPath, timeoutMs, envOverrides: {} }));
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
