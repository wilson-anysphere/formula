import { resolve } from 'node:path';

import { beforeEach, describe, expect, it, vi } from 'vitest';

import { runDesktopStartupBenchmarks } from './desktopStartupBench.ts';
import { repoRoot } from './desktopStartupUtil.ts';

function withTempEnv(vars: Record<string, string | undefined>, fn: () => Promise<void>): Promise<void> {
  const prev: Record<string, string | undefined> = {};
  for (const k of Object.keys(vars)) {
    prev[k] = process.env[k];
  }
  for (const [k, v] of Object.entries(vars)) {
    if (v === undefined) delete process.env[k];
    else process.env[k] = v;
  }
  return fn().finally(() => {
    for (const [k, v] of Object.entries(prev)) {
      if (v === undefined) delete process.env[k];
      else process.env[k] = v;
    }
  });
}

function findResult(results: Awaited<ReturnType<typeof runDesktopStartupBenchmarks>>, name: string) {
  return results.find((r) => r.name === name) ?? null;
}

describe('desktopStartupBench cold vs warm', () => {
  beforeEach(() => {
    // These benchmarks spawn a child process and rely on real timers for timeouts.
    // Some test suites use fake timers and can leak them across files when a test aborts early.
    // Force real timers here to keep the benchmark harness stable in full-suite runs.
    vi.useRealTimers();
  });

  it('emits cold metrics and does not reuse the profile between runs', async () => {
    const modulePath = resolve(repoRoot, 'apps/desktop/tests/performance/fixtures/fakeDesktopStartupModule.cjs');
    await withTempEnv(
      {
        // Ensure the benchmark runs.
        FORMULA_RUN_DESKTOP_STARTUP_BENCH: '1',
        // Use Node as our fake “desktop binary” and load a module that prints [startup] metrics.
        FORMULA_DESKTOP_BIN: process.execPath,
        NODE_OPTIONS: `--require ${modulePath}`,

        // Keep the benchmark fast for unit tests.
        FORMULA_DESKTOP_STARTUP_RUNS: '3',
        // Under full vitest runs (many tests + child process churn), spawning the fake desktop
        // process can occasionally take several seconds on CI. Keep the overall benchmark fast
        // while allowing some headroom to avoid flakes.
        FORMULA_DESKTOP_STARTUP_TIMEOUT_MS: '15000',
        FORMULA_DESKTOP_RSS_IDLE_DELAY_MS: '0',
        FORMULA_DESKTOP_STARTUP_MODE: 'cold',
        FORMULA_DESKTOP_STARTUP_BENCH_KIND: 'full',

        // Avoid requiring Xvfb for unit tests.
        CI: undefined,
        DISPLAY: ':99',
      },
      async () => {
        const results = await runDesktopStartupBenchmarks();

        const coldWindow = findResult(results, 'desktop.startup.cold.window_visible_ms.p95');
        const coldTti = findResult(results, 'desktop.startup.cold.tti_ms.p95');
        const aliasWindow = findResult(results, 'desktop.startup.window_visible_ms.p95');
        const aliasTti = findResult(results, 'desktop.startup.tti_ms.p95');

        expect(coldWindow).not.toBeNull();
        expect(coldTti).not.toBeNull();
        expect(aliasWindow).not.toBeNull();
        expect(aliasTti).not.toBeNull();

        // Our fake module reports `window_visible_ms=100` for cold launches. If the benchmark
        // accidentally reuses the same profile without resetting, later runs would report the
        // warm value (10) and the median would drop below 100.
        expect(coldWindow!.median).toBe(100);
        expect(coldWindow!.p95).toBe(100);
        expect(coldTti!.median).toBe(400);
        expect(coldTti!.p95).toBe(400);

        // Aliases should match the cold results.
        expect(aliasWindow!.p95).toBe(coldWindow!.p95);
        expect(aliasTti!.p95).toBe(coldTti!.p95);

        expect(findResult(results, 'desktop.startup.warm.window_visible_ms.p95')).toBeNull();
      },
    );
  });

  it('emits warm metrics and excludes the warmup run from the measured results', async () => {
    const modulePath = resolve(repoRoot, 'apps/desktop/tests/performance/fixtures/fakeDesktopStartupModule.cjs');
    await withTempEnv(
      {
        FORMULA_RUN_DESKTOP_STARTUP_BENCH: '1',
        FORMULA_DESKTOP_BIN: process.execPath,
        NODE_OPTIONS: `--require ${modulePath}`,

        FORMULA_DESKTOP_STARTUP_RUNS: '3',
        // This test can run under heavy parallel load in the monorepo vitest suite; give the
        // child process ample time to emit the `[startup]` metrics.
        FORMULA_DESKTOP_STARTUP_TIMEOUT_MS: '15000',
        FORMULA_DESKTOP_RSS_IDLE_DELAY_MS: '0',
        FORMULA_DESKTOP_STARTUP_MODE: 'warm',
        FORMULA_DESKTOP_STARTUP_BENCH_KIND: 'full',

        CI: undefined,
        DISPLAY: ':99',
      },
      async () => {
        const results = await runDesktopStartupBenchmarks();

        const warmWindow = findResult(results, 'desktop.startup.warm.window_visible_ms.p95');
        const warmTti = findResult(results, 'desktop.startup.warm.tti_ms.p95');
        expect(warmWindow).not.toBeNull();
        expect(warmTti).not.toBeNull();

        // The fake module reports `window_visible_ms=100` for a cold profile and 10 for warm.
        // Warm mode performs 1 warmup launch (cold) and then measures the subsequent warm runs.
        // If the warmup was mistakenly included, p95 would be 100.
        expect(warmWindow!.p95).toBe(10);
        expect(warmWindow!.median).toBe(10);
        expect(warmTti!.p95).toBe(40);
        expect(warmTti!.median).toBe(40);

        // Warm mode does not emit legacy unscoped aliases.
        expect(findResult(results, 'desktop.startup.window_visible_ms.p95')).toBeNull();
        expect(findResult(results, 'desktop.startup.tti_ms.p95')).toBeNull();
      },
    );
  });
});
