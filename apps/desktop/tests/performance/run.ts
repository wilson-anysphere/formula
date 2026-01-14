import { spawnSync } from 'node:child_process';
import {
  accessSync,
  constants as fsConstants,
  existsSync,
  mkdirSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

import { formatMb, formatMs, runBenchmark, type BenchmarkResult } from './benchmark.ts';
import { createCollaborationBenchmarks } from './benchmarks/collaboration.bench.ts';
import { createRenderBenchmarks } from './benchmarks/render.bench.ts';
import { createSharedGridRendererBenchmarks } from './benchmarks/sharedGridRenderer.bench.ts';
import { createStartupBenchmarks } from './benchmarks/startup.bench.ts';
import { repoRoot } from './desktopStartupUtil.ts';
import { runDesktopStartupBenchmarks } from './desktopStartupBench.ts';
import { runDesktopMemoryBenchmarks } from './desktopMemoryBench.ts';

type DetailedReport = {
  generatedAt: string;
  benchmarks: BenchmarkResult[];
};

type ActionBenchmark = { name: string; unit: BenchmarkResult['unit']; value: number };

function parseArgs(argv: string[]): { output: string; details: string } {
  const defaults = {
    output: 'benchmark-results.json',
    details: 'benchmark-details.json',
  };

  const args = [...argv];
  while (args.length > 0) {
    const arg = args.shift();
    if (arg === '--output' && args[0]) defaults.output = args.shift()!;
    else if (arg === '--details' && args[0]) defaults.details = args.shift()!;
  }

  return {
    output: defaults.output,
    details: defaults.details,
  };
}

function printSummary(results: BenchmarkResult[]): void {
  const longestName = Math.max(...results.map((r) => r.name.length));

  const formatValue = (value: number, unit: BenchmarkResult['unit']): string => {
    if (unit === 'mb') return formatMb(value);
    return formatMs(value);
  };

  for (const r of results) {
    const status = r.passed ? 'PASS' : 'FAIL';
    const name = r.name.padEnd(longestName);
    const p95 = formatValue(r.p95, r.unit).padStart(10);
    const target = (r.targetMs === undefined ? 'unset' : formatValue(r.targetMs, r.unit)).padStart(10);
    // eslint-disable-next-line no-console
    console.log(`${status}  ${name}  p95=${p95}  target=${target}`);
  }
}

function parseOptionalTargetMb(name: string): number | null {
  const raw = process.env[name];
  if (raw === undefined) return null;
  if (raw.trim() === '') return null;
  const val = Number(raw);
  if (!Number.isFinite(val) || val <= 0) {
    throw new Error(`Invalid ${name}=${JSON.stringify(raw)} (expected a number > 0)`);
  }
  return val;
}

function bytesToMb(bytes: number): number {
  // Use decimal MB so thresholds match common tooling / release size budgets.
  return bytes / 1_000_000;
}

function round(value: number, decimals: number): number {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

function directorySizeBytes(dirPath: string): number {
  let total = 0;
  const entries = readdirSync(dirPath, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = resolve(dirPath, entry.name);
    if (entry.isDirectory()) {
      total += directorySizeBytes(fullPath);
      continue;
    }
    if (entry.isFile()) {
      total += statSync(fullPath).size;
      continue;
    }
    // Ignore symlinks and special files to keep the metric stable.
  }
  return total;
}

function buildScalarResult(
  name: string,
  unit: BenchmarkResult['unit'],
  value: number,
  target: number | null,
): BenchmarkResult {
  const targetMs = target ?? undefined;
  return {
    name,
    iterations: 1,
    warmup: 0,
    unit,
    mean: value,
    median: value,
    p95: value,
    p99: value,
    stdDev: 0,
    targetMs,
    passed: targetMs === undefined ? true : value <= targetMs,
  };
}

function tryCreateDistTarGzBytes(distDir: string): number | null {
  const outDir = resolve(repoRoot, 'target', 'benchmark-artifacts');
  mkdirSync(outDir, { recursive: true });
  const outFile = resolve(outDir, 'desktop-dist.tar.gz');
  rmSync(outFile, { force: true });

  const proc = spawnSync('tar', ['-czf', outFile, '-C', distDir, '.'], {
    encoding: 'utf8',
    stdio: ['ignore', 'ignore', 'pipe'],
    cwd: repoRoot,
  });

  if (proc.error) return null;
  if (proc.status !== 0) {
    // If tar isn't available (or fails for any reason), treat this as an optional metric.
    rmSync(outFile, { force: true });
    return null;
  }

  try {
    const size = statSync(outFile).size;
    rmSync(outFile, { force: true });
    return size;
  } catch {
    rmSync(outFile, { force: true });
    return null;
  }
}

function collectOptionalSizeMetrics(): BenchmarkResult[] {
  const metrics: BenchmarkResult[] = [];

  // Optional enforcement (max sizes in decimal MB):
  // - FORMULA_DESKTOP_BINARY_SIZE_TARGET_MB
  // - FORMULA_DESKTOP_DIST_SIZE_TARGET_MB
  // - FORMULA_DESKTOP_DIST_GZIP_SIZE_TARGET_MB

  const exe = process.platform === 'win32' ? 'formula-desktop.exe' : 'formula-desktop';
  const binaryPath = resolve(repoRoot, 'target', 'release', exe);
  if (existsSync(binaryPath)) {
    const stats = statSync(binaryPath);
    if (stats.isFile()) {
      const bytes = stats.size;
      const mb = round(bytesToMb(bytes), 3);
      const targetMb = parseOptionalTargetMb('FORMULA_DESKTOP_BINARY_SIZE_TARGET_MB');
      metrics.push(buildScalarResult('desktop.size.binary_mb', 'mb', mb, targetMb));
    }
  }

  const distDir = resolve(repoRoot, 'apps', 'desktop', 'dist');
  if (existsSync(distDir)) {
    const stats = statSync(distDir);
    if (stats.isDirectory()) {
      const bytes = directorySizeBytes(distDir);
      const mb = round(bytesToMb(bytes), 3);
      const targetMb = parseOptionalTargetMb('FORMULA_DESKTOP_DIST_SIZE_TARGET_MB');
      metrics.push(buildScalarResult('desktop.size.dist_mb', 'mb', mb, targetMb));

      const gzBytes = tryCreateDistTarGzBytes(distDir);
      if (gzBytes !== null) {
        const gzMb = round(bytesToMb(gzBytes), 3);
        const gzTargetMb = parseOptionalTargetMb('FORMULA_DESKTOP_DIST_GZIP_SIZE_TARGET_MB');
        metrics.push(buildScalarResult('desktop.size.dist_gzip_mb', 'mb', gzMb, gzTargetMb));
      }
    }
  }

  return metrics;
}

function runRustBenchmarks(): BenchmarkResult[] {
  const cargoArgs = [
    '-q',
    '-p',
    'formula-engine',
    '--bin',
    'perf_bench',
    '--release',
  ];

  const defaultGlobalCargoHome = resolve(homedir(), '.cargo');
  const envCargoHome = process.env.CARGO_HOME;
  const normalizedEnvCargoHome = envCargoHome ? resolve(envCargoHome) : null;
  const cargoHome =
    !envCargoHome ||
    (!process.env.CI &&
      !process.env.FORMULA_ALLOW_GLOBAL_CARGO_HOME &&
      normalizedEnvCargoHome === defaultGlobalCargoHome)
      ? resolve(repoRoot, 'target', 'cargo-home')
      : envCargoHome;
  mkdirSync(cargoHome, { recursive: true });

  const safeRun = resolve(repoRoot, 'scripts/safe-cargo-run.sh');
  let canUseSafeRun = process.platform !== 'win32' && existsSync(safeRun);
  if (canUseSafeRun) {
    try {
      accessSync(safeRun, fsConstants.X_OK);
    } catch {
      canUseSafeRun = false;
    }
  }
  const command = canUseSafeRun ? safeRun : 'cargo';
  const args = canUseSafeRun ? cargoArgs : ['run', ...cargoArgs];

  const proc = spawnSync(command, args, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    cwd: repoRoot,
    env: { ...process.env, CARGO_HOME: cargoHome },
  });

  if (proc.error) throw proc.error;
  if (proc.status !== 0) {
    throw new Error(`cargo perf_bench failed (exit ${proc.status}):\n${proc.stderr}`);
  }

  const parsed = JSON.parse(proc.stdout) as { benchmarks: BenchmarkResult[] };
  return parsed.benchmarks;
}

async function main(): Promise<void> {
  const { output, details } = parseArgs(process.argv.slice(2));

  const benchmarks = [
    ...createStartupBenchmarks(),
    ...createRenderBenchmarks(),
    ...createCollaborationBenchmarks(),
  ];

  if (process.env.FORMULA_BENCH_DOCUMENT_CELL_PROVIDER === '1') {
    const { createDocumentCellProviderCacheKeyBenchmarks } = await import(
      './benchmarks/documentCellProviderCacheKey.bench.ts'
    );
    benchmarks.push(...createDocumentCellProviderCacheKeyBenchmarks());
  }

  // CanvasGridRenderer benchmarks install a global JSDOM + canvas mocks. Run them
  // last so other (pure Node) benchmarks aren't affected by the DOM globals.
  benchmarks.push(...createSharedGridRendererBenchmarks());

  const results: BenchmarkResult[] = [];
  for (const bench of benchmarks) {
    results.push(
      await runBenchmark(bench.name, bench.fn, {
        iterations: bench.iterations,
        warmup: bench.warmup,
        targetMs: bench.targetMs,
        clock: bench.clock,
      }),
    );
  }

  // Rust engine microbenchmarks (parse/eval/recalc).
  results.push(...runRustBenchmarks());

  // Optional: real desktop startup (Tauri binary) timings.
  // Supports cold vs warm profiles via `FORMULA_DESKTOP_STARTUP_MODE=cold|warm`.
  // This is gated because it requires a built binary + a usable display environment.
  results.push(...(await runDesktopStartupBenchmarks()));
  results.push(...(await runDesktopMemoryBenchmarks()));

  // Optional: size metrics for the release binary and/or built frontend assets.
  // Skip quietly when artifacts aren't present (e.g. perf.yml does not build them).
  results.push(...collectOptionalSizeMetrics());

  results.sort((a, b) => a.name.localeCompare(b.name));

  printSummary(results);

  const report: DetailedReport = {
    generatedAt: new Date().toISOString(),
    benchmarks: results,
  };

  const actionResults: ActionBenchmark[] = results.map((r) => ({
    name: r.name,
    unit: r.unit,
    value: r.p95,
  }));

  writeFileSync(resolve(repoRoot, details), JSON.stringify(report, null, 2));
  writeFileSync(resolve(repoRoot, output), JSON.stringify(actionResults, null, 2));

  const failed = results.filter((r) => !r.passed);
  if (failed.length > 0) {
    // eslint-disable-next-line no-console
    console.error(
      `\nPerformance regression: ${failed.length} benchmark(s) exceeded p95 targets.`,
    );
    process.exitCode = 1;
  }
}

await main();
