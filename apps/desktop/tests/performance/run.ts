import { spawnSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { formatMs, runBenchmark, type BenchmarkResult } from './benchmark.ts';
import { createCollaborationBenchmarks } from './benchmarks/collaboration.bench.ts';
import { createRenderBenchmarks } from './benchmarks/render.bench.ts';
import { createStartupBenchmarks } from './benchmarks/startup.bench.ts';

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

type DetailedReport = {
  generatedAt: string;
  benchmarks: BenchmarkResult[];
};

type ActionBenchmark = { name: string; unit: 'ms'; value: number };

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

  for (const r of results) {
    const status = r.passed ? 'PASS' : 'FAIL';
    const name = r.name.padEnd(longestName);
    const p95 = formatMs(r.p95).padStart(10);
    const target = formatMs(r.targetMs).padStart(10);
    // eslint-disable-next-line no-console
    console.log(`${status}  ${name}  p95=${p95}  target=${target}`);
  }
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

  const proc = spawnSync('bash', [resolve(repoRoot, 'scripts/safe-cargo-run.sh'), ...cargoArgs], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    cwd: repoRoot,
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

  const results: BenchmarkResult[] = [];
  for (const bench of benchmarks) {
    results.push(
      await runBenchmark(bench.name, bench.fn, {
        iterations: bench.iterations,
        warmup: bench.warmup,
        targetMs: bench.targetMs,
      }),
    );
  }

  // Rust engine microbenchmarks (parse/eval/recalc).
  results.push(...runRustBenchmarks());

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
