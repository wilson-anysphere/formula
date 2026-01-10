import { performance } from 'node:perf_hooks';

export type BenchmarkFn = () => void | Promise<void>;

export interface BenchmarkOptions {
  iterations?: number;
  warmup?: number;
  /**
   * Absolute p95 threshold in milliseconds.
   *
   * Benchmarks fail if `p95 > targetMs`.
   */
  targetMs: number;
  unit?: 'ms';
}

export interface BenchmarkResult {
  name: string;
  iterations: number;
  warmup: number;
  unit: 'ms';

  mean: number;
  median: number;
  p95: number;
  p99: number;
  stdDev: number;

  targetMs: number;
  passed: boolean;
}

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
  const variance =
    values.reduce((sum, x) => sum + Math.pow(x - avg, 2), 0) / values.length;
  return Math.sqrt(variance);
}

export async function runBenchmark(
  name: string,
  fn: BenchmarkFn,
  options: BenchmarkOptions,
): Promise<BenchmarkResult> {
  const iterations = options.iterations ?? 50;
  const warmup = options.warmup ?? 10;
  const targetMs = options.targetMs;
  const unit: 'ms' = options.unit ?? 'ms';

  for (let i = 0; i < warmup; i++) {
    await fn();
  }

  const results: number[] = [];
  for (let i = 0; i < iterations; i++) {
    const start = performance.now();
    await fn();
    results.push(performance.now() - start);
  }

  const sorted = [...results].sort((a, b) => a - b);
  const avg = mean(sorted);
  const med = median(sorted);
  const p95 = percentile(sorted, 0.95);
  const p99 = percentile(sorted, 0.99);
  const sd = stdDev(sorted, avg);

  return {
    name,
    iterations,
    warmup,
    unit,
    mean: avg,
    median: med,
    p95,
    p99,
    stdDev: sd,
    targetMs,
    passed: p95 <= targetMs,
  };
}

export function formatMs(value: number): string {
  if (value >= 1000) return `${(value / 1000).toFixed(2)}s`;
  if (value >= 10) return `${value.toFixed(1)}ms`;
  return `${value.toFixed(3)}ms`;
}

