/**
 * Micro-benchmark for TabCompletionEngine.
 *
 * Goal: keep tab completion latency comfortably below ~100ms p95 for common
 * scenarios, and catch accidental performance regressions.
 *
 * Run:
 *   node packages/ai-completion/bench/tabCompletionEngine.bench.mjs
 *   pnpm bench:tab-completion
 *   pnpm -C packages/ai-completion bench:tab-completion
 *
 * Passing options through pnpm:
 *   pnpm bench:tab-completion -- --runs 200 --warmup 50
 *
 * CI:
 *   The script automatically enforces the latency budget when `CI=1` or when
 *   `--ci` is passed.
 */
import { writeFileSync } from "node:fs";
import { TabCompletionEngine } from "../src/tabCompletionEngine.js";

// Keep the benchmark fast (<5s) even on slower machines.
const DEFAULT_RUNS = 50;
const DEFAULT_WARMUP = 10;
const DEFAULT_BUDGET_MS = 100;

function parseArgs(argv) {
  /** @type {{runs:number, warmup:number, budgetMs:number, ci:boolean, output:string|null, details:string|null, scenarios:string[]}} */
  const out = {
    runs: parsePositiveInt(process.env.BENCH_RUNS, DEFAULT_RUNS),
    warmup: parsePositiveInt(process.env.BENCH_WARMUP, DEFAULT_WARMUP),
    budgetMs: parsePositiveInt(process.env.BENCH_BUDGET_MS, DEFAULT_BUDGET_MS),
    ci: isTruthy(process.env.CI),
    output: null,
    details: null,
    scenarios: [],
  };

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    // Some package managers (e.g. `pnpm run <script> -- <args>`) will forward a
    // literal `--` separator to the underlying script. Treat it as a no-op so
    // users can invoke the benchmark with either:
    //   pnpm bench:tab-completion --runs 200
    // or
    //   pnpm bench:tab-completion -- --runs 200
    if (arg === "--") continue;
    if (arg === "--runs") {
      out.runs = parsePositiveInt(argv[++i], out.runs);
      continue;
    }
    if (arg === "--warmup") {
      out.warmup = parsePositiveInt(argv[++i], out.warmup);
      continue;
    }
    if (arg === "--budget-ms") {
      out.budgetMs = parsePositiveInt(argv[++i], out.budgetMs);
      continue;
    }
    if (arg === "--scenario") {
      const value = argv[++i];
      if (typeof value === "string" && value.length > 0) out.scenarios.push(value);
      continue;
    }
    if (arg === "--output") {
      out.output = argv[++i] ?? out.output;
      continue;
    }
    if (arg === "--details") {
      out.details = argv[++i] ?? out.details;
      continue;
    }
    if (arg === "--ci") {
      out.ci = true;
      continue;
    }
    if (arg === "--help" || arg === "-h") {
      printHelp();
      process.exit(0);
    }
    throw new Error(`Unknown arg: ${arg}`);
  }

  return out;
}

function printHelp() {
  console.log(`TabCompletionEngine benchmark

Usage:
  node packages/ai-completion/bench/tabCompletionEngine.bench.mjs [options]

Options:
  --runs <n>        Measured iterations per scenario (default: ${DEFAULT_RUNS})
  --warmup <n>      Warmup iterations per scenario (default: ${DEFAULT_WARMUP})
  --budget-ms <n>   CI p95 budget per scenario (default: ${DEFAULT_BUDGET_MS})
  --scenario <pat>  Run only scenarios whose name matches <pat> (repeatable)
  --output <path>   Write action-style JSON results (p95 ms) to a file
  --details <path>  Write detailed JSON results (p50/p95/mean/min/max) to a file
  --ci              Enforce the budget (also enabled automatically when CI=1)
`);
}

function isTruthy(value) {
  if (value === undefined || value === null) return false;
  const str = String(value).toLowerCase();
  return str === "1" || str === "true" || str === "yes";
}

function parsePositiveInt(value, fallback) {
  const n = Number(value);
  if (!Number.isFinite(n) || !Number.isInteger(n) || n <= 0) return fallback;
  return n;
}

function fmtMs(ms) {
  return `${ms.toFixed(3)}ms`;
}

function memUsage() {
  const { heapUsed } = process.memoryUsage();
  return `${(heapUsed / 1024 / 1024).toFixed(1)}MB`;
}

function quantile(sortedValues, q) {
  if (sortedValues.length === 0) return NaN;
  if (q <= 0) return sortedValues[0];
  if (q >= 1) return sortedValues[sortedValues.length - 1];

  const pos = (sortedValues.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  const next = sortedValues[base + 1];
  return next === undefined ? sortedValues[base] : sortedValues[base] + rest * (next - sortedValues[base]);
}

function summarize(samples) {
  const sorted = [...samples].sort((a, b) => a - b);
  const sum = samples.reduce((acc, v) => acc + v, 0);
  return {
    min: sorted[0],
    max: sorted[sorted.length - 1],
    mean: sum / samples.length,
    p50: quantile(sorted, 0.5),
    p95: quantile(sorted, 0.95),
  };
}

async function runScenario({ name, warmup, runs, fn }) {
  // Warm up JIT and one-time lazily computed things (e.g. FunctionRegistry sort).
  for (let i = 0; i < warmup; i++) {
    // Ensure work isn't optimized away.
    const out = await fn(i);
    if (Array.isArray(out) && out.length === 0) {
      // keep the value "live" even if empty
      void out;
    }
  }

  const samples = new Array(runs);
  let checksum = 0;
  for (let i = 0; i < runs; i++) {
    const start = process.hrtime.bigint();
    const out = await fn(i);
    const end = process.hrtime.bigint();
    samples[i] = Number(end - start) / 1e6;
    checksum += Array.isArray(out) ? out.length : 0;
  }

  return { name, ...summarize(samples), checksum };
}

async function main() {
  const { runs, warmup, budgetMs, ci, output, details, scenarios: scenarioFilters } = parseArgs(
    process.argv.slice(2)
  );

  // Disable caching so the benchmark exercises the full "cold path" per run.
  // The "per-keystroke" UX typically involves distinct inputs, so cache hits are
  // less common than steady-state profiling might suggest.
  const noCache = { get: () => undefined, set: () => {} };

  const engine = new TabCompletionEngine({ cache: noCache });

  const emptyCells = { getCellValue: () => null };

  // 10k populated values in column A.
  const filledRows = 10_000;
  const colA = new Float64Array(filledRows);
  for (let i = 0; i < filledRows; i++) colA[i] = i + 1;
  const tenKColCells = {
    getCellValue(row, col) {
      if (col !== 0) return null;
      if (!Number.isInteger(row) || row < 0 || row >= filledRows) return null;
      return colA[row];
    },
  };

  const patternColumn = new Array(200);
  for (let i = 0; i < patternColumn.length; i++) {
    if (i % 3 === 0) patternColumn[i] = "San Francisco";
    else if (i % 10 === 0) patternColumn[i] = "San Jose";
    else patternColumn[i] = null;
  }
  const patternCells = {
    getCellValue(row, col) {
      if (col !== 0) return null;
      if (!Number.isInteger(row) || row < 0) return null;
      if (row >= patternColumn.length) return null;
      return patternColumn[row] ?? null;
    },
  };

  const fnInput = "=VLO";
  const fnContext = {
    currentInput: fnInput,
    cursorPosition: fnInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: emptyCells,
  };

  const rangeInput = "=SUM(A";
  const rangeContext = {
    currentInput: rangeInput,
    cursorPosition: rangeInput.length,
    // Row 10001 (0-based 10000), below 10k populated cells in column A.
    cellRef: { row: filledRows, col: 0 },
    surroundingCells: tenKColCells,
  };

  const patternInput = "San";
  const patternContext = {
    currentInput: patternInput,
    cursorPosition: patternInput.length,
    cellRef: { row: 100, col: 0 },
    surroundingCells: patternCells,
  };

  // Sanity checks: make sure each scenario actually exercises the intended path.
  {
    const s1 = await engine.getSuggestions(fnContext);
    if (!s1.some((s) => typeof s?.text === "string" && s.text.toUpperCase().includes("VLOOKUP("))) {
      throw new Error(`Sanity check failed: function-name completion did not include VLOOKUP(. Got: ${s1.map((s) => s.text).join(", ")}`);
    }
    const s2 = await engine.getSuggestions(rangeContext);
    if (!s2.some((s) => s?.type === "range" && typeof s?.text === "string" && s.text.startsWith("=SUM("))) {
      throw new Error(`Sanity check failed: range completion did not include a SUM range suggestion. Got: ${s2.map((s) => s.text).join(", ")}`);
    }
    const s3 = await engine.getSuggestions(patternContext);
    if (!s3.some((s) => s?.type === "value" && typeof s?.text === "string")) {
      throw new Error(`Sanity check failed: pattern completion did not produce any value suggestions. Got: ${s3.map((s) => s.text).join(", ")}`);
    }
  }

  console.log("TabCompletionEngine benchmark");
  console.log(`Node ${process.version}`);
  console.log(`runs=${runs} warmup=${warmup} heap=${memUsage()}`);
  if (ci) console.log(`CI budget: p95 <= ${fmtMs(budgetMs)}`);
  console.log("");

  // Note: scenario names are used as stable identifiers in CI/perf tracking
  // (e.g. benchmark-action history). Avoid renaming them unless you are okay
  // with resetting the time-series.
  const scenarios = [
    {
      name: "Function-name completion (=VLO)",
      fn: () => engine.getSuggestions(fnContext),
    },
    {
      name: "Range completion (=SUM(A with 10k populated rows in column A)",
      fn: () => engine.getSuggestions(rangeContext),
    },
    {
      name: "Pattern completion (repeated strings, non-formula input)",
      fn: () => engine.getSuggestions(patternContext),
    },
  ];

  const selectedScenarios =
    scenarioFilters.length === 0
      ? scenarios
      : scenarios.filter((s) => scenarioFilters.some((f) => includesIgnoreCase(s.name, f)));

  if (selectedScenarios.length === 0) {
    throw new Error(
      `No scenarios matched filters: ${scenarioFilters.join(
        ", "
      )}. Available: ${scenarios.map((s) => s.name).join(" | ")}`
    );
  }

  /** @type {{name:string, p95:number}[]} */
  const failures = [];
  /** @type {any[]} */
  const detailedResults = [];
  for (const scenario of selectedScenarios) {
    const result = await runScenario({
      name: scenario.name,
      warmup,
      runs,
      fn: scenario.fn,
    });
    detailedResults.push(result);
    console.log(
      `${result.name}\n  p50 ${fmtMs(result.p50)}  p95 ${fmtMs(result.p95)}  mean ${fmtMs(result.mean)}  min ${fmtMs(result.min)}  max ${fmtMs(result.max)}`
    );
    console.log("");

    if (ci && result.p95 > budgetMs) {
      failures.push({ name: result.name, p95: result.p95 });
    }
  }

  if (failures.length > 0) {
    console.error(`TabCompletionEngine benchmark failed budget p95 <= ${fmtMs(budgetMs)}:`);
    for (const f of failures) console.error(`  ${f.name}: p95=${fmtMs(f.p95)}`);
    process.exitCode = 1;
  }

  if (output) {
    /** @type {{name:string, unit:"ms", value:number}[]} */
    const actionResults = detailedResults.map((r) => ({ name: r.name, unit: "ms", value: r.p95 }));
    writeFileSync(output, JSON.stringify(actionResults, null, 2));
  }

  if (details) {
    writeFileSync(
      details,
      JSON.stringify(
        {
          generatedAt: new Date().toISOString(),
          runs,
          warmup,
          budgetMs,
          results: detailedResults,
        },
        null,
        2
      )
    );
  }
}

await main();

function includesIgnoreCase(text, pat) {
  if (!pat) return true;
  return String(text).toLowerCase().includes(String(pat).toLowerCase());
}
