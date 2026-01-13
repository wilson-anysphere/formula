import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

import {
  defaultDesktopBinPath,
  percentile,
  runOnce,
  type StartupMetrics,
} from "./desktopStartupUtil.ts";

// Benchmark environment knobs:
// - `FORMULA_DISABLE_STARTUP_UPDATE_CHECK=1` prevents the release updater from running a
//   background check/download on startup, which can add nondeterministic CPU/memory/network
//   activity and skew startup/idle-memory measurements.
// - `FORMULA_STARTUP_METRICS=1` enables the Rust-side one-line startup metrics log we parse.

type Summary = {
  runs: number;
  windowVisible: { p50: number; p95: number; targetMs: number };
  firstRender: { p50: number; p95: number };
  tti: { p50: number; p95: number; targetMs: number };
  enforce: boolean;
  webviewLoaded?: { p50: number; p95: number };
};

function parseArgs(argv: string[]): {
  runs: number;
  timeoutMs: number;
  binPath: string | null;
  windowTargetMs: number;
  ttiTargetMs: number;
  allowInCi: boolean;
  enforce: boolean;
  jsonPath: string | null;
} {
  const args = [...argv];
  const envRuns = Number(process.env.FORMULA_DESKTOP_STARTUP_RUNS ?? "") || 20;
  const envTimeoutMs = Number(process.env.FORMULA_DESKTOP_STARTUP_TIMEOUT_MS ?? "") || 15_000;
  const envBin = process.env.FORMULA_DESKTOP_BIN ?? null;
  const envWindowTargetMs = Number(process.env.FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS ?? "") || 500;
  const envTtiTargetMs = Number(process.env.FORMULA_DESKTOP_TTI_TARGET_MS ?? "") || 1000;
  const envEnforce = process.env.FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH === "1";
  const out = {
    runs: Math.max(1, envRuns),
    timeoutMs: Math.max(1, envTimeoutMs),
    binPath: envBin as string | null,
    windowTargetMs: Math.max(1, envWindowTargetMs),
    ttiTargetMs: Math.max(1, envTtiTargetMs),
    allowInCi: false,
    enforce: envEnforce,
    jsonPath: null as string | null,
  };

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;
    if (arg === "--runs" && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === "--timeout-ms" && args[0]) out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if ((arg === "--bin" || arg === "--bin-path") && args[0]) out.binPath = args.shift()!;
    else if ((arg === "--window-target-ms" || arg === "--window-visible-target-ms") && args[0])
      out.windowTargetMs = Math.max(1, Number(args.shift()) || out.windowTargetMs);
    else if (arg === "--tti-target-ms" && args[0])
      out.ttiTargetMs = Math.max(1, Number(args.shift()) || out.ttiTargetMs);
    else if ((arg === "--json" || arg === "--json-path") && args[0]) out.jsonPath = args.shift()!;
    else if (arg === "--allow-ci") out.allowInCi = true;
    else if (arg === "--enforce") out.enforce = true;
  }

  return out;
}

function printSummary(summary: Summary): void {
  const windowStatus = summary.windowVisible.p95 <= summary.windowVisible.targetMs ? "PASS" : "FAIL";
  const ttiStatus = summary.tti.p95 <= summary.tti.targetMs ? "PASS" : "FAIL";
  // eslint-disable-next-line no-console
  console.log(
    [
      "[desktop-startup]",
      `runs=${summary.runs}`,
      `windowVisible(${windowStatus} p50=${summary.windowVisible.p50}ms,p95=${summary.windowVisible.p95}ms,target=${summary.windowVisible.targetMs}ms)`,
      `firstRender(p50=${summary.firstRender.p50}ms,p95=${summary.firstRender.p95}ms)`,
      summary.webviewLoaded
        ? `webviewLoaded(p50=${summary.webviewLoaded.p50}ms,p95=${summary.webviewLoaded.p95}ms)`
        : "webviewLoaded(n/a)",
      `tti(${ttiStatus} p50=${summary.tti.p50}ms,p95=${summary.tti.p95}ms,target=${summary.tti.targetMs}ms)`,
      summary.enforce ? "enforced=1" : "enforced=0",
    ].join(" "),
  );
}

async function main(): Promise<void> {
  const { runs, timeoutMs, binPath: argBin, windowTargetMs, ttiTargetMs, allowInCi, enforce, jsonPath } = parseArgs(
    process.argv.slice(2),
  );

  if (process.env.CI && !allowInCi && process.env.FORMULA_RUN_DESKTOP_STARTUP_BENCH !== "1") {
    // eslint-disable-next-line no-console
    console.log(
      "[desktop-startup] skipping in CI (set FORMULA_RUN_DESKTOP_STARTUP_BENCH=1 or pass --allow-ci to run)",
    );
    return;
  }

  const binPath = argBin ? resolve(argBin) : defaultDesktopBinPath();
  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      "Desktop binary not found. Build it via `bash scripts/cargo_agent.sh build -p formula-desktop-tauri --bin formula-desktop --release --features desktop` and pass --bin <path> (or set FORMULA_DESKTOP_BIN).",
    );
  }

  // eslint-disable-next-line no-console
  console.log(
    "[desktop-startup] measuring real desktop cold-start timings (window-visible + TTI).\n" +
      `- runs: ${runs} (override via --runs or FORMULA_DESKTOP_STARTUP_RUNS)\n` +
      `- timeout: ${timeoutMs}ms (override via --timeout-ms or FORMULA_DESKTOP_STARTUP_TIMEOUT_MS)\n` +
      `- window target: ${windowTargetMs}ms (override via --window-target-ms or FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS)\n` +
      `- tti target: ${ttiTargetMs}ms (override via --tti-target-ms or FORMULA_DESKTOP_TTI_TARGET_MS)\n` +
      `- home: target/perf-home (repo-local; override with FORMULA_PERF_HOME; set FORMULA_DESKTOP_BENCH_RESET_HOME=1 to reset between iterations)\n` +
      (enforce
        ? "- enforcement: enabled (set FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=0 to disable)\n"
        : "- enforcement: disabled (set FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=1 or pass --enforce to fail on regression)\n"),
  );

  const results: StartupMetrics[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-startup] run ${i + 1}/${runs}...`);
    results.push(
      await runOnce({
        binPath,
        timeoutMs,
        envOverrides: { FORMULA_DISABLE_STARTUP_UPDATE_CHECK: "1" },
      }),
    );
  }

  const windowVisible = results.map((r) => r.windowVisibleMs).sort((a, b) => a - b);
  const firstRender = results
    .map((r) => r.firstRenderMs)
    .filter((v): v is number => typeof v === "number" && Number.isFinite(v))
    .sort((a, b) => a - b);
  const webviewLoaded = results
    .map((r) => r.webviewLoadedMs)
    .filter((v): v is number => typeof v === "number" && Number.isFinite(v))
    .sort((a, b) => a - b);
  const tti = results.map((r) => r.ttiMs).sort((a, b) => a - b);

  // Mirror the benchmark harness policy for handling missing `webview_loaded_ms` values:
  // skip reporting it unless we have a representative sample.
  const minWebviewLoadedFraction = 0.8;
  const minWebviewLoadedRuns = Math.ceil(results.length * minWebviewLoadedFraction);

  const summary: Summary = {
    runs: results.length,
    windowVisible: {
      p50: percentile(windowVisible, 0.5),
      p95: percentile(windowVisible, 0.95),
      targetMs: windowTargetMs,
    },
    firstRender: {
      p50: percentile(firstRender, 0.5),
      p95: percentile(firstRender, 0.95),
    },
    ...(webviewLoaded.length >= minWebviewLoadedRuns
      ? {
          webviewLoaded: {
            p50: percentile(webviewLoaded, 0.5),
            p95: percentile(webviewLoaded, 0.95),
          },
        }
      : {}),
    tti: {
      p50: percentile(tti, 0.5),
      p95: percentile(tti, 0.95),
      targetMs: ttiTargetMs,
    },
    enforce,
  };

  printSummary(summary);

  if (jsonPath) {
    const outputPath = resolve(jsonPath);
    mkdirSync(dirname(outputPath), { recursive: true });
    writeFileSync(
      outputPath,
      JSON.stringify(
        {
          generatedAt: new Date().toISOString(),
          platform: process.platform,
          binPath,
          runs: results.length,
          samples: results,
          summary,
        },
        null,
        2,
      ),
      "utf8",
    );
  }

  if (enforce) {
    const failed =
      summary.windowVisible.p95 > summary.windowVisible.targetMs || summary.tti.p95 > summary.tti.targetMs;
    if (failed) process.exitCode = 1;
  }
}

await main();
