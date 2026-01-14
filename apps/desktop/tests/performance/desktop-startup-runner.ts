import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

import {
  defaultDesktopBinPath,
  percentile,
  buildDesktopStartupProfileRoot,
  parseDesktopStartupMode,
  runDesktopStartupIterations,
  resolveDesktopStartupBenchKind,
  resolveDesktopStartupArgv,
  resolveDesktopStartupMode,
  resolveDesktopStartupTargets,
  resolvePerfHome,
  type DesktopStartupBenchKind,
  type DesktopStartupMode,
  type StartupMetrics,
} from "./desktopStartupUtil.ts";

// Benchmark environment knobs:
// - `FORMULA_DISABLE_STARTUP_UPDATE_CHECK=1` prevents the release updater from running a
//   background check/download on startup, which can add nondeterministic CPU/memory/network
//   activity and skew startup/idle-memory measurements.
// - `FORMULA_STARTUP_METRICS=1` enables the Rust-side one-line startup metrics log we parse.

type Summary = {
  mode: DesktopStartupMode;
  runs: number;
  windowVisible: { p50: number; p95: number; targetMs: number };
  // `first_render_ms` is only meaningful for the full-app benchmark (the shell benchmark uses a
  // minimal page and exits before the app grid is rendered).
  firstRender: { p50: number | null; p95: number | null; targetMs: number | null };
  tti: { p50: number; p95: number; targetMs: number };
  enforce: boolean;
  webviewLoaded?: { p50: number; p95: number; targetMs: number };
};

function usage(): string {
  return [
    "Desktop startup benchmark runner (real Tauri binary).",
    "",
    "Usage:",
    "  node scripts/run-node-ts.mjs apps/desktop/tests/performance/desktop-startup-runner.ts [options]",
    "",
    "Options:",
    "  --mode <cold|warm>               Startup mode (env: FORMULA_DESKTOP_STARTUP_MODE, default: cold)",
    "  --runs <n>                       Iterations (env: FORMULA_DESKTOP_STARTUP_RUNS, default: 20)",
    "  --timeout-ms <ms>                Timeout per run (env: FORMULA_DESKTOP_STARTUP_TIMEOUT_MS, default: 15000)",
    "  --bin, --bin-path <path>         Desktop binary path (env: FORMULA_DESKTOP_BIN)",
    "  --startup-bench, --shell         Shell-only startup (default in CI)",
    "  --full                           Full app startup (default locally)",
    "  --window-target-ms <ms>          p95 target (overrides env targets)",
    "  --first-render-target-ms <ms>    p95 target for first_render_ms (full only; overrides env targets)",
    "  --webview-loaded-target-ms <ms>  p95 target (env: FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS / FORMULA_DESKTOP_SHELL_WEBVIEW_LOADED_TARGET_MS)",
    "  --tti-target-ms <ms>             p95 target (overrides env targets)",
    "  --json, --json-path <path>       Write JSON output (samples + summary) to this path",
    "  --enforce                        Exit non-zero if p95 exceeds targets (env: FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=1)",
    "  --allow-ci                       Allow running under CI without FORMULA_RUN_DESKTOP_STARTUP_BENCH=1",
    "  -h, --help                       Show this help and exit",
    "",
    "Notes:",
    "  - Uses isolated profile directories under target/perf-home by default (override via FORMULA_PERF_HOME).",
    "    Each invocation picks a unique profile root to avoid cache pollution across runs.",
    "  - Sets FORMULA_DISABLE_STARTUP_UPDATE_CHECK=1 for stability.",
    "",
  ].join("\n");
}

function parseArgs(argv: string[]): {
  mode: DesktopStartupMode;
  runs: number;
  timeoutMs: number;
  binPath: string | null;
  windowTargetMs: number;
  firstRenderTargetMs: number;
  webviewLoadedTargetMs: number;
  ttiTargetMs: number;
  allowInCi: boolean;
  enforce: boolean;
  jsonPath: string | null;
  benchKind: DesktopStartupBenchKind;
} {
  const args = [...argv];
  let mode: DesktopStartupMode = resolveDesktopStartupMode();

  const envRuns = Number(process.env.FORMULA_DESKTOP_STARTUP_RUNS ?? "") || 20;
  const envTimeoutMs = Number(process.env.FORMULA_DESKTOP_STARTUP_TIMEOUT_MS ?? "") || 15_000;
  const envBin = process.env.FORMULA_DESKTOP_BIN ?? null;

  const envEnforce = process.env.FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH === "1";
  const defaultKind: DesktopStartupBenchKind = resolveDesktopStartupBenchKind();

  let windowTargetMsOverride: number | null = null;
  let firstRenderTargetMsOverride: number | null = null;
  let webviewLoadedTargetMsOverride: number | null = null;
  let ttiTargetMsOverride: number | null = null;

  const out = {
    mode,
    runs: Math.max(1, envRuns),
    timeoutMs: Math.max(1, envTimeoutMs),
    binPath: envBin as string | null,
    windowTargetMs: 0,
    firstRenderTargetMs: 0,
    webviewLoadedTargetMs: 0,
    ttiTargetMs: 0,
    allowInCi: false,
    enforce: envEnforce,
    jsonPath: null as string | null,
    benchKind: defaultKind,
  };

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;

    if (arg === "--mode" && args[0]) {
      const raw = String(args.shift());
      const parsed = parseDesktopStartupMode(raw);
      if (!parsed) {
        throw new Error(`Invalid --mode ${JSON.stringify(raw)} (expected "cold" or "warm")`);
      }
      mode = parsed;
      out.mode = parsed;
    } else if (arg === "--runs" && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === "--timeout-ms" && args[0]) out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if ((arg === "--bin" || arg === "--bin-path") && args[0]) out.binPath = args.shift()!;
    else if ((arg === "--window-target-ms" || arg === "--window-visible-target-ms") && args[0])
      windowTargetMsOverride = Math.max(1, Number(args.shift()) || 0);
    else if (arg === "--first-render-target-ms" && args[0])
      firstRenderTargetMsOverride = Math.max(1, Number(args.shift()) || 0);
    else if ((arg === "--webview-loaded-target-ms" || arg === "--webview-target-ms") && args[0])
      webviewLoadedTargetMsOverride = Math.max(1, Number(args.shift()) || 0);
    else if (arg === "--tti-target-ms" && args[0]) ttiTargetMsOverride = Math.max(1, Number(args.shift()) || 0);
    else if ((arg === "--json" || arg === "--json-path") && args[0]) out.jsonPath = args.shift()!;
    else if (arg === "--allow-ci") out.allowInCi = true;
    else if (arg === "--enforce") out.enforce = true;
    else if (arg === "--startup-bench" || arg === "--shell") out.benchKind = "shell";
    else if (arg === "--full") out.benchKind = "full";
  }

  const targets = resolveDesktopStartupTargets({ benchKind: out.benchKind, mode: out.mode });
  out.windowTargetMs = windowTargetMsOverride ?? Math.max(1, targets.windowVisibleTargetMs);
  out.firstRenderTargetMs = firstRenderTargetMsOverride ?? Math.max(1, targets.firstRenderTargetMs);
  out.webviewLoadedTargetMs = webviewLoadedTargetMsOverride ?? Math.max(1, targets.webviewLoadedTargetMs);
  out.ttiTargetMs = ttiTargetMsOverride ?? Math.max(1, targets.ttiTargetMs);

  return out;
}

function formatMaybeMs(ms: number | null): string {
  if (ms === null || !Number.isFinite(ms)) return "n/a";
  return `${ms}ms`;
}

function printSummary(summary: Summary, benchKind: DesktopStartupBenchKind): void {
  const windowStatus = summary.windowVisible.p95 <= summary.windowVisible.targetMs ? "PASS" : "FAIL";
  const ttiStatus = summary.tti.p95 <= summary.tti.targetMs ? "PASS" : "FAIL";

  const firstRenderPart =
    summary.firstRender.p50 !== null &&
    summary.firstRender.p95 !== null &&
    summary.firstRender.targetMs !== null
      ? (() => {
          const status = summary.firstRender.p95 <= summary.firstRender.targetMs ? "PASS" : "FAIL";
          return `firstRender(${status} p50=${summary.firstRender.p50}ms,p95=${summary.firstRender.p95}ms,target=${summary.firstRender.targetMs}ms)`;
        })()
      : `firstRender(p50=${formatMaybeMs(summary.firstRender.p50)},p95=${formatMaybeMs(summary.firstRender.p95)})`;

  const webviewLoadedStatus =
    summary.webviewLoaded && summary.webviewLoaded.p95 <= summary.webviewLoaded.targetMs ? "PASS" : "FAIL";

  // eslint-disable-next-line no-console
  console.log(
    [
      benchKind === "shell" ? "[desktop-shell-startup]" : "[desktop-startup]",
      `mode=${summary.mode}`,
      `runs=${summary.runs}`,
      `windowVisible(${windowStatus} p50=${summary.windowVisible.p50}ms,p95=${summary.windowVisible.p95}ms,target=${summary.windowVisible.targetMs}ms)`,
      firstRenderPart,
      summary.webviewLoaded
        ? `webviewLoaded(${webviewLoadedStatus} p50=${summary.webviewLoaded.p50}ms,p95=${summary.webviewLoaded.p95}ms,target=${summary.webviewLoaded.targetMs}ms)`
        : "webviewLoaded(n/a)",
      `tti(${ttiStatus} p50=${summary.tti.p50}ms,p95=${summary.tti.p95}ms,target=${summary.tti.targetMs}ms)`,
      summary.enforce ? "enforced=1" : "enforced=0",
    ].join(" "),
  );
}

async function main(): Promise<void> {
  const cliArgs = process.argv.slice(2);
  if (cliArgs.includes("--help") || cliArgs.includes("-h")) {
    // eslint-disable-next-line no-console
    console.log(usage());
    return;
  }

  const {
    mode,
    runs,
    timeoutMs,
    binPath: argBin,
    windowTargetMs,
    firstRenderTargetMs,
    webviewLoadedTargetMs,
    ttiTargetMs,
    allowInCi,
    enforce,
    jsonPath,
    benchKind,
  } = parseArgs(cliArgs);

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

  const argv = resolveDesktopStartupArgv(benchKind);

  const perfHome = resolvePerfHome();
  const profileRoot = buildDesktopStartupProfileRoot({ perfHome, benchKind, mode });

  // eslint-disable-next-line no-console
  console.log(
    "[desktop-startup] measuring desktop startup timings (window-visible + first-render + TTI).\n" +
      `- kind: ${benchKind} (set FORMULA_DESKTOP_STARTUP_BENCH_KIND=shell|full or pass --startup-bench/--shell/--full)\n` +
      `- mode: ${mode} (set FORMULA_DESKTOP_STARTUP_MODE=cold|warm or pass --mode)\n` +
      `- runs: ${runs} (override via --runs or FORMULA_DESKTOP_STARTUP_RUNS)\n` +
      `- timeout: ${timeoutMs}ms (override via --timeout-ms or FORMULA_DESKTOP_STARTUP_TIMEOUT_MS)\n` +
      `- window target: ${windowTargetMs}ms (override via --window-target-ms)\n` +
      `- first render target: ${firstRenderTargetMs}ms (only for kind=full; override via --first-render-target-ms)\n` +
      `- webviewLoaded target: ${webviewLoadedTargetMs}ms (override via --webview-loaded-target-ms)\n` +
      `- tti target: ${ttiTargetMs}ms (override via --tti-target-ms)\n` +
      "- targets env (full): FORMULA_DESKTOP_{COLD,WARM}_{WINDOW_VISIBLE,FIRST_RENDER,TTI}_TARGET_MS + FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS\n" +
      "- targets env (shell overrides): FORMULA_DESKTOP_SHELL_{COLD,WARM}_{WINDOW_VISIBLE,TTI}_TARGET_MS + FORMULA_DESKTOP_SHELL_WEBVIEW_LOADED_TARGET_MS\n" +
      `- perf-home: ${perfHome} (repo-local; override with FORMULA_PERF_HOME)\n` +
      `- profile-root: ${profileRoot}\n` +
      (enforce
        ? "- enforcement: enabled (set FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=0 to disable)\n"
        : "- enforcement: disabled (set FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=1 or pass --enforce to fail on regression)\n"),
  );

  const envOverrides: NodeJS.ProcessEnv = { FORMULA_DISABLE_STARTUP_UPDATE_CHECK: "1" };

  const results: StartupMetrics[] = await runDesktopStartupIterations({
    mode,
    runs,
    timeoutMs,
    binPath,
    argv,
    envOverrides,
    profileRoot,
    onProgress: ({ phase, mode: runMode, iteration, total, profileDir }) => {
      // eslint-disable-next-line no-console
      if (phase === "warmup") {
        console.log(`[desktop-${benchKind}-startup] warmup run 1/1 (warm, profile=${profileDir})...`);
      } else {
        console.log(
          `[desktop-${benchKind}-startup] run ${iteration}/${total} (${runMode}, profile=${profileDir})...`,
        );
      }
    },
  });

  const windowVisible = results.map((r) => r.windowVisibleMs).sort((a, b) => a - b);
  const firstRenderValues =
    benchKind === "full"
      ? results
          .map((r) => r.firstRenderMs)
          .filter((v): v is number => typeof v === "number" && Number.isFinite(v))
          .sort((a, b) => a - b)
      : [];
  const webviewLoadedValues = results
    .map((r) => r.webviewLoadedMs)
    .filter((v): v is number => typeof v === "number" && Number.isFinite(v))
    .sort((a, b) => a - b);
  const tti = results.map((r) => r.ttiMs).sort((a, b) => a - b);

  if (benchKind === "full" && firstRenderValues.length !== results.length) {
    throw new Error(
      "Desktop did not report first_render_ms. Ensure the frontend calls `report_startup_first_render` when the grid becomes visible.",
    );
  }

  // `webview_loaded_ms` is recorded by the Rust host (via a native page-load callback) and should
  // generally be available for every run. Keep this best-effort skip policy anyway so the runner
  // can still work against older binaries and so we don't compute p95 over a biased tiny sample.
  const minWebviewLoadedFraction = 0.8;
  const minWebviewLoadedRuns = Math.ceil(results.length * minWebviewLoadedFraction);

  if (webviewLoadedValues.length === 0) {
    // eslint-disable-next-line no-console
    console.log("[desktop-startup] webview_loaded_ms unavailable (0 runs reported it); skipping metric");
  } else if (webviewLoadedValues.length < minWebviewLoadedRuns) {
    // eslint-disable-next-line no-console
    console.log(
      `[desktop-startup] webview_loaded_ms only available for ${webviewLoadedValues.length}/${results.length} runs (<${Math.round(
        minWebviewLoadedFraction * 100,
      )}%); skipping metric`,
    );
  }

  const summary: Summary = {
    mode,
    runs: results.length,
    windowVisible: {
      p50: percentile(windowVisible, 0.5),
      p95: percentile(windowVisible, 0.95),
      targetMs: windowTargetMs,
    },
    firstRender:
      benchKind === "full" && firstRenderValues.length > 0
        ? {
            p50: percentile(firstRenderValues, 0.5),
            p95: percentile(firstRenderValues, 0.95),
            targetMs: firstRenderTargetMs,
          }
        : { p50: null, p95: null, targetMs: null },
    ...(webviewLoadedValues.length >= minWebviewLoadedRuns
      ? {
          webviewLoaded: {
            p50: percentile(webviewLoadedValues, 0.5),
            p95: percentile(webviewLoadedValues, 0.95),
            targetMs: webviewLoadedTargetMs,
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

  printSummary(summary, benchKind);

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
          perfHome,
          profileRoot,
          mode,
          benchKind,
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
      summary.windowVisible.p95 > summary.windowVisible.targetMs ||
      summary.tti.p95 > summary.tti.targetMs ||
      (summary.webviewLoaded !== undefined && summary.webviewLoaded.p95 > summary.webviewLoaded.targetMs) ||
      (summary.firstRender.targetMs !== null &&
        summary.firstRender.p95 !== null &&
        summary.firstRender.p95 > summary.firstRender.targetMs);
    if (failed) process.exitCode = 1;
  }
}

await main();
