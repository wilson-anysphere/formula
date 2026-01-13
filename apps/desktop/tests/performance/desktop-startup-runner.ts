import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

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

type StartupBenchKind = "shell" | "full";
type StartupMode = "cold" | "warm";

type Summary = {
  runs: number;
  windowVisible: { p50: number; p95: number; targetMs: number };
  // `first_render_ms` is only meaningful for the full-app benchmark (the shell benchmark uses a
  // minimal page and exits before the app grid is rendered).
  firstRender: { p50: number | null; p95: number | null };
  tti: { p50: number; p95: number; targetMs: number };
  enforce: boolean;
  webviewLoaded?: { p50: number; p95: number; targetMs: number };
};

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "../../../..");

function parseBenchKindFromEnv(): StartupBenchKind | null {
  const raw = (process.env.FORMULA_DESKTOP_STARTUP_BENCH_KIND ?? "").trim().toLowerCase();
  if (!raw) return null;
  if (raw === "shell") return "shell";
  if (raw === "full") return "full";
  return null;
}

function parseArgs(argv: string[]): {
  runs: number;
  timeoutMs: number;
  binPath: string | null;
  mode: StartupMode;
  windowTargetMs: number;
  webviewLoadedTargetMs: number;
  ttiTargetMs: number;
  allowInCi: boolean;
  enforce: boolean;
  jsonPath: string | null;
  benchKind: StartupBenchKind;
} {
  const args = [...argv];

  const modeRaw = (process.env.FORMULA_DESKTOP_STARTUP_MODE ?? "cold").trim().toLowerCase();
  if (modeRaw !== "cold" && modeRaw !== "warm") {
    throw new Error(
      `Invalid FORMULA_DESKTOP_STARTUP_MODE=${JSON.stringify(modeRaw)} (expected "cold" or "warm")`,
    );
  }
  let mode: StartupMode = modeRaw;

  const envRuns = Number(process.env.FORMULA_DESKTOP_STARTUP_RUNS ?? "") || 20;
  const envTimeoutMs = Number(process.env.FORMULA_DESKTOP_STARTUP_TIMEOUT_MS ?? "") || 15_000;
  const envBin = process.env.FORMULA_DESKTOP_BIN ?? null;

  const coldWindowTargetMs =
    Number(
      process.env.FORMULA_DESKTOP_COLD_WINDOW_VISIBLE_TARGET_MS ??
        process.env.FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS ??
        "",
    ) || 500;
  const coldTtiTargetMs =
    Number(
      process.env.FORMULA_DESKTOP_COLD_TTI_TARGET_MS ??
        process.env.FORMULA_DESKTOP_TTI_TARGET_MS ??
        "",
    ) || 1000;
  const warmWindowTargetMs =
    Number(process.env.FORMULA_DESKTOP_WARM_WINDOW_VISIBLE_TARGET_MS ?? "") || coldWindowTargetMs;
  const warmTtiTargetMs = Number(process.env.FORMULA_DESKTOP_WARM_TTI_TARGET_MS ?? "") || coldTtiTargetMs;

  const envWebviewLoadedTargetMs = Number(process.env.FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS ?? "") || 800;
  const envEnforce = process.env.FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH === "1";

  const envKind = parseBenchKindFromEnv();
  const defaultKind: StartupBenchKind = envKind ?? (process.env.CI ? "shell" : "full");

  let windowTargetMsOverride: number | null = null;
  let ttiTargetMsOverride: number | null = null;

  const out = {
    runs: Math.max(1, envRuns),
    timeoutMs: Math.max(1, envTimeoutMs),
    binPath: envBin as string | null,
    mode,
    windowTargetMs: 0,
    webviewLoadedTargetMs: Math.max(1, envWebviewLoadedTargetMs),
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
      const raw = String(args.shift()).trim().toLowerCase();
      if (raw !== "cold" && raw !== "warm") {
        throw new Error(`Invalid --mode ${JSON.stringify(raw)} (expected "cold" or "warm")`);
      }
      mode = raw;
      out.mode = mode;
    } else if (arg === "--runs" && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === "--timeout-ms" && args[0]) out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if ((arg === "--bin" || arg === "--bin-path") && args[0]) out.binPath = args.shift()!;
    else if ((arg === "--window-target-ms" || arg === "--window-visible-target-ms") && args[0])
      windowTargetMsOverride = Math.max(1, Number(args.shift()) || 0);
    else if ((arg === "--webview-loaded-target-ms" || arg === "--webview-target-ms") && args[0])
      out.webviewLoadedTargetMs = Math.max(1, Number(args.shift()) || out.webviewLoadedTargetMs);
    else if (arg === "--tti-target-ms" && args[0])
      ttiTargetMsOverride = Math.max(1, Number(args.shift()) || 0);
    else if ((arg === "--json" || arg === "--json-path") && args[0]) out.jsonPath = args.shift()!;
    else if (arg === "--allow-ci") out.allowInCi = true;
    else if (arg === "--enforce") out.enforce = true;
    else if (arg === "--startup-bench" || arg === "--shell") out.benchKind = "shell";
    else if (arg === "--full") out.benchKind = "full";
  }

  out.windowTargetMs =
    windowTargetMsOverride ??
    (mode === "warm" ? Math.max(1, warmWindowTargetMs) : Math.max(1, coldWindowTargetMs));
  out.ttiTargetMs =
    ttiTargetMsOverride ?? (mode === "warm" ? Math.max(1, warmTtiTargetMs) : Math.max(1, coldTtiTargetMs));

  return out;
}

function formatMaybeMs(ms: number | null): string {
  if (ms === null || !Number.isFinite(ms)) return "n/a";
  return `${ms}ms`;
}

function printSummary(summary: Summary, benchKind: StartupBenchKind): void {
  const windowStatus = summary.windowVisible.p95 <= summary.windowVisible.targetMs ? "PASS" : "FAIL";
  const webviewLoadedStatus =
    summary.webviewLoaded && summary.webviewLoaded.p95 <= summary.webviewLoaded.targetMs ? "PASS" : "FAIL";
  const ttiStatus = summary.tti.p95 <= summary.tti.targetMs ? "PASS" : "FAIL";
  // eslint-disable-next-line no-console
  console.log(
    [
      benchKind === "shell" ? "[desktop-shell-startup]" : "[desktop-startup]",
      `runs=${summary.runs}`,
      `windowVisible(${windowStatus} p50=${summary.windowVisible.p50}ms,p95=${summary.windowVisible.p95}ms,target=${summary.windowVisible.targetMs}ms)`,
      `firstRender(p50=${formatMaybeMs(summary.firstRender.p50)},p95=${formatMaybeMs(summary.firstRender.p95)})`,
      summary.webviewLoaded
        ? `webviewLoaded(${webviewLoadedStatus} p50=${summary.webviewLoaded.p50}ms,p95=${summary.webviewLoaded.p95}ms,target=${summary.webviewLoaded.targetMs}ms)`
        : "webviewLoaded(n/a)",
      `tti(${ttiStatus} p50=${summary.tti.p50}ms,p95=${summary.tti.p95}ms,target=${summary.tti.targetMs}ms)`,
      summary.enforce ? "enforced=1" : "enforced=0",
    ].join(" "),
  );
}

async function main(): Promise<void> {
  const {
    runs,
    timeoutMs,
    binPath: argBin,
    mode,
    windowTargetMs,
    webviewLoadedTargetMs,
    ttiTargetMs,
    allowInCi,
    enforce,
    jsonPath,
    benchKind,
  } = parseArgs(process.argv.slice(2));

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

  const argv = benchKind === "shell" ? ["--startup-bench"] : [];

  // eslint-disable-next-line no-console
  console.log(
    "[desktop-startup] measuring desktop startup timings (window-visible + first-render + TTI).\n" +
      `- kind: ${benchKind} (set FORMULA_DESKTOP_STARTUP_BENCH_KIND=shell|full or pass --startup-bench/--full)\n` +
      `- mode: ${mode} (set FORMULA_DESKTOP_STARTUP_MODE=cold|warm or pass --mode)\n` +
      `- runs: ${runs} (override via --runs or FORMULA_DESKTOP_STARTUP_RUNS)\n` +
      `- timeout: ${timeoutMs}ms (override via --timeout-ms or FORMULA_DESKTOP_STARTUP_TIMEOUT_MS)\n` +
      `- window target: ${windowTargetMs}ms (override via --window-target-ms)\n` +
      `- webviewLoaded target: ${webviewLoadedTargetMs}ms (override via --webview-loaded-target-ms or FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS)\n` +
      `- tti target: ${ttiTargetMs}ms (override via --tti-target-ms)\n` +
      `- home: target/perf-home (repo-local; override with FORMULA_PERF_HOME)\n` +
      (enforce
        ? "- enforcement: enabled (set FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=0 to disable)\n"
        : "- enforcement: disabled (set FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=1 or pass --enforce to fail on regression)\n"),
  );

  const results: StartupMetrics[] = [];
  const envOverrides: NodeJS.ProcessEnv = { FORMULA_DISABLE_STARTUP_UPDATE_CHECK: "1" };

  const perfHome =
    process.env.FORMULA_PERF_HOME && process.env.FORMULA_PERF_HOME.trim() !== ""
      ? resolve(repoRoot, process.env.FORMULA_PERF_HOME)
      : resolve(repoRoot, "target", "perf-home");
  const profileRoot = resolve(perfHome, `desktop-startup-${benchKind}-${mode}-${Date.now()}-${process.pid}`);

  // `desktopStartupUtil.runOnce()` supports resetting the selected profile dir via
  // `FORMULA_DESKTOP_BENCH_RESET_HOME=1`. Manage it here so cold/warm semantics remain clear.
  const prevResetHome = process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;
  const setResetHome = (value: string | undefined) => {
    if (value === undefined) {
      delete process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;
    } else {
      process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = value;
    }
  };

  try {
    if (mode === "warm") {
      const profileDir = resolve(profileRoot, "profile");
      setResetHome("1");
      // eslint-disable-next-line no-console
      console.log(`[desktop-${benchKind}-startup] warmup run 1/1 (warm, profile=${profileDir})...`);
      await runOnce({ binPath, timeoutMs, argv, envOverrides, profileDir });

      setResetHome(undefined);
      for (let i = 0; i < runs; i += 1) {
        // eslint-disable-next-line no-console
        console.log(
          `[desktop-${benchKind}-startup] run ${i + 1}/${runs} (warm, profile=${profileDir})...`,
        );
        results.push(await runOnce({ binPath, timeoutMs, argv, envOverrides, profileDir }));
      }
    } else {
      setResetHome("1");
      for (let i = 0; i < runs; i += 1) {
        const profileDir = resolve(profileRoot, `run-${String(i + 1).padStart(2, "0")}`);
        // eslint-disable-next-line no-console
        console.log(
          `[desktop-${benchKind}-startup] run ${i + 1}/${runs} (cold, profile=${profileDir})...`,
        );
        results.push(await runOnce({ binPath, timeoutMs, argv, envOverrides, profileDir }));
      }
    }
  } finally {
    setResetHome(prevResetHome);
  }

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

  // `webview_loaded_ms` is recorded by the Rust host (via a native page-load callback) and should
  // generally be available for every run. Keep this best-effort skip policy anyway so the runner
  // can still work against older binaries and so we don't compute p95 over a biased tiny sample
  // if something regresses.
  const minWebviewLoadedFraction = 0.8;
  const minWebviewLoadedRuns = Math.ceil(results.length * minWebviewLoadedFraction);

  const summary: Summary = {
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
          }
        : { p50: null, p95: null },
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
      (summary.webviewLoaded !== undefined && summary.webviewLoaded.p95 > summary.webviewLoaded.targetMs);
    if (failed) process.exitCode = 1;
  }
}

await main();
