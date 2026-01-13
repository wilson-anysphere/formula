import { existsSync } from "node:fs";
import { resolve } from "node:path";

import {
  defaultDesktopBinPath,
  percentile,
  runOnce,
  type StartupMetrics,
} from "./desktopStartupRunnerShared.ts";

type Summary = {
  runs: number;
  windowVisible: { p50: number; p95: number };
  tti: { p50: number; p95: number };
};

function parseArgs(argv: string[]): {
  runs: number;
  timeoutMs: number;
  binPath: string | null;
  allowInCi: boolean;
} {
  const args = [...argv];
  const out = {
    runs: 20,
    timeoutMs: 15_000,
    binPath: null as string | null,
    allowInCi: false,
  };

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;
    if (arg === "--runs" && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === "--timeout-ms" && args[0]) out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if ((arg === "--bin" || arg === "--bin-path") && args[0]) out.binPath = args.shift()!;
    else if (arg === "--allow-ci") out.allowInCi = true;
  }

  return out;
}

function printSummary(summary: Summary): void {
  // eslint-disable-next-line no-console
  console.log(
    [
      "[desktop-startup]",
      `runs=${summary.runs}`,
      `windowVisible(p50=${summary.windowVisible.p50}ms,p95=${summary.windowVisible.p95}ms)`,
      `tti(p50=${summary.tti.p50}ms,p95=${summary.tti.p95}ms)`,
    ].join(" "),
  );
}

async function main(): Promise<void> {
  const { runs, timeoutMs, binPath: argBin, allowInCi } = parseArgs(process.argv.slice(2));

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
      "Desktop binary not found. Build it via `(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)` and pass --bin <path>.",
    );
  }

  const results: StartupMetrics[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-startup] run ${i + 1}/${runs}...`);
    results.push(await runOnce({ binPath, timeoutMs, envOverrides: {} }));
  }

  const windowVisible = results.map((r) => r.windowVisibleMs).sort((a, b) => a - b);
  const tti = results.map((r) => r.ttiMs).sort((a, b) => a - b);

  const summary: Summary = {
    runs: results.length,
    windowVisible: {
      p50: percentile(windowVisible, 0.5),
      p95: percentile(windowVisible, 0.95),
    },
    tti: {
      p50: percentile(tti, 0.5),
      p95: percentile(tti, 0.95),
    },
  };

  printSummary(summary);
}

await main();
