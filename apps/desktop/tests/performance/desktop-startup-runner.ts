import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { createInterface } from "node:readline";
import { fileURLToPath } from "node:url";

type StartupMetrics = {
  windowVisibleMs: number;
  webviewLoadedMs: number | null;
  ttiMs: number;
};

type Summary = {
  runs: number;
  windowVisible: { p50: number; p95: number };
  tti: { p50: number; p95: number };
};

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "../../../..");

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

function defaultDesktopBinPath(): string | null {
  const exe = process.platform === "win32" ? "formula-desktop.exe" : "formula-desktop";
  const candidates = [
    resolve(repoRoot, "apps/desktop/src-tauri/target/release", exe),
    resolve(repoRoot, "apps/desktop/src-tauri/target/debug", exe),
  ];
  for (const p of candidates) {
    if (existsSync(p)) return p;
  }
  return null;
}

function shouldUseXvfb(): boolean {
  if (process.platform !== "linux") return false;
  // If DISPLAY is set, assume an X server is already available.
  if (process.env.DISPLAY && process.env.DISPLAY.trim() !== "") return false;
  const xvfb = resolve(repoRoot, "scripts/xvfb-run-safe.sh");
  return existsSync(xvfb);
}

function percentile(values: number[], p: number): number {
  if (values.length === 0) return NaN;
  const sorted = [...values].sort((a, b) => a - b);
  const idx = Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * p) - 1));
  return sorted[idx]!;
}

function parseStartupLine(line: string): StartupMetrics | null {
  // Example:
  // [startup] window_visible_ms=123 webview_loaded_ms=234 tti_ms=456
  const match = line.match(
    /^\[startup\]\s+window_visible_ms=(\d+)\s+webview_loaded_ms=(\d+|n\/a)\s+tti_ms=(\d+)\s*$/,
  );
  if (!match) return null;
  const windowVisibleMs = Number(match[1]);
  const webviewLoadedRaw = match[2]!;
  const webviewLoadedMs = webviewLoadedRaw === "n/a" ? null : Number(webviewLoadedRaw);
  const ttiMs = Number(match[3]);
  if (!Number.isFinite(windowVisibleMs) || !Number.isFinite(ttiMs)) return null;
  return { windowVisibleMs, webviewLoadedMs, ttiMs };
}

async function runOnce(binPath: string, timeoutMs: number): Promise<StartupMetrics> {
  const useXvfb = shouldUseXvfb();
  const command = useXvfb ? resolve(repoRoot, "scripts/xvfb-run-safe.sh") : binPath;
  const args = useXvfb ? [binPath] : [];

  return await new Promise<StartupMetrics>((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: ["ignore", "pipe", "pipe"],
      env: {
        ...process.env,
        // Enable the Rust-side single-line log in release builds.
        FORMULA_STARTUP_METRICS: "1",
      },
    });

    let done = false;
    let captured: StartupMetrics | null = null;
    let captureKillTimer: NodeJS.Timeout | null = null;
    let exitDeadline: NodeJS.Timeout | null = null;
    const deadline = setTimeout(() => {
      if (done) return;
      done = true;
      child.kill();
      cleanup();
      rejectPromise(new Error(`Timed out after ${timeoutMs}ms waiting for startup metrics`));
    }, timeoutMs);

    const cleanup = () => {
      clearTimeout(deadline);
      if (captureKillTimer) clearTimeout(captureKillTimer);
      if (exitDeadline) clearTimeout(exitDeadline);
      rlOut.close();
      rlErr.close();
    };

    const onLine = (line: string) => {
      if (done || captured) return;
      const parsed = parseStartupLine(line.trim());
      if (!parsed) return;
      captured = parsed;
      // We got the data we came for; don't fail the run just because shutdown is slow.
      clearTimeout(deadline);
      // Stop the app after capturing the metrics so we can run multiple iterations.
      child.kill();
      exitDeadline = setTimeout(() => {
        if (done) return;
        done = true;
        try {
          child.kill("SIGKILL");
        } catch {
          child.kill();
        }
        cleanup();
        rejectPromise(new Error("Timed out waiting for desktop process to exit after capturing metrics"));
      }, 5000);
      // If the process doesn't exit quickly, force-kill it so we don't accumulate
      // background GUI processes during a multi-run benchmark.
      captureKillTimer = setTimeout(() => {
        try {
          child.kill("SIGKILL");
        } catch {
          // Ignore; not all platforms support SIGKILL.
          child.kill();
        }
      }, 2000);
    };

    const rlOut = createInterface({ input: child.stdout! });
    const rlErr = createInterface({ input: child.stderr! });
    rlOut.on("line", onLine);
    rlErr.on("line", onLine);

    child.on("error", (err) => {
      if (done) return;
      done = true;
      cleanup();
      rejectPromise(err);
    });

    child.on("exit", (code, signal) => {
      cleanup();
      if (done) return;
      done = true;
      if (captured) {
        resolvePromise(captured);
        return;
      }
      rejectPromise(
        new Error(`Desktop process exited before reporting metrics (code=${code}, signal=${signal})`),
      );
    });
  });
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
      "Desktop binary not found. Build it via `cargo tauri build` (apps/desktop) and pass --bin <path>.",
    );
  }

  const results: StartupMetrics[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-startup] run ${i + 1}/${runs}...`);
    results.push(await runOnce(binPath, timeoutMs));
  }

  const windowVisible = results.map((r) => r.windowVisibleMs);
  const tti = results.map((r) => r.ttiMs);

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
