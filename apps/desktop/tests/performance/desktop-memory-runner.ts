import { spawn, spawnSync } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
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
  rssMb: { p50: number; p95: number };
  targetMb: number | null;
};

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "../../../..");

const perfHome = resolve(repoRoot, "target", "perf-home");

function parseArgs(argv: string[]): {
  runs: number;
  timeoutMs: number;
  settleMs: number;
  binPath: string | null;
  targetMb: number | null;
  allowInCi: boolean;
  enforce: boolean;
} {
  const args = [...argv];
  const envRuns = Number(process.env.FORMULA_DESKTOP_MEMORY_RUNS ?? "") || 10;
  const envTimeoutMs = Number(process.env.FORMULA_DESKTOP_MEMORY_TIMEOUT_MS ?? "") || 30_000;
  const envSettleMs = Number(process.env.FORMULA_DESKTOP_MEMORY_SETTLE_MS ?? "") || 5_000;
  const envTargetMb = Number(process.env.FORMULA_DESKTOP_MEMORY_TARGET_MB ?? "") || null;
  const envEnforce = process.env.FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH === "1";
  const envBin = process.env.FORMULA_DESKTOP_BIN ?? null;

  const out = {
    runs: Math.max(1, envRuns),
    timeoutMs: Math.max(1, envTimeoutMs),
    settleMs: Math.max(0, envSettleMs),
    binPath: envBin as string | null,
    targetMb: envTargetMb ? Math.max(1, envTargetMb) : null,
    allowInCi: false,
    enforce: envEnforce,
  };

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;
    if (arg === "--runs" && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === "--timeout-ms" && args[0])
      out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if (arg === "--settle-ms" && args[0]) out.settleMs = Math.max(0, Number(args.shift()) || out.settleMs);
    else if ((arg === "--bin" || arg === "--bin-path") && args[0]) out.binPath = args.shift()!;
    else if (arg === "--target-mb" && args[0]) {
      const raw = Number(args.shift());
      if (Number.isFinite(raw) && raw > 0) out.targetMb = raw;
    }
    else if (arg === "--allow-ci") out.allowInCi = true;
    else if (arg === "--enforce") out.enforce = true;
  }

  return out;
}

function defaultDesktopBinPath(): string | null {
  const exe = process.platform === "win32" ? "formula-desktop.exe" : "formula-desktop";
  const candidates = [
    // Cargo workspace default target dir (most common).
    resolve(repoRoot, "target/release", exe),
    resolve(repoRoot, "target/debug", exe),
    // Fallbacks in case a caller built with a custom target dir rooted under the app.
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

function parsePsTable(output: string): { pid: number; ppid: number; rssKb: number }[] {
  const rows: { pid: number; ppid: number; rssKb: number }[] = [];
  for (const line of output.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const parts = trimmed.split(/\s+/);
    if (parts.length < 3) continue;
    const pid = Number(parts[0]);
    const ppid = Number(parts[1]);
    const rssKb = Number(parts[2]);
    if (!Number.isFinite(pid) || !Number.isFinite(ppid) || !Number.isFinite(rssKb)) continue;
    rows.push({ pid, ppid, rssKb });
  }
  return rows;
}

function processTreeRssKb(rootPid: number): number {
  // `ps` RSS is reported in KB on both Linux and macOS.
  // BSD/mac: `ps -ax -o pid= -o ppid= -o rss=`
  // GNU/Linux: same flags work.
  const proc = spawnSync("ps", ["-ax", "-o", "pid=", "-o", "ppid=", "-o", "rss="], {
    encoding: "utf8",
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  if (proc.status !== 0) {
    throw new Error(`ps failed (exit ${proc.status}):\n${proc.stderr}`);
  }
  const rows = parsePsTable(proc.stdout);
  const childrenByParent = new Map<number, number[]>();
  const rssByPid = new Map<number, number>();
  for (const row of rows) {
    rssByPid.set(row.pid, row.rssKb);
    const list = childrenByParent.get(row.ppid);
    if (list) list.push(row.pid);
    else childrenByParent.set(row.ppid, [row.pid]);
  }

  let total = 0;
  const stack = [rootPid];
  const seen = new Set<number>();
  while (stack.length > 0) {
    const pid = stack.pop();
    if (pid == null) continue;
    if (seen.has(pid)) continue;
    seen.add(pid);
    const rss = rssByPid.get(pid);
    if (rss != null) total += rss;
    const kids = childrenByParent.get(pid);
    if (kids) stack.push(...kids);
  }
  return total;
}

async function runOnce(binPath: string, timeoutMs: number, settleMs: number): Promise<number> {
  const useXvfb = shouldUseXvfb();
  const command = useXvfb ? resolve(repoRoot, "scripts/xvfb-run-safe.sh") : binPath;
  const args = useXvfb ? [binPath] : [];

  mkdirSync(perfHome, { recursive: true });

  return await new Promise<number>((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: ["ignore", "pipe", "pipe"],
      env: {
        ...process.env,
        FORMULA_STARTUP_METRICS: "1",
        HOME: perfHome,
        USERPROFILE: perfHome,
      },
    });

    let done = false;
    let captured: StartupMetrics | null = null;
    let sampledRssMb: number | null = null;
    let sampleTimer: NodeJS.Timeout | null = null;
    let captureKillTimer: NodeJS.Timeout | null = null;
    let exitDeadline: NodeJS.Timeout | null = null;
    const deadline = setTimeout(() => {
      if (done) return;
      done = true;
      try {
        child.kill();
      } catch {
        // ignore
      }
      cleanup();
      rejectPromise(new Error(`Timed out after ${timeoutMs}ms waiting for startup metrics`));
    }, timeoutMs);

    const cleanup = () => {
      clearTimeout(deadline);
      if (sampleTimer) clearTimeout(sampleTimer);
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
      clearTimeout(deadline);

      sampleTimer = setTimeout(() => {
        if (done) return;
        try {
          const rootPid = child.pid;
          if (!rootPid || rootPid <= 0) {
            throw new Error("Desktop process PID was not available for memory sampling");
          }
          const rssKb = processTreeRssKb(rootPid);
          const rssMb = rssKb / 1024;
          sampledRssMb = rssMb;
        } catch (err) {
          done = true;
          cleanup();
          rejectPromise(err instanceof Error ? err : new Error(String(err)));
          return;
        } finally {
          try {
            child.kill();
          } catch {
            // ignore
          }
        }

        // If the process doesn't exit quickly, force-kill it so we don't accumulate
        // background GUI processes during a multi-run benchmark.
        captureKillTimer = setTimeout(() => {
          try {
            child.kill("SIGKILL");
          } catch {
            child.kill();
          }
        }, 2000);

        exitDeadline = setTimeout(() => {
          if (done) return;
          done = true;
          cleanup();
          rejectPromise(new Error("Timed out waiting for desktop process to exit after sampling memory"));
        }, 5000);
      }, settleMs);
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
      if (done) return;
      done = true;
      cleanup();
      if (sampledRssMb != null) {
        resolvePromise(sampledRssMb);
      } else if (captured) {
        rejectPromise(new Error(`Desktop process exited before memory could be sampled (code=${code}, signal=${signal})`));
      } else {
        rejectPromise(
          new Error(
            `Desktop process exited before reporting startup metrics (code=${code}, signal=${signal})`,
          ),
        );
      }
    });
  });
}

function printSummary(summary: Summary): void {
  const targetSuffix = summary.targetMb ? ` target=${summary.targetMb}MB` : "";
  // eslint-disable-next-line no-console
  console.log(
    [
      "[desktop-memory]",
      `runs=${summary.runs}`,
      `idleRssMb(p50=${summary.rssMb.p50.toFixed(1)}MB,p95=${summary.rssMb.p95.toFixed(1)}MB${targetSuffix})`,
    ].join(" "),
  );
}

async function main(): Promise<void> {
  const { runs, timeoutMs, settleMs, binPath: argBin, targetMb, allowInCi, enforce } = parseArgs(
    process.argv.slice(2),
  );

  if (process.env.CI && !allowInCi && process.env.FORMULA_RUN_DESKTOP_MEMORY_BENCH !== "1") {
    // eslint-disable-next-line no-console
    console.log(
      "[desktop-memory] skipping in CI (set FORMULA_RUN_DESKTOP_MEMORY_BENCH=1 or pass --allow-ci to run)",
    );
    return;
  }

  const binPath = argBin ? resolve(argBin) : defaultDesktopBinPath();
  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      "Desktop binary not found. Build it via `bash scripts/cargo_agent.sh build -p formula-desktop-tauri --bin formula-desktop --release --features desktop` and pass --bin <path>.",
    );
  }

  // eslint-disable-next-line no-console
  console.log(
    "[desktop-memory] measuring process RSS (resident set size) for the desktop app after TTI.\n" +
      `- runs: ${runs} (override via --runs or FORMULA_DESKTOP_MEMORY_RUNS)\n` +
      `- settle: ${settleMs}ms (override via --settle-ms or FORMULA_DESKTOP_MEMORY_SETTLE_MS)\n` +
      `- home: ${resolve(repoRoot, "target", "perf-home")} (repo-local)\n` +
      (targetMb ? `- target: ${targetMb}MB (override via --target-mb or FORMULA_DESKTOP_MEMORY_TARGET_MB)\n` : ""),
  );

  const results: number[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory] run ${i + 1}/${runs}...`);
    const rss = await runOnce(binPath, timeoutMs, settleMs);
    results.push(rss);
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory]   idleRssMb=${rss.toFixed(1)}MB`);
  }

  const summary: Summary = {
    runs: results.length,
    rssMb: {
      p50: percentile(results, 0.5),
      p95: percentile(results, 0.95),
    },
    targetMb,
  };

  printSummary(summary);

  if (enforce && targetMb && summary.rssMb.p95 > targetMb) {
    // eslint-disable-next-line no-console
    console.error(
      `[desktop-memory] FAIL: p95=${summary.rssMb.p95.toFixed(1)}MB exceeds target=${targetMb}MB (set FORMULA_DESKTOP_MEMORY_TARGET_MB to adjust)`,
    );
    process.exitCode = 1;
  }
}

await main();
