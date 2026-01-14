import { spawn, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readlinkSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { dirname, isAbsolute, parse, resolve, relative } from "node:path";
import { createInterface, type Interface } from "node:readline";

import {
  defaultDesktopBinPath,
  parseStartupLine,
  percentile,
  repoRoot,
  resolvePerfHome,
  shouldUseXvfb,
  terminateProcessTree,
  type StartupMetrics,
  type TerminateProcessTreeMode,
} from "./desktopStartupUtil.ts";

type Summary = {
  runs: number;
  rssMb: { p50: number; p95: number; targetMb: number };
  enforce: boolean;
};

const perfHome = resolvePerfHome();
const perfTmp = resolve(perfHome, "tmp");
const perfXdgConfig = resolve(perfHome, "xdg-config");
const perfXdgCache = resolve(perfHome, "xdg-cache");
const perfXdgState = resolve(perfHome, "xdg-state");
const perfXdgData = resolve(perfHome, "xdg-data");
const perfAppData = resolve(perfHome, "AppData", "Roaming");
const perfLocalAppData = resolve(perfHome, "AppData", "Local");

function isSubpath(parentDir: string, maybeChild: string): boolean {
  const rel = relative(parentDir, maybeChild);
  if (rel === "" || rel.startsWith("..")) return false;
  // `path.relative()` can return an absolute path on Windows when drives differ.
  if (isAbsolute(rel)) return false;
  return true;
}

function usage(): string {
  return [
    "Desktop idle memory benchmark runner (real Tauri binary).",
    "",
    "Usage:",
    "  node scripts/run-node-ts.mjs apps/desktop/tests/performance/desktop-memory-runner.ts [options]",
    "",
    "Options:",
    "  --runs <n>                 Iterations (env: FORMULA_DESKTOP_MEMORY_RUNS, default: 10)",
    "  --timeout-ms <ms>          Timeout per run (env: FORMULA_DESKTOP_MEMORY_TIMEOUT_MS, default: 30000)",
    "  --settle-ms <ms>           Delay after startup before sampling (env: FORMULA_DESKTOP_MEMORY_SETTLE_MS, default: 5000)",
    "  --bin, --bin-path <path>   Desktop binary path (env: FORMULA_DESKTOP_BIN)",
    "  --target-mb <mb>           p95 target (env: FORMULA_DESKTOP_IDLE_RSS_TARGET_MB, default: 100)",
    "  --json, --json-path <path> Write JSON output (samples + summary) to this path",
    "  --enforce                  Exit non-zero if p95 exceeds target (env: FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=1)",
    "  --allow-ci                 Allow running under CI without FORMULA_RUN_DESKTOP_MEMORY_BENCH=1",
    "  -h, --help                 Show this help and exit",
    "",
    "Notes:",
    "  - Uses an isolated HOME under target/perf-home by default (override via FORMULA_PERF_HOME).",
    "  - Windows reports process-tree Working Set (closest analogue to RSS).",
    "",
  ].join("\n");
}

function parseArgs(argv: string[]): {
  runs: number;
  timeoutMs: number;
  settleMs: number;
  binPath: string | null;
  targetMb: number;
  allowInCi: boolean;
  enforce: boolean;
  jsonPath: string | null;
} {
  const args = [...argv];
  const envRuns = Number(process.env.FORMULA_DESKTOP_MEMORY_RUNS ?? "") || 10;
  const envTimeoutMs = Number(process.env.FORMULA_DESKTOP_MEMORY_TIMEOUT_MS ?? "") || 30_000;
  const envSettleMs = Number(process.env.FORMULA_DESKTOP_MEMORY_SETTLE_MS ?? "") || 5_000;

  const rawTarget =
    process.env.FORMULA_DESKTOP_IDLE_RSS_TARGET_MB ?? process.env.FORMULA_DESKTOP_MEMORY_TARGET_MB ?? "";
  const envTargetMb = Number(rawTarget) || 100;

  const envEnforce = process.env.FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH === "1";
  const envBin = process.env.FORMULA_DESKTOP_BIN ?? null;

  const out = {
    runs: Math.max(1, envRuns),
    timeoutMs: Math.max(1, envTimeoutMs),
    settleMs: Math.max(0, envSettleMs),
    binPath: envBin as string | null,
    targetMb: Math.max(1, envTargetMb),
    allowInCi: false,
    enforce: envEnforce,
    jsonPath: null as string | null,
  };

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;
    if (arg === "--runs" && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === "--timeout-ms" && args[0]) out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if (arg === "--settle-ms" && args[0]) out.settleMs = Math.max(0, Number(args.shift()) || out.settleMs);
    else if ((arg === "--bin" || arg === "--bin-path") && args[0]) out.binPath = args.shift()!;
    else if (arg === "--target-mb" && args[0]) {
      const raw = Number(args.shift());
      if (Number.isFinite(raw) && raw > 0) out.targetMb = raw;
    } else if (arg === "--allow-ci") out.allowInCi = true;
    else if (arg === "--enforce") out.enforce = true;
    else if ((arg === "--json" || arg === "--json-path") && args[0]) out.jsonPath = args.shift()!;
  }

  return out;
}

function closeReadline(rl: Interface | null): void {
  if (!rl) return;
  try {
    rl.close();
  } catch {
    // ignore
  }
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

function readProcChildrenPidsLinux(pid: number): number[] {
  try {
    // `/proc/<pid>/task/<pid>/children` contains whitespace-separated child PIDs.
    // (This is sufficient for our usage here since the xvfb wrapper is single-threaded.)
    const content = readFileSync(`/proc/${pid}/task/${pid}/children`, "utf8").trim();
    if (!content) return [];
    return content
      .split(/\s+/g)
      .map((token: string) => Number(token))
      .filter((n: number) => Number.isInteger(n) && n > 0);
  } catch {
    return [];
  }
}

function readProcExeLinux(pid: number): string | null {
  try {
    const target = readlinkSync(`/proc/${pid}/exe`, { encoding: "utf8" });
    return target.replace(/ \(deleted\)$/, "");
  } catch {
    return null;
  }
}

function findDesktopPidUnderWrapperLinux(wrapperPid: number, binPath: string): number | null {
  let binReal = binPath;
  try {
    binReal = realpathSync(binPath);
  } catch {
    // ignore; best-effort match.
  }

  const children = readProcChildrenPidsLinux(wrapperPid);
  for (const pid of children) {
    const exe = readProcExeLinux(pid);
    if (!exe) continue;
    if (exe === binReal || exe === binPath) return pid;
  }

  // Fallback: look for a process whose exe basename matches `formula-desktop`.
  for (const pid of children) {
    const exe = readProcExeLinux(pid);
    if (!exe) continue;
    if (exe.endsWith("/formula-desktop")) return pid;
  }

  return null;
}

function processTreeRssKb(rootPid: number): number {
  // `ps` RSS is reported in KB on both Linux and macOS.
  // BSD/mac: `ps -ax -o pid= -o ppid= -o rss=`
  // GNU/Linux: same flags work.
  const proc = spawnSync("ps", ["-ax", "-o", "pid=", "-o", "ppid=", "-o", "rss="], {
    encoding: "utf8",
    cwd: repoRoot,
    // `ps -ax` can print many lines on CI runners; bump the buffer for safety.
    maxBuffer: 5 * 1024 * 1024,
    timeout: 5000,
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

function processTreeWorkingSetBytesWindows(rootPid: number): number {
  // Aggregate in PowerShell to avoid streaming/parsing a huge JSON payload in Node.
  const script = [
    "$ErrorActionPreference = 'SilentlyContinue'",
    `$rootPid = ${rootPid}`,
    "$procs = Get-CimInstance Win32_Process | Select-Object ProcessId, ParentProcessId, WorkingSetSize",
    "$children = @{}",
    "$ws = @{}",
    "foreach ($p in $procs) {",
    "  $pid = [int]$p.ProcessId",
    "  $ppid = [int]$p.ParentProcessId",
    "  if (-not $children.ContainsKey($ppid)) { $children[$ppid] = @() }",
    "  $children[$ppid] += $pid",
    "  $ws[$pid] = [int64]$p.WorkingSetSize",
    "}",
    "$stack = New-Object System.Collections.Generic.Stack[int]",
    "$seen = New-Object System.Collections.Generic.HashSet[int]",
    "$stack.Push($rootPid)",
    "$total = [int64]0",
    "while ($stack.Count -gt 0) {",
    "  $pid = $stack.Pop()",
    "  if (-not $seen.Add($pid)) { continue }",
    "  if ($ws.ContainsKey($pid)) { $total += $ws[$pid] }",
    "  if ($children.ContainsKey($pid)) { foreach ($c in $children[$pid]) { $stack.Push([int]$c) } }",
    "}",
    "Write-Output $total",
  ].join("\n");

  const proc = spawnSync("powershell", ["-NoProfile", "-Command", script], {
    encoding: "utf8",
    cwd: repoRoot,
    maxBuffer: 1024 * 1024,
    timeout: 15000,
  });
  if (proc.error) throw proc.error;
  if (proc.status !== 0) {
    throw new Error(`powershell memory sampling failed (exit ${proc.status}):\n${proc.stderr}`);
  }
  const stdout = (proc.stdout ?? "").trim();
  if (!stdout) return 0;
  const bytes = Number(stdout);
  if (!Number.isFinite(bytes) || bytes < 0) return 0;
  return bytes;
}

function processTreeMemoryMb(rootPid: number): number {
  if (process.platform === "win32") {
    return processTreeWorkingSetBytesWindows(rootPid) / (1024 * 1024);
  }
  return processTreeRssKb(rootPid) / 1024;
}

async function runOnce(binPath: string, timeoutMs: number, settleMs: number): Promise<number> {
  if (process.env.FORMULA_DESKTOP_BENCH_RESET_HOME === "1") {
    const rootDir = parse(perfHome).root;
    if (perfHome === rootDir || perfHome === repoRoot) {
      throw new Error(`Refusing to reset unsafe desktop benchmark home dir: ${perfHome}`);
    }
    const safeRoot = resolve(repoRoot, "target");
    const allowUnsafe =
      process.env.FORMULA_PERF_ALLOW_UNSAFE_CLEAN === "1" ||
      String(process.env.FORMULA_PERF_ALLOW_UNSAFE_CLEAN ?? "")
        .trim()
        .toLowerCase() === "true";
    if (!isSubpath(safeRoot, perfHome) && !allowUnsafe) {
      throw new Error(
        `Refusing to reset FORMULA_PERF_HOME outside ${safeRoot} (got ${perfHome}).\n` +
          "Pick a path under target/ (recommended), or set FORMULA_PERF_ALLOW_UNSAFE_CLEAN=1 to override (DANGEROUS).",
      );
    }
    rmSync(perfHome, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 });
  }
  mkdirSync(perfHome, { recursive: true });
  mkdirSync(perfTmp, { recursive: true });
  mkdirSync(perfXdgConfig, { recursive: true });
  mkdirSync(perfXdgCache, { recursive: true });
  mkdirSync(perfXdgState, { recursive: true });
  mkdirSync(perfXdgData, { recursive: true });
  mkdirSync(perfAppData, { recursive: true });
  mkdirSync(perfLocalAppData, { recursive: true });

  const useXvfb = shouldUseXvfb();
  const xvfbPath = resolve(repoRoot, "scripts/xvfb-run-safe.sh");
  const command = useXvfb ? "bash" : binPath;
  const args = useXvfb ? [xvfbPath, binPath] : [];

  const env = {
    ...process.env,
    // Keep perf benchmarks stable/quiet by disabling the automatic startup update check.
    FORMULA_DISABLE_STARTUP_UPDATE_CHECK: "1",
    // Enable the Rust-side single-line log in release builds.
    FORMULA_STARTUP_METRICS: "1",
    // Optional: allow downstream tooling to discover the chosen HOME root.
    FORMULA_PERF_HOME: perfHome,
    // In case the app reads $HOME / XDG dirs for config, keep per-run caches out of the real home dir.
    HOME: perfHome,
    USERPROFILE: perfHome,
    XDG_CONFIG_HOME: perfXdgConfig,
    XDG_CACHE_HOME: perfXdgCache,
    XDG_STATE_HOME: perfXdgState,
    XDG_DATA_HOME: perfXdgData,
    APPDATA: perfAppData,
    LOCALAPPDATA: perfLocalAppData,
    TMPDIR: perfTmp,
    TEMP: perfTmp,
    TMP: perfTmp,
  } satisfies NodeJS.ProcessEnv;

  return await new Promise<number>((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: ["ignore", "pipe", "pipe"],
      env,
      // On POSIX, start the app in its own process group so we can terminate the whole tree.
      detached: process.platform !== "win32",
      windowsHide: true,
    });

    let rlOut: Interface | null = null;
    let rlErr: Interface | null = null;

    let settled = false;
    let captured: StartupMetrics | null = null;
    let sampledRssMb: number | null = null;

    let startupTimeout: NodeJS.Timeout | null = null;
    let settleTimer: NodeJS.Timeout | null = null;
    let forceKillTimer: NodeJS.Timeout | null = null;
    let exitDeadline: NodeJS.Timeout | null = null;
    let timedOutWaitingForMetrics = false;

    const cleanup = () => {
      if (startupTimeout) clearTimeout(startupTimeout);
      if (settleTimer) clearTimeout(settleTimer);
      if (forceKillTimer) clearTimeout(forceKillTimer);
      if (exitDeadline) clearTimeout(exitDeadline);
      closeReadline(rlOut);
      closeReadline(rlErr);
    };

    const settle = (kind: "resolve" | "reject", value: any) => {
      if (settled) return;
      settled = true;
      cleanup();
      if (kind === "resolve") resolvePromise(value);
      else rejectPromise(value);
    };

    const beginShutdown = (reason: "sampled" | "timeout") => {
      if (exitDeadline) return;

      const initialMode: TerminateProcessTreeMode =
        process.platform === "win32" || reason === "timeout" ? "force" : "graceful";

      terminateProcessTree(child, initialMode);
      forceKillTimer = setTimeout(() => terminateProcessTree(child, "force"), 2000);
      exitDeadline = setTimeout(() => {
        terminateProcessTree(child, "force");

        // Extremely defensive: don't hang the parent process even if kill fails.
        try {
          child.unref();
        } catch {
          // ignore
        }
        try {
          child.stdout?.destroy();
        } catch {
          // ignore
        }
        try {
          child.stderr?.destroy();
        } catch {
          // ignore
        }

        const msg =
          reason === "sampled"
            ? "Timed out waiting for desktop process to exit after sampling memory"
            : "Timed out waiting for desktop process to exit after timing out waiting for startup metrics";
        settle("reject", new Error(msg));
      }, 5000);
    };

    const onLine = (line: string) => {
      if (captured || timedOutWaitingForMetrics) return;
      const parsed = parseStartupLine(line);
      if (!parsed) return;
      captured = parsed;
      if (startupTimeout) {
        clearTimeout(startupTimeout);
        startupTimeout = null;
      }

      settleTimer = setTimeout(() => {
        try {
          const wrapperPid = child.pid;
          if (!wrapperPid || wrapperPid <= 0) {
            throw new Error("Desktop process PID was not available for memory sampling");
          }
          let rootPid = wrapperPid;
          // When running under the xvfb wrapper, the spawned process is a bash script
          // that also owns the Xvfb server. To keep the reported RSS scoped to the
          // desktop app (and its WebView children), locate the actual desktop PID and
          // measure its process tree instead of the wrapper's.
          if (process.platform === "linux" && useXvfb) {
            const found = findDesktopPidUnderWrapperLinux(wrapperPid, binPath);
            if (found) rootPid = found;
          }
          sampledRssMb = processTreeMemoryMb(rootPid);
        } catch (err) {
          terminateProcessTree(child, "force");
          try {
            child.unref();
          } catch {
            // ignore
          }
          try {
            child.stdout?.destroy();
          } catch {
            // ignore
          }
          try {
            child.stderr?.destroy();
          } catch {
            // ignore
          }
          settle("reject", err instanceof Error ? err : new Error(String(err)));
          return;
        }
        beginShutdown("sampled");
      }, settleMs);
    };

    if (child.stdout) {
      rlOut = createInterface({ input: child.stdout });
      rlOut.on("line", onLine);
    }
    if (child.stderr) {
      rlErr = createInterface({ input: child.stderr });
      rlErr.on("line", onLine);
    }

    startupTimeout = setTimeout(() => {
      timedOutWaitingForMetrics = true;
      beginShutdown("timeout");
    }, timeoutMs);

    child.on("error", (err) => {
      settle("reject", err);
    });

    // Use `close` (not `exit`) so stdout/stderr are fully drained before we decide whether we
    // observed the `[startup] ...` line. This keeps error reporting stable even if the desktop
    // process exits quickly after logging.
    child.on("close", (code, signal) => {
      if (settled) return;

      if (timedOutWaitingForMetrics) {
        settle("reject", new Error(`Timed out after ${timeoutMs}ms waiting for startup metrics`));
        return;
      }

      if (sampledRssMb != null) {
        settle("resolve", sampledRssMb);
        return;
      }

      // If the desktop process exits early (before we sample memory), still attempt to kill the
      // full process tree. WebView helpers can survive parent crashes and leak across runs.
      terminateProcessTree(child, "force");

      if (captured) {
        settle(
          "reject",
          new Error(`Desktop process exited before memory could be sampled (code=${code}, signal=${signal})`),
        );
        return;
      }

      settle(
        "reject",
        new Error(`Desktop process exited before reporting startup metrics (code=${code}, signal=${signal})`),
      );
    });
  });
}

function printSummary(summary: Summary): void {
  const status = summary.rssMb.p95 <= summary.rssMb.targetMb ? "PASS" : "FAIL";
  const measurement = process.platform === "win32" ? "working_set" : "rss";
  // eslint-disable-next-line no-console
  console.log(
    [
      "[desktop-memory]",
      `runs=${summary.runs}`,
      `idleRssMb(${status} p50=${summary.rssMb.p50.toFixed(1)}MB,p95=${summary.rssMb.p95.toFixed(1)}MB,target=${summary.rssMb.targetMb}MB)`,
      `kind=${measurement}`,
      summary.enforce ? "enforced=1" : "enforced=0",
    ].join(" "),
  );
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  if (argv.includes("--help") || argv.includes("-h")) {
    // eslint-disable-next-line no-console
    console.log(usage());
    return;
  }

  const { runs, timeoutMs, settleMs, binPath: argBin, targetMb, allowInCi, enforce, jsonPath } = parseArgs(argv);

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
      "Desktop binary not found. Build it via `bash scripts/cargo_agent.sh build -p formula-desktop-tauri --bin formula-desktop --release --features desktop` and pass --bin <path> (or set FORMULA_DESKTOP_BIN).",
    );
  }

  const memoryKind = process.platform === "win32" ? "Working Set" : "RSS";
  // eslint-disable-next-line no-console
  console.log(
    `[desktop-memory] measuring idle memory for the desktop app (${memoryKind} after TTI).\n` +
      `- runs: ${runs} (override via --runs or FORMULA_DESKTOP_MEMORY_RUNS)\n` +
      `- timeout: ${timeoutMs}ms (override via --timeout-ms or FORMULA_DESKTOP_MEMORY_TIMEOUT_MS)\n` +
      `- settle: ${settleMs}ms (override via --settle-ms or FORMULA_DESKTOP_MEMORY_SETTLE_MS)\n` +
      `- target: ${targetMb}MB (override via --target-mb or FORMULA_DESKTOP_IDLE_RSS_TARGET_MB)\n` +
      `- home: ${perfHome} (repo-local; override with FORMULA_PERF_HOME)\n` +
      (enforce
        ? "- enforcement: enabled (set FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=0 to disable)\n"
        : "- enforcement: disabled (set FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=1 or pass --enforce to fail on regression)\n"),
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

  const sorted = [...results].sort((a, b) => a - b);
  const summary: Summary = {
    runs: results.length,
    rssMb: {
      p50: percentile(sorted, 0.5),
      p95: percentile(sorted, 0.95),
      targetMb,
    },
    enforce,
  };

  printSummary(summary);

  if (jsonPath) {
    const outputPath = resolve(jsonPath);
    mkdirSync(dirname(outputPath), { recursive: true });
    const measurement = process.platform === "win32" ? "working_set" : "rss";
    writeFileSync(
      outputPath,
      JSON.stringify(
        {
          generatedAt: new Date().toISOString(),
          platform: process.platform,
          // On Windows we record process-tree Working Set (not true RSS). On Unix we record RSS.
          // Keeping this explicit in the JSON helps cross-platform comparisons.
          measurement,
          binPath,
          runs: results.length,
          settleMs,
          targetMb,
          samples: results,
          summary,
        },
        null,
        2,
      ),
      "utf8",
    );
  }

  if (enforce && summary.rssMb.p95 > summary.rssMb.targetMb) {
    process.exitCode = 1;
  }
}

await main();
