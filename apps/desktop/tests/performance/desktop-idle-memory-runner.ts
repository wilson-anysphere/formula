import { spawn, spawnSync, type ChildProcess } from "node:child_process";
import { existsSync, mkdirSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { readFile, readlink, readdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { createInterface, type Interface } from "node:readline";
import { fileURLToPath } from "node:url";

import {
  defaultDesktopBinPath,
  parseStartupLine,
  percentile,
  shouldUseXvfb,
  type StartupMetrics,
} from "./desktopStartupRunnerShared.ts";

// Ensure paths are rooted at repo root even when invoked from elsewhere.
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "../../../..");

const perfHome = resolve(repoRoot, "target", "perf-home");
const perfTmp = resolve(perfHome, "tmp");
const perfXdgConfig = resolve(perfHome, "xdg-config");
const perfXdgCache = resolve(perfHome, "xdg-cache");
const perfXdgState = resolve(perfHome, "xdg-state");
const perfXdgData = resolve(perfHome, "xdg-data");
const perfAppData = resolve(perfHome, "AppData", "Roaming");
const perfLocalAppData = resolve(perfHome, "AppData", "Local");

type MemorySample = {
  rssMb: number;
  startup: StartupMetrics;
};

type Summary = {
  runs: number;
  idleRssMb: { p50: number; p95: number };
};

function parseArgs(argv: string[]): {
  runs: number;
  timeoutMs: number;
  idleWaitMs: number;
  binPath: string | null;
  allowInCi: boolean;
  jsonPath: string | null;
} {
  const args = [...argv];
  const out = {
    runs: 5,
    timeoutMs: 30_000,
    idleWaitMs: 5_000,
    binPath: null as string | null,
    allowInCi: false,
    jsonPath: null as string | null,
  };

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;
    if (arg === "--runs" && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === "--timeout-ms" && args[0])
      out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if (arg === "--idle-wait-ms" && args[0])
      out.idleWaitMs = Math.max(0, Number(args.shift()) || out.idleWaitMs);
    else if ((arg === "--bin" || arg === "--bin-path") && args[0]) out.binPath = args.shift()!;
    else if ((arg === "--json" || arg === "--json-path") && args[0]) out.jsonPath = args.shift()!;
    else if (arg === "--allow-ci") out.allowInCi = true;
  }

  return out;
}

function printSummary(summary: Summary): void {
  // eslint-disable-next-line no-console
  console.log(
    [
      "[desktop-idle-memory]",
      `runs=${summary.runs}`,
      `idleRssMb(p50=${summary.idleRssMb.p50}mb,p95=${summary.idleRssMb.p95}mb)`,
    ].join(" "),
  );
}

async function sleep(ms: number): Promise<void> {
  await new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

function closeReadline(rl: Interface | null): void {
  if (!rl) return;
  try {
    rl.close();
  } catch {
    // ignore
  }
}

function terminate(child: ChildProcess): void {
  if (!child.pid) return;

  if (process.platform !== "win32") {
    try {
      process.kill(-child.pid, "SIGTERM");
      return;
    } catch {
      // ignore
    }
  }

  try {
    child.kill();
  } catch {
    // ignore
  }
}

function forceKill(child: ChildProcess): void {
  if (!child.pid) return;

  if (process.platform === "win32") {
    try {
      spawnSync("taskkill", ["/PID", String(child.pid), "/T", "/F"], { stdio: "ignore" });
      return;
    } catch {
      // Fall through to best-effort `child.kill()`.
    }
  }

  if (process.platform !== "win32") {
    try {
      process.kill(-child.pid, "SIGKILL");
      return;
    } catch {
      // ignore
    }
  }

  try {
    child.kill("SIGKILL");
  } catch {
    try {
      child.kill();
    } catch {
      // ignore
    }
  }
}

async function waitForExit(child: ChildProcess, timeoutMs: number): Promise<void> {
  if (child.exitCode !== null || child.signalCode !== null) return;
  await new Promise<void>((resolvePromise, rejectPromise) => {
    const deadline = setTimeout(() => {
      cleanup();
      rejectPromise(new Error(`Timed out after ${timeoutMs}ms waiting for desktop process tree to exit`));
    }, timeoutMs);

    const onExit = () => {
      cleanup();
      resolvePromise();
    };

    const cleanup = () => {
      clearTimeout(deadline);
      child.off("exit", onExit);
    };

    child.on("exit", onExit);
  });
}

async function waitForStartupMetrics(child: ChildProcess, timeoutMs: number): Promise<StartupMetrics> {
  return await new Promise<StartupMetrics>((resolvePromise, rejectPromise) => {
    let settled = false;
    let rlOut: Interface | null = null;
    let rlErr: Interface | null = null;

    const deadline = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();
      rejectPromise(new Error(`Timed out after ${timeoutMs}ms waiting for [startup] metrics log line`));
    }, timeoutMs);

    const onLine = (line: string) => {
      if (settled) return;
      const parsed = parseStartupLine(line);
      if (!parsed) return;
      settled = true;
      cleanup();
      resolvePromise(parsed);
    };

    const onExit = (code: number | null, signal: NodeJS.Signals | null) => {
      if (settled) return;
      settled = true;
      cleanup();
      rejectPromise(
        new Error(`Desktop process exited before reporting startup metrics (code=${code}, signal=${signal})`),
      );
    };

    const onError = (err: Error) => {
      if (settled) return;
      settled = true;
      cleanup();
      rejectPromise(err);
    };

    rlOut = createInterface({ input: child.stdout! });
    rlErr = createInterface({ input: child.stderr! });
    rlOut.on("line", onLine);
    rlErr.on("line", onLine);
    child.on("exit", onExit);
    child.on("error", onError);

    const cleanup = () => {
      clearTimeout(deadline);
      closeReadline(rlOut);
      closeReadline(rlErr);
      child.off("exit", onExit);
      child.off("error", onError);
    };
  });
}

function parseProcChildrenPids(content: string): number[] {
  const trimmed = content.trim();
  if (!trimmed) return [];
  return trimmed
    .split(/\s+/g)
    .map((x) => Number(x))
    .filter((n) => Number.isInteger(n) && n > 0);
}

function parseProcStatusVmRssKb(content: string): number | null {
  const match = content.match(/^VmRSS:\s+(\d+)\s+kB\s*$/m);
  if (!match) return null;
  const kb = Number(match[1]);
  if (!Number.isFinite(kb)) return null;
  return kb;
}

async function readUtf8(path: string): Promise<string | null> {
  try {
    return await readFile(path, "utf8");
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT" || code === "ESRCH" || code === "EACCES") return null;
    throw err;
  }
}

async function readProcExe(pid: number): Promise<string | null> {
  try {
    const target = await readlink(`/proc/${pid}/exe`);
    return target.replace(/ \(deleted\)$/, "");
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT" || code === "ESRCH" || code === "EACCES") return null;
    throw err;
  }
}

async function getChildPidsLinux(pid: number): Promise<number[]> {
  let tids: string[];
  try {
    tids = await readdir(`/proc/${pid}/task`);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT" || code === "ESRCH" || code === "EACCES") return [];
    throw err;
  }

  const out = new Set<number>();
  for (const tid of tids) {
    const content = await readUtf8(`/proc/${pid}/task/${tid}/children`);
    if (!content) continue;
    for (const child of parseProcChildrenPids(content)) out.add(child);
  }

  return [...out];
}

async function collectProcessTreePidsLinux(rootPid: number): Promise<number[]> {
  const seen = new Set<number>();
  const stack: number[] = [rootPid];
  while (stack.length > 0) {
    const pid = stack.pop()!;
    if (seen.has(pid)) continue;
    seen.add(pid);
    const children = await getChildPidsLinux(pid);
    for (const child of children) {
      if (!seen.has(child)) stack.push(child);
    }
  }
  return [...seen];
}

async function getProcessRssBytesLinux(pid: number): Promise<number> {
  const status = await readUtf8(`/proc/${pid}/status`);
  if (!status) return 0;
  const kb = parseProcStatusVmRssKb(status);
  if (!kb) return 0;
  return kb * 1024;
}

async function getProcessTreeRssBytesLinux(rootPid: number): Promise<number> {
  const pids = await collectProcessTreePidsLinux(rootPid);
  let total = 0;
  for (const pid of pids) {
    total += await getProcessRssBytesLinux(pid);
  }
  return total;
}

async function findPidForExecutableLinux(
  rootPid: number,
  binPath: string,
  timeoutMs: number,
): Promise<number | null> {
  const deadline = Date.now() + timeoutMs;
  const binResolved = resolve(binPath);
  // Best-effort resolve of symlinks so we match `/proc/<pid>/exe` consistently.
  let binReal = binResolved;
  try {
    binReal = realpathSync(binResolved);
  } catch {
    // ignore
  }

  while (Date.now() < deadline) {
    const pids = await collectProcessTreePidsLinux(rootPid);
    for (const pid of pids) {
      const exe = await readProcExe(pid);
      if (!exe) continue;
      if (exe === binReal || exe === binResolved) return pid;
    }
    await sleep(50);
  }
  return null;
}

function getProcessTreeRssBytesDarwin(rootPid: number): number {
  const proc = spawnSync("ps", ["-axo", "pid=,ppid=,rss="], { encoding: "utf8" });
  if (proc.error || proc.status !== 0) return 0;

  const childrenByPpid = new Map<number, number[]>();
  const rssKbByPid = new Map<number, number>();

  for (const rawLine of proc.stdout.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    const parts = line.split(/\s+/);
    if (parts.length < 3) continue;
    const pid = Number(parts[0]);
    const ppid = Number(parts[1]);
    const rssKb = Number(parts[2]);
    if (!Number.isInteger(pid) || pid <= 0) continue;
    if (!Number.isInteger(ppid) || ppid < 0) continue;
    if (!Number.isFinite(rssKb) || rssKb < 0) continue;

    rssKbByPid.set(pid, rssKb);
    const children = childrenByPpid.get(ppid);
    if (children) children.push(pid);
    else childrenByPpid.set(ppid, [pid]);
  }

  const seen = new Set<number>();
  const stack: number[] = [rootPid];
  while (stack.length > 0) {
    const pid = stack.pop()!;
    if (seen.has(pid)) continue;
    seen.add(pid);
    const children = childrenByPpid.get(pid) ?? [];
    for (const child of children) {
      if (!seen.has(child)) stack.push(child);
    }
  }

  let totalKb = 0;
  for (const pid of seen) {
    totalKb += rssKbByPid.get(pid) ?? 0;
  }
  return totalKb * 1024;
}

function getProcessTreeWorkingSetBytesWindows(rootPid: number): number {
  const script = [
    "$ErrorActionPreference = 'SilentlyContinue'",
    "$procs = Get-CimInstance Win32_Process | Select-Object ProcessId, ParentProcessId, WorkingSetSize",
    "$procs | ConvertTo-Json -Compress",
  ].join("; ");
  const proc = spawnSync("powershell", ["-NoProfile", "-Command", script], { encoding: "utf8" });
  if (proc.error || proc.status !== 0) return 0;
  const stdout = (proc.stdout ?? "").trim();
  if (!stdout) return 0;

  let parsed: unknown;
  try {
    parsed = JSON.parse(stdout);
  } catch {
    return 0;
  }

  const rows = Array.isArray(parsed) ? parsed : [parsed];
  const childrenByPpid = new Map<number, number[]>();
  const wsByPid = new Map<number, number>();

  for (const row of rows) {
    if (!row || typeof row !== "object") continue;
    const pid = Number((row as any).ProcessId);
    const ppid = Number((row as any).ParentProcessId);
    const ws = Number((row as any).WorkingSetSize);
    if (!Number.isInteger(pid) || pid <= 0) continue;
    if (!Number.isInteger(ppid) || ppid < 0) continue;
    if (!Number.isFinite(ws) || ws < 0) continue;

    wsByPid.set(pid, ws);
    const children = childrenByPpid.get(ppid);
    if (children) children.push(pid);
    else childrenByPpid.set(ppid, [pid]);
  }

  const seen = new Set<number>();
  const stack: number[] = [rootPid];
  while (stack.length > 0) {
    const pid = stack.pop()!;
    if (seen.has(pid)) continue;
    seen.add(pid);
    const children = childrenByPpid.get(pid) ?? [];
    for (const child of children) {
      if (!seen.has(child)) stack.push(child);
    }
  }

  let total = 0;
  for (const pid of seen) {
    total += wsByPid.get(pid) ?? 0;
  }
  return total;
}

async function getProcessTreeRssBytes(rootPid: number): Promise<number> {
  if (process.platform === "linux") {
    return await getProcessTreeRssBytesLinux(rootPid);
  }
  if (process.platform === "darwin") {
    return getProcessTreeRssBytesDarwin(rootPid);
  }
  if (process.platform === "win32") {
    return getProcessTreeWorkingSetBytesWindows(rootPid);
  }
  return 0;
}

async function runOnce({
  binPath,
  timeoutMs,
  idleWaitMs,
}: {
  binPath: string;
  timeoutMs: number;
  idleWaitMs: number;
}): Promise<MemorySample> {
  // Best-effort isolation: keep the desktop app from mutating a developer's real home directory.
  if (process.env.FORMULA_DESKTOP_BENCH_RESET_HOME === "1") {
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

  const child = spawn(command, args, {
    cwd: repoRoot,
    stdio: ["ignore", "pipe", "pipe"],
    env: {
      ...process.env,
      FORMULA_DISABLE_STARTUP_UPDATE_CHECK: "1",
      FORMULA_STARTUP_METRICS: "1",
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
    },
    detached: process.platform !== "win32",
    windowsHide: true,
  });

  if (!child.pid) {
    child.kill();
    throw new Error("Failed to spawn desktop process (missing pid)");
  }

  try {
    const startup = await waitForStartupMetrics(child, timeoutMs);
    if (idleWaitMs > 0) await sleep(idleWaitMs);

    let rootPid = child.pid;
    if (process.platform === "linux" && useXvfb) {
      const found =
        (await findPidForExecutableLinux(child.pid, binPath, Math.min(2000, timeoutMs))) ?? null;
      if (found) rootPid = found;
    }

    const rssBytes = await getProcessTreeRssBytes(rootPid);
    return { rssMb: rssBytes / (1024 * 1024), startup };
  } finally {
    try {
      terminate(child);
      await waitForExit(child, 5000);
    } catch {
      // ignore
    } finally {
      forceKill(child);
      await waitForExit(child, 5000).catch(() => {});
    }
  }
}

async function main(): Promise<void> {
  const { runs, timeoutMs, idleWaitMs, binPath: argBin, allowInCi, jsonPath } = parseArgs(
    process.argv.slice(2),
  );

  if (process.env.CI && !allowInCi && process.env.FORMULA_RUN_DESKTOP_MEMORY_BENCH !== "1") {
    // eslint-disable-next-line no-console
    console.log(
      "[desktop-idle-memory] skipping in CI (set FORMULA_RUN_DESKTOP_MEMORY_BENCH=1 or pass --allow-ci to run)",
    );
    return;
  }

  const binPath = argBin ? resolve(argBin) : defaultDesktopBinPath();
  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      "Desktop binary not found. Build it via `(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)` and pass --bin <path>.",
    );
  }

  const samples: MemorySample[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-idle-memory] run ${i + 1}/${runs}...`);
    samples.push(await runOnce({ binPath, timeoutMs, idleWaitMs }));
  }

  const rssSorted = samples.map((s) => s.rssMb).sort((a, b) => a - b);
  const summary: Summary = {
    runs: samples.length,
    idleRssMb: {
      p50: percentile(rssSorted, 0.5),
      p95: percentile(rssSorted, 0.95),
    },
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
          runs: samples.length,
          idleWaitMs,
          samples,
          summary,
        },
        null,
        2,
      ),
      "utf8",
    );
  }
}

await main();
