import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

import { buildBenchmarkResultFromValues } from './benchmark.ts';
import {
  defaultDesktopBinPath,
  findPidForExecutableLinux,
  formatPerfPath,
  getProcessTreeRssBytesLinux,
  repoRoot,
  resolvePerfHome,
  runOnce as runDesktopOnce,
  sleep,
} from './desktopStartupUtil.ts';

type Summary = {
  runs: number;
  rssMb: { p50: number; p95: number; targetMb: number };
  enforce: boolean;
};

function usage(): string {
  return [
    'Desktop idle memory benchmark runner (real Tauri binary).',
    '',
    'Usage:',
    '  node scripts/run-node-ts.mjs apps/desktop/tests/performance/desktop-memory-runner.ts [options]',
    '',
    'Options:',
    '  --runs <n>                 Iterations (env: FORMULA_DESKTOP_MEMORY_RUNS, default: 10)',
    '  --timeout-ms <ms>          Timeout per run (env: FORMULA_DESKTOP_MEMORY_TIMEOUT_MS, default: 20000)',
    '  --settle-ms <ms>           Delay after startup before sampling (env: FORMULA_DESKTOP_MEMORY_SETTLE_MS, default: 5000)',
    '  --bin, --bin-path <path>   Desktop binary path (env: FORMULA_DESKTOP_BIN)',
    '  --target-mb <mb>           p95 target (env: FORMULA_DESKTOP_IDLE_RSS_TARGET_MB, default: 100)',
    '  --json, --json-path <path> Write JSON output (samples + summary) to this path',
    '  --enforce                  Exit non-zero if p95 exceeds target (env: FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=1)',
    '  --allow-ci                 Allow running under CI without FORMULA_RUN_DESKTOP_MEMORY_BENCH=1',
    '  -h, --help                 Show this help and exit',
    '',
    'Notes:',
    '  - Uses an isolated profile directory under target/perf-home by default (override via FORMULA_PERF_HOME).',
    '    Each invocation uses a unique profile dir to avoid persistent cache pollution across runs.',
    '  - Set FORMULA_DESKTOP_BENCH_RESET_HOME=1 to delete the profile dir before each iteration.',
    '  - Windows reports process-tree Working Set (closest analogue to RSS).',
    '',
  ].join('\n');
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
  const envRuns = Number(process.env.FORMULA_DESKTOP_MEMORY_RUNS ?? '') || 10;
  const envTimeoutMs = Number(process.env.FORMULA_DESKTOP_MEMORY_TIMEOUT_MS ?? '') || 20_000;
  // Allow explicitly setting `FORMULA_DESKTOP_MEMORY_SETTLE_MS=0` to sample immediately.
  // Treat unset/blank/invalid values as the default.
  const settleRaw = process.env.FORMULA_DESKTOP_MEMORY_SETTLE_MS;
  const settleParsed = settleRaw && settleRaw.trim() !== '' ? Number(settleRaw) : 5_000;
  const envSettleMs = Number.isFinite(settleParsed) ? Math.max(0, settleParsed) : 5_000;

  const rawTarget =
    process.env.FORMULA_DESKTOP_IDLE_RSS_TARGET_MB ?? process.env.FORMULA_DESKTOP_MEMORY_TARGET_MB ?? '';
  const envTargetMb = Number(rawTarget) || 100;

  const envEnforce = process.env.FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH === '1';
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
    if (arg === '--runs' && args[0]) out.runs = Math.max(1, Number(args.shift()) || out.runs);
    else if (arg === '--timeout-ms' && args[0]) out.timeoutMs = Math.max(1, Number(args.shift()) || out.timeoutMs);
    else if (arg === '--settle-ms' && args[0]) {
      const raw = String(args.shift());
      const parsed = Number(raw);
      if (Number.isFinite(parsed)) out.settleMs = Math.max(0, parsed);
    } else if ((arg === '--bin' || arg === '--bin-path') && args[0]) out.binPath = args.shift()!;
    else if (arg === '--target-mb' && args[0]) {
      const raw = Number(args.shift());
      if (Number.isFinite(raw) && raw > 0) out.targetMb = raw;
    } else if (arg === '--allow-ci') out.allowInCi = true;
    else if (arg === '--enforce') out.enforce = true;
    else if ((arg === '--json' || arg === '--json-path') && args[0]) out.jsonPath = args.shift()!;
  }

  return out;
}

function parsePsTable(output: string): { pid: number; ppid: number; rssKb: number }[] {
  const rows: { pid: number; ppid: number; rssKb: number }[] = [];
  for (const line of output.split('\n')) {
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
  const proc = spawnSync('ps', ['-ax', '-o', 'pid=', '-o', 'ppid=', '-o', 'rss='], {
    encoding: 'utf8',
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
    '$procs = Get-CimInstance Win32_Process | Select-Object ProcessId, ParentProcessId, WorkingSetSize',
    '$children = @{}',
    '$ws = @{}',
    'foreach ($p in $procs) {',
    '  $pid = [int]$p.ProcessId',
    '  $ppid = [int]$p.ParentProcessId',
    '  if (-not $children.ContainsKey($ppid)) { $children[$ppid] = @() }',
    '  $children[$ppid] += $pid',
    '  $ws[$pid] = [int64]$p.WorkingSetSize',
    '}',
    '$stack = New-Object System.Collections.Generic.Stack[int]',
    '$seen = New-Object System.Collections.Generic.HashSet[int]',
    '$stack.Push($rootPid)',
    '$total = [int64]0',
    'while ($stack.Count -gt 0) {',
    '  $pid = $stack.Pop()',
    '  if (-not $seen.Add($pid)) { continue }',
    '  if ($ws.ContainsKey($pid)) { $total += $ws[$pid] }',
    '  if ($children.ContainsKey($pid)) { foreach ($c in $children[$pid]) { $stack.Push([int]$c) } }',
    '}',
    'Write-Output $total',
  ].join('\n');

  const proc = spawnSync('powershell.exe', ['-NoProfile', '-Command', script], {
    encoding: 'utf8',
    cwd: repoRoot,
    maxBuffer: 1024 * 1024,
    timeout: 15000,
    windowsHide: true,
  });
  if (proc.error) throw proc.error;
  if (proc.status !== 0) {
    throw new Error(`powershell memory sampling failed (exit ${proc.status}):\n${proc.stderr}`);
  }

  const stdout = (proc.stdout ?? '').trim();
  if (!stdout) return 0;
  const bytes = Number(stdout);
  if (!Number.isFinite(bytes) || bytes < 0) return 0;
  return bytes;
}

async function processTreeMemoryMb(options: {
  rootPid: number;
  binPath: string;
  timeoutMs: number;
  signal?: AbortSignal;
}): Promise<number> {
  const { rootPid, binPath, timeoutMs, signal } = options;

  if (process.platform === 'win32') {
    return processTreeWorkingSetBytesWindows(rootPid) / (1024 * 1024);
  }

  if (process.platform === 'linux') {
    const resolvedPid = await findPidForExecutableLinux(rootPid, binPath, Math.min(2000, timeoutMs), signal);
    if (!resolvedPid) {
      throw new Error('Failed to resolve desktop PID for RSS sampling');
    }

    const bytes = await getProcessTreeRssBytesLinux(resolvedPid);
    return bytes / (1024 * 1024);
  }

  return processTreeRssKb(rootPid) / 1024;
}

async function runOnce(binPath: string, timeoutMs: number, settleMs: number, profileDir: string): Promise<number> {
  let sampledRssMb: number | null = null;
  let sampleError: Error | null = null;

  await runDesktopOnce({
    binPath,
    timeoutMs,
    profileDir,
    afterCapture: async (child, _metrics, signal) => {
      try {
        if (settleMs > 0) {
          await sleep(settleMs, signal);
        }

        const pid = child.pid;
        if (!pid || pid <= 0) {
          throw new Error('Desktop process PID was not available for memory sampling');
        }

        const rssMb = await processTreeMemoryMb({ rootPid: pid, binPath, timeoutMs, signal });
        if (!Number.isFinite(rssMb) || rssMb <= 0) {
          throw new Error('Failed to sample desktop memory (process may have exited)');
        }
        sampledRssMb = rssMb;
      } catch (err) {
        sampleError = err instanceof Error ? err : new Error(String(err));
      }
    },
    // Covers settle delay + `ps`/PowerShell timeouts.
    afterCaptureTimeoutMs: settleMs + (process.platform === 'win32' ? 20_000 : 10_000),
  });

  if (sampleError) throw sampleError;
  if (sampledRssMb == null) throw new Error('Desktop memory sampling failed');
  return sampledRssMb;
}

function printSummary(summary: Summary): void {
  const status = summary.rssMb.p95 <= summary.rssMb.targetMb ? 'PASS' : 'FAIL';
  const measurement = process.platform === 'win32' ? 'working_set' : 'rss';
  // eslint-disable-next-line no-console
  console.log(
    [
      '[desktop-memory]',
      `runs=${summary.runs}`,
      `idleRssMb(${status} p50=${summary.rssMb.p50.toFixed(1)}MB,p95=${summary.rssMb.p95.toFixed(1)}MB,target=${summary.rssMb.targetMb}MB)`,
      `kind=${measurement}`,
      summary.enforce ? 'enforced=1' : 'enforced=0',
    ].join(' '),
  );
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  if (argv.includes('--help') || argv.includes('-h')) {
    // eslint-disable-next-line no-console
    console.log(usage());
    return;
  }

  const { runs, timeoutMs, settleMs, binPath: argBin, targetMb, allowInCi, enforce, jsonPath } = parseArgs(argv);

  if (process.env.CI && !allowInCi && process.env.FORMULA_RUN_DESKTOP_MEMORY_BENCH !== '1') {
    // eslint-disable-next-line no-console
    console.log(
      '[desktop-memory] skipping in CI (set FORMULA_RUN_DESKTOP_MEMORY_BENCH=1 or pass --allow-ci to run)',
    );
    return;
  }

  const binPath = argBin ? resolve(argBin) : defaultDesktopBinPath();
  if (!binPath || !existsSync(binPath)) {
    throw new Error(
      'Desktop binary not found. Build it via `bash scripts/cargo_agent.sh build -p formula-desktop-tauri --bin formula-desktop --release --features desktop` and pass --bin <path> (or set FORMULA_DESKTOP_BIN).',
    );
  }

  const memoryKind = process.platform === 'win32' ? 'Working Set' : 'RSS';
  const perfHome = resolvePerfHome();
  const profileRoot = resolve(perfHome, `desktop-memory-${Date.now()}-${process.pid}`);
  // eslint-disable-next-line no-console
  console.log(
    `[desktop-memory] measuring idle memory for the desktop app (${memoryKind} after TTI).\n` +
      `- runs: ${runs} (override via --runs or FORMULA_DESKTOP_MEMORY_RUNS)\n` +
      `- timeout: ${timeoutMs}ms (override via --timeout-ms or FORMULA_DESKTOP_MEMORY_TIMEOUT_MS)\n` +
      `- settle: ${settleMs}ms (override via --settle-ms or FORMULA_DESKTOP_MEMORY_SETTLE_MS)\n` +
      `- target: ${targetMb}MB (override via --target-mb or FORMULA_DESKTOP_IDLE_RSS_TARGET_MB)\n` +
      `- perf-home: ${formatPerfPath(perfHome)} (override with FORMULA_PERF_HOME)\n` +
      `- profile: ${formatPerfPath(profileRoot)}\n` +
      (enforce
        ? '- enforcement: enabled (set FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=0 to disable)\n'
        : '- enforcement: disabled (set FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=1 or pass --enforce to fail on regression)\n'),
  );

  const results: number[] = [];
  for (let i = 0; i < runs; i += 1) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory] run ${i + 1}/${runs}...`);
    const rss = await runOnce(binPath, timeoutMs, settleMs, profileRoot);
    results.push(rss);
    // eslint-disable-next-line no-console
    console.log(`[desktop-memory]   idleRssMb=${rss.toFixed(1)}MB`);
  }

  const stats = buildBenchmarkResultFromValues('desktop.memory.idle_rss_mb.p95', results, targetMb, 'mb');
  const summary: Summary = {
    runs: results.length,
    rssMb: {
      p50: stats.median,
      p95: stats.p95,
      targetMb,
    },
    enforce,
  };

  printSummary(summary);

  if (jsonPath) {
    const outputPath = resolve(jsonPath);
    mkdirSync(dirname(outputPath), { recursive: true });
    const measurement = process.platform === 'win32' ? 'working_set' : 'rss';
    const perfHomeRel = formatPerfPath(perfHome);
    const profileRootRel = formatPerfPath(profileRoot);
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
          perfHome,
          perfHomeRel,
          profileRoot,
          profileRootRel,
          // Backwards-compatible name: older consumers expect `profileDir`.
          profileDir: profileRoot,
          runs: results.length,
          settleMs,
          targetMb,
          samples: results,
          summary,
        },
        null,
        2,
      ),
      'utf8',
    );
  }

  if (enforce && summary.rssMb.p95 > summary.rssMb.targetMb) {
    process.exitCode = 1;
  }
}

await main();
