import { spawnSync, type SpawnSyncOptions, type ChildProcess } from 'node:child_process';

export type TerminateProcessTreeMode = 'graceful' | 'force';

export type TerminateProcessTreeDeps = {
  /**
   * Override the platform for tests.
   *
   * Defaults to `process.platform`.
   */
  platform?: NodeJS.Platform;
  /**
   * Override the `process.kill` implementation for tests.
   */
  processKill?: typeof process.kill;
  /**
   * Override `spawnSync` for tests.
   */
  spawnSync?: typeof spawnSync;
  /**
   * Timeout (ms) for the Windows `taskkill` invocation.
   *
   * This must remain bounded so benchmark cleanup never hangs.
   */
  taskkillTimeoutMs?: number;
};

export type TaskkillCommand = { command: string; args: string[] };

export function buildTaskkillCommand(pid: number): TaskkillCommand {
  return { command: 'taskkill', args: ['/PID', String(pid), '/T', '/F'] };
}

/**
 * Best-effort termination of a spawned process *and its children*.
 *
 * Why this exists:
 * - Tauri/WebView runtimes can spawn child processes (e.g. WebKit WebProcess on Linux).
 * - Killing only the root pid can leave these GUI processes orphaned across benchmark runs.
 *
 * On POSIX:
 * - Spawns should be created with `detached: true` so the process is the leader of a new process
 *   group and we can signal the whole group via `process.kill(-pid, ...)`.
 *
 * On Windows:
 * - Node's `child.kill()` only targets the root pid; `taskkill /T` is the best-effort way to tear
 *   down the full tree.
 *
 * This helper is intentionally synchronous and best-effort; callers keep their own hard deadlines.
 */
export function terminateProcessTree(
  child: Pick<ChildProcess, 'pid' | 'kill'>,
  mode: TerminateProcessTreeMode,
  deps: TerminateProcessTreeDeps = {},
): void {
  const pid = child.pid;
  if (!pid || pid <= 0) return;

  const platform = deps.platform ?? process.platform;

  if (platform === 'win32') {
    if (mode === 'force') {
      const spawnSyncImpl = deps.spawnSync ?? spawnSync;
      const { command, args } = buildTaskkillCommand(pid);
      const options: SpawnSyncOptions = {
        stdio: 'ignore',
        windowsHide: true,
        timeout: deps.taskkillTimeoutMs ?? 2000,
      };
      try {
        const result = spawnSyncImpl(command, args, options);
        // `spawnSync` only throws for argument/IPC issues; command failures (e.g. access denied)
        // are surfaced via `status`/`error`.
        if (!result.error && result.status === 0) {
          return;
        }
      } catch {
        // Fall back to killing just the root pid. This is not ideal (children can survive),
        // but it's better than leaving the full tree running if taskkill is unavailable.
      }
    }

    try {
      // Note: on Windows this is effectively a forceful termination.
      child.kill();
    } catch {
      // Ignore (pid may already be gone).
    }
    return;
  }

  const processKill = deps.processKill ?? process.kill;
  const signal: NodeJS.Signals = mode === 'force' ? 'SIGKILL' : 'SIGTERM';

  // If the process was spawned with `detached: true`, its pid is also the process group id.
  // Killing the group reliably terminates WebView child processes.
  try {
    processKill(-pid, signal);
    return;
  } catch {
    // If the caller didn't spawn the process in its own group (or the platform doesn't support
    // negative pids), fall back to killing just the root pid.
  }

  try {
    processKill(pid, signal);
  } catch {
    // Ignore (pid may already be gone).
  }
}
