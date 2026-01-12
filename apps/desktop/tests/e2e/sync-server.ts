import { spawn, type ChildProcess } from "node:child_process";
import os from "node:os";
import path from "node:path";

export function stablePortFromString(input: string, { base = 23_000, range = 2000 } = {}): number {
  // Deterministic port selection avoids collisions when multiple agent checkouts run
  // Playwright tests on the same host.
  let hash = 0;
  for (let i = 0; i < input.length; i += 1) {
    hash = (hash * 31 + input.charCodeAt(i)) >>> 0;
  }
  return base + (hash % range);
}

export async function waitForHealthz(url: string, timeoutMs = 30_000): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(url);
      if (res.ok) return;
    } catch {
      // ignore
    }
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  throw new Error(`Timed out waiting for sync server health check at ${url}`);
}

export type LocalSyncServerHandle = {
  proc: ChildProcess;
  wsUrl: string;
  port: number;
  stop: () => Promise<void>;
};

export async function startLocalSyncServer(options: { repoRoot: string; portOffset?: number }): Promise<LocalSyncServerHandle> {
  const repoRoot = options.repoRoot;
  const portOffset = options.portOffset ?? 0;
  const syncServerRoot = path.resolve(repoRoot, "services/sync-server");

  const port = stablePortFromString(repoRoot) + portOffset;
  const wsUrl = `ws://127.0.0.1:${port}`;
  const dataDir = path.join(os.tmpdir(), `formula-sync-server-e2e-${port}`);

  const proc = spawn(process.execPath, ["scripts/node-with-tsx.mjs", "src/index.ts"], {
    cwd: syncServerRoot,
    env: {
      ...process.env,
      NODE_ENV: "test",
      SYNC_SERVER_HOST: "127.0.0.1",
      SYNC_SERVER_PORT: String(port),
      // Avoid optional leveldb deps in e2e environments.
      SYNC_SERVER_PERSISTENCE_BACKEND: "file",
      SYNC_SERVER_DATA_DIR: dataDir,
    },
    stdio: "inherit",
  });

  await waitForHealthz(`http://127.0.0.1:${port}/healthz`);

  const stop = async (): Promise<void> => {
    if (proc.exitCode != null) return;
    if (!proc.killed) {
      proc.kill("SIGTERM");
    }

    // Ensure the Playwright worker doesn't hang on a still-running child process.
    await new Promise<void>((resolve) => {
      const timeout = setTimeout(() => {
        // Best-effort: force-kill if graceful shutdown hangs.
        if (proc.exitCode == null && !proc.killed) {
          try {
            proc.kill("SIGKILL");
          } catch {
            // ignore
          }
        }
        resolve();
      }, 5_000);

      proc.once("exit", () => {
        clearTimeout(timeout);
        resolve();
      });
    });
  };

  return { proc, wsUrl, port, stop };
}

