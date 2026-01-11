import { spawn } from "node:child_process";
import type { ChildProcessWithoutNullStreams } from "node:child_process";
import type { AddressInfo } from "node:net";
import net from "node:net";
import path from "node:path";
import { fileURLToPath } from "node:url";

export async function waitForCondition(
  condition: () => boolean | Promise<boolean>,
  timeoutMs: number,
  intervalMs: number = 25
): Promise<void> {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (await condition()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

export async function getAvailablePort(): Promise<number> {
  return await new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address() as AddressInfo;
      const port = address.port;
      server.close(() => resolve(port));
    });
  });
}

export async function waitForServerReady(baseUrl: string): Promise<void> {
  await waitForCondition(async () => {
    try {
      const res = await fetch(`${baseUrl}/healthz`);
      return res.ok;
    } catch {
      return false;
    }
  }, 30_000);
}

export function waitForProviderSync(provider: {
  on: (event: string, cb: (...args: any[]) => void) => void;
  off: (event: string, cb: (...args: any[]) => void) => void;
}): Promise<void> {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      provider.off("sync", handler);
      reject(new Error("Timed out waiting for provider sync"));
    }, 10_000);
    timeout.unref();

    const handler = (isSynced: boolean) => {
      if (!isSynced) return;
      clearTimeout(timeout);
      provider.off("sync", handler);
      resolve();
    };
    provider.on("sync", handler);
  });
}

export type SyncServerAuthConfig =
  | {
      mode: "opaque";
      token: string;
    }
  | {
      mode: "jwt";
      secret: string;
      audience?: string;
      issuer?: string;
    };

export type StartedSyncServer = {
  port: number;
  httpUrl: string;
  wsUrl: string;
  stop: () => Promise<void>;
};

export async function startSyncServer(opts: {
  port?: number;
  dataDir: string;
  auth: SyncServerAuthConfig;
  env?: Record<string, string | undefined>;
}): Promise<StartedSyncServer> {
  const port = opts.port ?? (await getAvailablePort());
  const httpUrl = `http://127.0.0.1:${port}`;
  const wsUrl = `ws://127.0.0.1:${port}`;

  const serviceDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    ".."
  );
  const entry = path.join(serviceDir, "src", "index.ts");
  const nodeWithTsx = path.join(serviceDir, "scripts", "node-with-tsx.mjs");

  let child: ChildProcessWithoutNullStreams | null = null;
  let stdout = "";
  let stderr = "";

  const killChild = async () => {
    const proc = child;
    if (!proc) return;
    child = null;
    proc.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      proc.once("exit", () => resolve());
    });
  };

  const authEnv: Record<string, string> =
    opts.auth.mode === "opaque"
      ? {
          SYNC_SERVER_AUTH_TOKEN: opts.auth.token,
          SYNC_SERVER_JWT_SECRET: "",
          SYNC_SERVER_JWT_AUDIENCE: "",
          SYNC_SERVER_JWT_ISSUER: "",
        }
      : {
          SYNC_SERVER_AUTH_TOKEN: "",
          SYNC_SERVER_JWT_SECRET: opts.auth.secret,
          SYNC_SERVER_JWT_AUDIENCE: opts.auth.audience ?? "formula-sync",
          SYNC_SERVER_JWT_ISSUER: opts.auth.issuer ?? "",
        };

  child = spawn(process.execPath, [nodeWithTsx, entry], {
    cwd: serviceDir,
    env: {
      ...process.env,
      NODE_ENV: "test",
      LOG_LEVEL: "silent",
      SYNC_SERVER_HOST: "127.0.0.1",
      SYNC_SERVER_PORT: String(port),
      SYNC_SERVER_DATA_DIR: opts.dataDir,
      SYNC_SERVER_PERSISTENCE_BACKEND: "file",
      SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "10",
      ...authEnv,
      ...(opts.env ?? {}),
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  child.stdout.on("data", (d) => {
    stdout += d.toString();
    stdout = stdout.slice(-10_000);
  });
  child.stderr.on("data", (d) => {
    stderr += d.toString();
    stderr = stderr.slice(-10_000);
  });

  try {
    await waitForServerReady(httpUrl);
  } catch (err) {
    await killChild();
    throw new Error(
      `Server failed to start: ${String(err)}\nstdout:\n${stdout}\nstderr:\n${stderr}`
    );
  }

  return {
    port,
    httpUrl,
    wsUrl,
    stop: killChild,
  };
}
