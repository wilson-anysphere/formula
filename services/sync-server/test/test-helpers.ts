import { spawn } from "node:child_process";
import type { ChildProcessWithoutNullStreams } from "node:child_process";
import http from "node:http";
import https from "node:https";
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
    const url = new URL("/healthz", baseUrl);
    try {
      const ok = await new Promise<boolean>((resolve) => {
        const cb = (res: http.IncomingMessage) => {
          res.resume();
          resolve(Boolean(res.statusCode && res.statusCode >= 200 && res.statusCode < 300));
        };

        const req =
          url.protocol === "https:"
            ? https.request(
                {
                  protocol: url.protocol,
                  hostname: url.hostname,
                  port: url.port,
                  path: `${url.pathname}${url.search}`,
                  method: "GET",
                  rejectUnauthorized: false,
                },
                cb
              )
            : http.request(
                {
                  protocol: url.protocol,
                  hostname: url.hostname,
                  port: url.port,
                  path: `${url.pathname}${url.search}`,
                  method: "GET",
                },
                cb
              );
        req.on("error", () => resolve(false));
        req.end();
      });
      return ok;
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
  const tlsEnabled = Boolean(
    opts.env?.SYNC_SERVER_TLS_CERT_PATH && opts.env?.SYNC_SERVER_TLS_KEY_PATH
  );
  const httpScheme = tlsEnabled ? "https" : "http";
  const wsScheme = tlsEnabled ? "wss" : "ws";
  const httpUrl = `${httpScheme}://127.0.0.1:${port}`;
  const wsUrl = `${wsScheme}://127.0.0.1:${port}`;

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
      const timeout = setTimeout(() => {
        proc.kill("SIGKILL");
        resolve();
      }, 10_000);
      timeout.unref();
      proc.once("exit", () => {
        clearTimeout(timeout);
        resolve();
      });
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
      // Prevent TLS-related env vars from leaking into unrelated tests.
      SYNC_SERVER_TLS_CERT_PATH: "",
      SYNC_SERVER_TLS_KEY_PATH: "",
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
