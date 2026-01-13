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
  intervalMs: number = 25,
  signal?: AbortSignal
): Promise<void> {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (signal?.aborted) {
      const reason = (signal as any).reason;
      if (reason instanceof Error) throw reason;
      throw new Error(reason ? String(reason) : "Aborted");
    }
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

export async function waitForServerReady(
  baseUrl: string,
  opts: { signal?: AbortSignal } = {}
): Promise<void> {
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
  }, 30_000, 25, opts.signal);
}

export function waitForProviderSync(provider: {
  on: (event: string, cb: (...args: any[]) => void) => void;
  off: (event: string, cb: (...args: any[]) => void) => void;
  synced?: boolean;
}): Promise<void> {
  if (provider.synced === true) return Promise.resolve();
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

    // The provider may have already synced before we attached our listener,
    // particularly when multiple providers are created before awaiting sync.
    if (provider.synced === true) {
      clearTimeout(timeout);
      provider.off("sync", handler);
      resolve();
    }
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
    if (proc.exitCode !== null || proc.signalCode !== null) return;
    await new Promise<void>((resolve) => {
      let timeout: NodeJS.Timeout | null = null;
      const finish = () => {
        if (timeout) clearTimeout(timeout);
        timeout = null;
        resolve();
      };

      const onExit = () => {
        finish();
      };

      proc.once("exit", onExit);

      if (proc.exitCode !== null || proc.signalCode !== null) {
        proc.off("exit", onExit);
        finish();
        return;
      }

      try {
        proc.kill("SIGTERM");
      } catch {
        proc.off("exit", onExit);
        finish();
        return;
      }

      timeout = setTimeout(() => {
        proc.off("exit", onExit);
        try {
          proc.kill("SIGKILL");
        } catch {
          // ignore
        }
        finish();
      }, 10_000);
      timeout.unref();
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
      // Prevent local ops hardening flags from affecting test behavior.
      SYNC_SERVER_DISABLE_PUBLIC_METRICS: "",
      // Prevent per-document websocket connection limits from leaking into tests.
      SYNC_SERVER_MAX_CONNECTIONS_PER_DOC: "",
      // Prevent websocket origin allowlist settings from leaking into tests.
      SYNC_SERVER_ALLOWED_ORIGINS: "",
      // Prevent per-IP websocket message rate limiting env vars from leaking into tests.
      SYNC_SERVER_MAX_MESSAGES_PER_IP_WINDOW: "",
      SYNC_SERVER_IP_MESSAGE_WINDOW_MS: "",
      // Ensure tests default to immediate shutdown so the child-process harness
      // doesn't race the configured grace period.
      SYNC_SERVER_SHUTDOWN_GRACE_MS: "0",
      // Prevent JWT hardening flags from leaking into tests (individual tests can opt in).
      SYNC_SERVER_JWT_REQUIRE_SUB: "",
      SYNC_SERVER_JWT_REQUIRE_EXP: "",
      // Prevent persistence backpressure settings from leaking into tests.
      SYNC_SERVER_PERSISTENCE_MAX_QUEUE_DEPTH_PER_DOC: "",
      SYNC_SERVER_PERSISTENCE_MAX_QUEUE_DEPTH_TOTAL: "",
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

  const proc = child;
  if (!proc) {
    throw new Error("Failed to spawn sync-server process");
  }

  const readyAbortController = new AbortController();
  const onExit = (code: number | null, signal: NodeJS.Signals | null) => {
    readyAbortController.abort(
      new Error(
        `sync-server exited before becoming ready (code=${code ?? 0}, signal=${signal ?? "none"})`
      )
    );
  };

  proc.once("exit", onExit);
  if (proc.exitCode !== null || proc.signalCode !== null) {
    onExit(proc.exitCode, proc.signalCode);
  }

  try {
    await waitForServerReady(httpUrl, { signal: readyAbortController.signal });
  } catch (err) {
    await killChild();
    throw new Error(
      `Server failed to start: ${String(err)}\nstdout:\n${stdout}\nstderr:\n${stderr}`
    );
  } finally {
    proc.off("exit", onExit);
  }

  return {
    port,
    httpUrl,
    wsUrl,
    stop: killChild,
  };
}
