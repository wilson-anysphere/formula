import { spawn } from "node:child_process";
import type { ChildProcessWithoutNullStreams } from "node:child_process";
import net from "node:net";
import path from "node:path";
import { setTimeout as sleep } from "node:timers/promises";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";

import jwt from "jsonwebtoken";
import WebSocket from "ws";

// Keep the stress harness on the CommonJS build of Yjs + y-websocket so we don't
// end up loading both the `import` (ESM) and `require` (CJS) entrypoints in the
// same process. Mixing them triggers Yjs' "already imported" warning and can
// break `instanceof` checks inside the CRDT / provider implementations.
const require = createRequire(import.meta.url);

// eslint-disable-next-line @typescript-eslint/no-var-requires
const Y = require("yjs") as typeof import("yjs");

// eslint-disable-next-line @typescript-eslint/no-var-requires
const yWebsocket: typeof import("y-websocket") = require("y-websocket");
const { WebsocketProvider } = yWebsocket;

type AuthMode = "opaque" | "jwt";

type StressOptions = {
  clients: number;
  docs: number;
  durationMs: number;
  opsPerClient: number;
  awarenessEveryMs: number;
  maxMessageBytes: number | null;
  authMode: AuthMode;
};

type StartedSyncServer = {
  port: number;
  httpUrl: string;
  wsUrl: string;
  dataDir: string;
  stop: () => Promise<void>;
};

type StressClient = {
  index: number;
  docId: string;
  userId: string;
  token: string;
  doc: InstanceType<typeof Y.Doc>;
  provider: import("y-websocket").WebsocketProvider;
  connectedAtMs: number | null;
  syncedAtMs: number | null;
  opsAttempted: number;
  opsSucceeded: number;
};

const DEFAULTS: StressOptions = {
  clients: 50,
  docs: 1,
  durationMs: 10_000,
  opsPerClient: 200,
  awarenessEveryMs: 1_000,
  maxMessageBytes: null,
  authMode: "opaque",
};

function parseEnvInt(name: string): number | null {
  const raw = process.env[name];
  if (raw === undefined) return null;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return null;
  const parsed = Number.parseInt(trimmed, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseEnvString(name: string): string | null {
  const raw = process.env[name];
  if (raw === undefined) return null;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return null;
  return trimmed;
}

function parsePositiveInt(name: string, value: unknown, fallback: number): number {
  if (value === undefined || value === null) return fallback;
  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed) || parsed < 0) {
    throw new Error(`${name} must be a non-negative integer (got ${String(value)})`);
  }
  return parsed;
}

function formatMs(ms: number): string {
  if (!Number.isFinite(ms)) return String(ms);
  if (ms < 1000) return `${ms}ms`;
  return `${(ms / 1000).toFixed(2)}s`;
}

async function getAvailablePort(): Promise<number> {
  return await new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address();
      if (!addr || typeof addr === "string") {
        server.close(() => reject(new Error("Failed to allocate an available port")));
        return;
      }
      const port = addr.port;
      server.close(() => resolve(port));
    });
  });
}

async function waitForServerReady(baseUrl: string, timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(new URL("/healthz", baseUrl), {
        signal: AbortSignal.timeout(1_000),
      });
      res.body?.cancel().catch(() => {
        // ignore
      });
      if (res.ok) return;
    } catch {
      // ignore
    }
    await sleep(50);
  }
  throw new Error(`Timed out waiting for sync-server to become ready (${baseUrl})`);
}

async function startSyncServer(opts: {
  port?: number;
  dataDir: string;
  authMode: AuthMode;
  opaqueToken: string;
  jwtSecret: string;
  maxMessageBytes: number | null;
  maxConnectionsPerIp: number;
  maxConnections: number;
}): Promise<StartedSyncServer> {
  const port = opts.port ?? (await getAvailablePort());
  const httpUrl = `http://127.0.0.1:${port}`;
  const wsUrl = `ws://127.0.0.1:${port}`;

  const serviceDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const entry = path.join(serviceDir, "src", "index.ts");
  const nodeWithTsx = path.join(serviceDir, "scripts", "node-with-tsx.mjs");

  const authEnv: Record<string, string> =
    opts.authMode === "opaque"
      ? {
          SYNC_SERVER_AUTH_MODE: "opaque",
          SYNC_SERVER_AUTH_TOKEN: opts.opaqueToken,
          SYNC_SERVER_JWT_SECRET: "",
          SYNC_SERVER_JWT_AUDIENCE: "",
          SYNC_SERVER_JWT_ISSUER: "",
        }
      : {
          SYNC_SERVER_AUTH_MODE: "jwt-hs256",
          SYNC_SERVER_AUTH_TOKEN: "",
          SYNC_SERVER_JWT_SECRET: opts.jwtSecret,
          SYNC_SERVER_JWT_AUDIENCE: "formula-sync",
          // Intentionally empty: jsonwebtoken treats falsy issuer as "no check".
          SYNC_SERVER_JWT_ISSUER: "",
        };

  const maxMessageBytesEnv =
    opts.maxMessageBytes === null ? {} : { SYNC_SERVER_MAX_MESSAGE_BYTES: String(opts.maxMessageBytes) };

  // Defaults in sync-server are tuned for production abuse prevention, not for
  // local load tests. Lift the per-IP cap so `--clients` works out of the box.
  const limitsEnv: Record<string, string> = {
    SYNC_SERVER_MAX_CONNECTIONS: String(opts.maxConnections),
    SYNC_SERVER_MAX_CONNECTIONS_PER_IP: String(opts.maxConnectionsPerIp),
    // Avoid 429s during initial connection spikes.
    SYNC_SERVER_MAX_CONN_ATTEMPTS_PER_WINDOW: String(Math.max(60, opts.maxConnectionsPerIp * 4)),
    SYNC_SERVER_CONN_ATTEMPT_WINDOW_MS: "60000",
    // Avoid 1013s for moderate workloads; users can still dial these down.
    SYNC_SERVER_MAX_MESSAGES_PER_WINDOW: "1000000",
    SYNC_SERVER_MESSAGE_WINDOW_MS: "10000",
    SYNC_SERVER_MAX_MESSAGES_PER_DOC_WINDOW: "1000000",
    SYNC_SERVER_DOC_MESSAGE_WINDOW_MS: "10000",
  };

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

  child = spawn(process.execPath, [nodeWithTsx, entry], {
    cwd: serviceDir,
    env: {
      ...process.env,
      NODE_ENV: "test",
      LOG_LEVEL: "error",
      SYNC_SERVER_HOST: "127.0.0.1",
      SYNC_SERVER_PORT: String(port),
      SYNC_SERVER_DATA_DIR: opts.dataDir,
      SYNC_SERVER_PERSISTENCE_BACKEND: "file",
      SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "200",
      // Clear optional features that are commonly set in dev shells so the
      // stress harness is repeatable.
      SYNC_SERVER_INTROSPECTION_URL: "",
      SYNC_SERVER_INTROSPECTION_TOKEN: "",
      SYNC_SERVER_INTROSPECTION_CACHE_TTL_MS: "",
      SYNC_SERVER_INTROSPECT_URL: "",
      SYNC_SERVER_INTROSPECT_TOKEN: "",
      SYNC_SERVER_INTROSPECT_CACHE_MS: "",
      SYNC_SERVER_INTROSPECT_FAIL_OPEN: "",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "",
      SYNC_SERVER_RETENTION_TTL_MS: "",
      SYNC_SERVER_RETENTION_SWEEP_INTERVAL_MS: "",
      SYNC_SERVER_TOMBSTONE_TTL_MS: "",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: "",
      SYNC_SERVER_ENCRYPTION_KEYRING_PATH: "",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: "",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT: "",
      SYNC_SERVER_LEVELDB_DOCNAME_HASHING: "",
      // Keep metrics public so the harness can snapshot `/metrics`.
      SYNC_SERVER_DISABLE_PUBLIC_METRICS: "",
      // Keep TLS disabled for local stress runs.
      SYNC_SERVER_TLS_CERT_PATH: "",
      SYNC_SERVER_TLS_KEY_PATH: "",
      ...authEnv,
      ...limitsEnv,
      ...maxMessageBytesEnv,
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  child.stdout.on("data", (d) => {
    stdout += d.toString();
    stdout = stdout.slice(-50_000);
  });
  child.stderr.on("data", (d) => {
    stderr += d.toString();
    stderr = stderr.slice(-50_000);
  });

  const proc = child;
  if (!proc) throw new Error("Failed to spawn sync-server process");

  try {
    await waitForServerReady(httpUrl, 30_000);
  } catch (err) {
    await killChild();
    throw new Error(
      `sync-server failed to start: ${String(err)}\n\n--- stdout (tail) ---\n${stdout}\n\n--- stderr (tail) ---\n${stderr}`
    );
  }

  return {
    port,
    httpUrl,
    wsUrl,
    dataDir: opts.dataDir,
    stop: killChild,
  };
}

function waitForProviderSync(
  provider: Pick<import("y-websocket").WebsocketProvider, "on" | "off" | "synced">
): Promise<void> {
  if (provider.synced === true) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      provider.off("sync", handler);
      reject(new Error("Timed out waiting for provider sync"));
    }, 30_000);
    timeout.unref();

    const handler = (isSynced: boolean) => {
      if (!isSynced) return;
      clearTimeout(timeout);
      provider.off("sync", handler);
      resolve();
    };
    provider.on("sync", handler);

    // The provider may have already synced before we attached our listener.
    if (provider.synced === true) {
      clearTimeout(timeout);
      provider.off("sync", handler);
      resolve();
    }
  });
}

function randomInt(maxExclusive: number): number {
  return Math.floor(Math.random() * maxExclusive);
}

function buildDocIds(count: number): string[] {
  const runId = `${Date.now().toString(36)}-${Math.floor(Math.random() * 1e6).toString(36)}`;
  return Array.from({ length: count }, (_, i) => `stress-${runId}-doc${i}`);
}

function buildOpaqueToken(): string {
  return `stress-${Math.random().toString(36).slice(2)}-${Date.now().toString(36)}`;
}

function buildJwtSecret(): string {
  return `stress-${Math.random().toString(36).slice(2)}-${Date.now().toString(36)}`;
}

function signJwtToken(params: {
  secret: string;
  docId: string;
  userId: string;
}): string {
  return jwt.sign(
    {
      sub: params.userId,
      docId: params.docId,
      role: "owner",
      orgId: "stress",
    },
    params.secret,
    {
      algorithm: "HS256",
      audience: "formula-sync",
      expiresIn: "10m",
    }
  );
}

async function fetchMetricsSnapshot(httpUrl: string): Promise<string[] | null> {
  try {
    const res = await fetch(new URL("/metrics", httpUrl), { signal: AbortSignal.timeout(5_000) });
    if (!res.ok) return null;
    const text = await res.text();
    const wanted = [
      "sync_server_ws_connections_total",
      "sync_server_ws_connections_current",
      "sync_server_ws_connections_rejected_total",
      "sync_server_ws_messages_rate_limited_total",
      "sync_server_ws_messages_too_large_total",
      "sync_server_ws_closes_total",
    ];
    const lines = text
      .split("\n")
      .map((l) => l.trimEnd())
      .filter((l) => l.length > 0 && !l.startsWith("#"));
    const selected = lines.filter((l) => wanted.some((m) => l.startsWith(m)));
    return selected;
  } catch {
    return null;
  }
}

async function waitForConvergence(opts: {
  clients: StressClient[];
  expectedCountsByDoc: Map<string, Map<string, number>>;
  timeoutMs: number;
}): Promise<void> {
  const start = Date.now();
  while (Date.now() - start <= opts.timeoutMs) {
    let allOk = true;
    for (const client of opts.clients) {
      const expected = opts.expectedCountsByDoc.get(client.docId);
      if (!expected) continue;
      const counters = client.doc.getMap<number>("stress:counters");

      if (counters.size < expected.size) {
        allOk = false;
        break;
      }

      for (const [key, expectedValue] of expected) {
        if (counters.get(key) !== expectedValue) {
          allOk = false;
          break;
        }
      }
      if (!allOk) break;
    }

    if (allOk) return;
    await sleep(50);
  }
  throw new Error(`Timed out waiting for convergence (${formatMs(opts.timeoutMs)})`);
}

async function main(): Promise<void> {
  // `pnpm ... stress -- --clients 10` ends up passing an extra standalone `--`
  // through to the script. Strip it so Node's parseArgs doesn't treat the rest
  // of the flags as positionals.
  const rawArgs = process.argv.slice(2).filter((a) => a !== "--");

  const { values } = parseArgs({
    args: rawArgs,
    options: {
      clients: { type: "string" },
      docs: { type: "string" },
      durationMs: { type: "string" },
      opsPerClient: { type: "string" },
      awarenessEveryMs: { type: "string" },
      maxMessageBytes: { type: "string" },
      authMode: { type: "string" },
      help: { type: "boolean", short: "h" },
    },
    allowPositionals: false,
  });

  if (values.help) {
    // eslint-disable-next-line no-console
    console.log(`
sync-server stress harness

Usage:
  pnpm -C services/sync-server stress -- [options]

Options (CLI flags override env vars):
  --clients             Number of concurrent clients (env: STRESS_CLIENTS) [default: ${DEFAULTS.clients}]
  --docs                Number of docs/rooms to spread clients across (env: STRESS_DOCS) [default: ${DEFAULTS.docs}]
  --durationMs          Workload duration in ms (env: STRESS_DURATION_MS) [default: ${DEFAULTS.durationMs}]
  --opsPerClient         Max ops per client (env: STRESS_OPS_PER_CLIENT) [default: ${DEFAULTS.opsPerClient}]
  --awarenessEveryMs    Presence update interval (0 disables) (env: STRESS_AWARENESS_EVERY_MS) [default: ${DEFAULTS.awarenessEveryMs}]
  --maxMessageBytes     Override SYNC_SERVER_MAX_MESSAGE_BYTES (env: STRESS_MAX_MESSAGE_BYTES) [default: unset]
  --authMode            Auth mode: opaque | jwt (env: STRESS_AUTH_MODE) [default: ${DEFAULTS.authMode}]

Examples:
  pnpm -C services/sync-server stress
  pnpm -C services/sync-server stress -- --clients 200 --docs 10 --durationMs 30000
  pnpm -C services/sync-server stress -- --clients 100 --opsPerClient 1000 --durationMs 60000
  pnpm -C services/sync-server stress -- --authMode jwt
`);
    return;
  }

  const opts: StressOptions = {
    clients: parsePositiveInt(
      "--clients",
      values.clients ?? parseEnvInt("STRESS_CLIENTS"),
      DEFAULTS.clients
    ),
    docs: parsePositiveInt("--docs", values.docs ?? parseEnvInt("STRESS_DOCS"), DEFAULTS.docs),
    durationMs: parsePositiveInt(
      "--durationMs",
      values.durationMs ?? parseEnvInt("STRESS_DURATION_MS"),
      DEFAULTS.durationMs
    ),
    opsPerClient: parsePositiveInt(
      "--opsPerClient",
      values.opsPerClient ?? parseEnvInt("STRESS_OPS_PER_CLIENT"),
      DEFAULTS.opsPerClient
    ),
    awarenessEveryMs: parsePositiveInt(
      "--awarenessEveryMs",
      values.awarenessEveryMs ?? parseEnvInt("STRESS_AWARENESS_EVERY_MS"),
      DEFAULTS.awarenessEveryMs
    ),
    maxMessageBytes: (() => {
      const raw = values.maxMessageBytes ?? parseEnvInt("STRESS_MAX_MESSAGE_BYTES");
      if (raw === undefined || raw === null) return DEFAULTS.maxMessageBytes;
      const parsed = Number.parseInt(String(raw), 10);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        throw new Error(`--maxMessageBytes must be a positive integer (got ${String(raw)})`);
      }
      return parsed;
    })(),
    authMode: (() => {
      const raw = (values.authMode ?? parseEnvString("STRESS_AUTH_MODE") ?? DEFAULTS.authMode)
        .toLowerCase()
        .trim();
      if (raw === "opaque" || raw === "jwt") return raw;
      throw new Error(`--authMode must be "opaque" or "jwt" (got ${raw})`);
    })(),
  };

  if (opts.clients <= 0) throw new Error("--clients must be > 0");
  if (opts.docs <= 0) throw new Error("--docs must be > 0");
  if (opts.durationMs <= 0 && opts.opsPerClient <= 0) {
    throw new Error("At least one of --durationMs or --opsPerClient must be > 0");
  }

  const docIds = buildDocIds(opts.docs);
  const opaqueToken = buildOpaqueToken();
  const jwtSecret = buildJwtSecret();

  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-stress-"));

  const maxConnectionsPerIp = Math.max(25, opts.clients + 10);
  const maxConnections = Math.max(1000, opts.clients + 50);

  const server = await startSyncServer({
    dataDir,
    authMode: opts.authMode,
    opaqueToken,
    jwtSecret,
    maxMessageBytes: opts.maxMessageBytes,
    maxConnectionsPerIp,
    maxConnections,
  });

  let exitCode = 0;
  const closeCodes = new Map<number, number>();
  const disconnectedCountByClient = new Map<number, number>();

  const abortController = new AbortController();
  const abort = (reason: string) => abortController.abort(new Error(reason));
  const onSigint = () => abort("SIGINT");
  const onSigterm = () => abort("SIGTERM");
  process.on("SIGINT", onSigint);
  process.on("SIGTERM", onSigterm);

  const clients: StressClient[] = [];
  const awarenessTimers: NodeJS.Timeout[] = [];

  const now = Date.now();

  try {
    // eslint-disable-next-line no-console
    console.log(
      [
        "sync-server stress starting",
        `  server: ${server.httpUrl} (${opts.authMode})`,
        `  clients: ${opts.clients}`,
        `  docs: ${opts.docs}`,
        `  durationMs: ${opts.durationMs}`,
        `  opsPerClient: ${opts.opsPerClient}`,
        `  awarenessEveryMs: ${opts.awarenessEveryMs}`,
        `  maxMessageBytes: ${opts.maxMessageBytes ?? "default"}`,
      ].join("\n")
    );

    // Create providers/documents.
    for (let i = 0; i < opts.clients; i += 1) {
      const docId = docIds[i % docIds.length]!;
      const userId = `u${i}`;
      const token = opts.authMode === "opaque" ? opaqueToken : signJwtToken({ secret: jwtSecret, docId, userId });

      const doc = new Y.Doc();

      const createdAt = Date.now();
      const provider = new WebsocketProvider(server.wsUrl, docId, doc, {
        WebSocketPolyfill: WebSocket,
        disableBc: true,
        params: { token },
      });

      const client: StressClient = {
        index: i,
        docId,
        userId,
        token,
        doc,
        provider,
        connectedAtMs: null,
        syncedAtMs: null,
        opsAttempted: 0,
        opsSucceeded: 0,
      };
      clients.push(client);

      let lastWs: WebSocket | null = null;
      provider.on("status", (event: { status: "connected" | "disconnected" }) => {
        if (event.status === "connected") {
          if (client.connectedAtMs === null) client.connectedAtMs = Date.now();
          const ws = (provider as unknown as { ws?: WebSocket }).ws;
          if (ws && ws !== lastWs) {
            lastWs = ws;
            ws.on("close", (code: number) => {
              closeCodes.set(code, (closeCodes.get(code) ?? 0) + 1);
            });
          }
        } else {
          disconnectedCountByClient.set(i, (disconnectedCountByClient.get(i) ?? 0) + 1);
        }
      });

      provider.on("sync", (isSynced: boolean) => {
        if (isSynced && client.syncedAtMs === null) {
          client.syncedAtMs = Date.now();
          const syncMs = client.syncedAtMs - createdAt;
          // eslint-disable-next-line no-console
          console.log(`client ${i} synced to ${docId} in ${formatMs(syncMs)}`);
        }
      });

      if (opts.awarenessEveryMs > 0) {
        const timer = setInterval(() => {
          try {
            provider.awareness.setLocalStateField("presence", {
              cursor: {
                row: randomInt(1000),
                col: randomInt(1000),
              },
              // Monotonic-ish value so updates are not deduplicated.
              t: Date.now(),
              // Intentionally spoofable field; server should sanitize.
              id: `spoof-${userId}`,
            });
          } catch {
            // ignore
          }
        }, opts.awarenessEveryMs);
        timer.unref();
        awarenessTimers.push(timer);
      }
    }

    // Wait for all clients to sync before starting the workload.
    await Promise.all(clients.map((c) => waitForProviderSync(c.provider)));

    const workloadStart = Date.now();
    const workloadEnd = opts.durationMs > 0 ? workloadStart + opts.durationMs : Number.POSITIVE_INFINITY;
    const pacingIntervalMs =
      opts.durationMs > 0 && opts.opsPerClient > 0 ? opts.durationMs / opts.opsPerClient : 0;

    const RESERVED_ROOT_FRACTION = 0.01;

    const runClientWorkload = async (client: StressClient): Promise<void> => {
      const counters = client.doc.getMap<number>("stress:counters");
      const cells = client.doc.getMap<unknown>("cells");

      const maxOps = opts.opsPerClient > 0 ? opts.opsPerClient : Number.POSITIVE_INFINITY;

      for (let opIndex = 0; opIndex < maxOps; opIndex += 1) {
        if (abortController.signal.aborted) break;
        const nowMs = Date.now();
        if (nowMs >= workloadEnd) break;

        if (pacingIntervalMs > 0) {
          const targetMs = workloadStart + opIndex * pacingIntervalMs;
          const delay = targetMs - nowMs;
          if (delay > 0) {
            await sleep(delay, undefined, { signal: abortController.signal });
          }
        } else {
          // Avoid tight loops when no pacing is configured.
          await sleep(0, undefined, { signal: abortController.signal });
        }

        client.opsAttempted += 1;

        const row = randomInt(200);
        const col = randomInt(200);
        const cellKey = `Sheet1:${row}:${col}`;

        client.doc.transact(() => {
          // Baseline "cell-ish" update.
          let cell = cells.get(cellKey);
          if (!(cell instanceof Y.Map)) {
            cell = new Y.Map();
            cells.set(cellKey, cell);
          }
          (cell as Y.Map<unknown>).set("value", `${client.userId}:${opIndex}`);
          (cell as Y.Map<unknown>).set("modifiedBy", client.userId);
          (cell as Y.Map<unknown>).set("ts", Date.now());

          // Optional reserved roots exercise.
          if (Math.random() < RESERVED_ROOT_FRACTION) {
            const branching = client.doc.getMap<unknown>("branching:stress");
            branching.set(`${client.userId}:${opIndex}`, Date.now());
          }
          if (Math.random() < RESERVED_ROOT_FRACTION) {
            const versions = client.doc.getMap<unknown>("versions");
            versions.set(`${client.userId}`, opIndex);
          }
        });

        client.opsSucceeded += 1;
      }

      // Final convergence key: each client writes its final op count to a stable,
      // per-client map key so other clients can verify convergence.
      counters.set(`c${client.index}`, client.opsSucceeded);
    };

    await Promise.all(clients.map((c) => runClientWorkload(c)));

    const workloadDurationMs = Date.now() - workloadStart;

    // Convergence: ensure all clients have observed the final per-client counts.
    const expectedCountsByDoc = new Map<string, Map<string, number>>();
    for (const docId of docIds) {
      const expected = new Map<string, number>();
      for (const client of clients) {
        if (client.docId !== docId) continue;
        expected.set(`c${client.index}`, client.opsSucceeded);
      }
      expectedCountsByDoc.set(docId, expected);
    }

    const convergenceStart = Date.now();
    await waitForConvergence({
      clients,
      expectedCountsByDoc,
      timeoutMs: Math.max(10_000, opts.durationMs * 3),
    });
    const convergenceMs = Date.now() - convergenceStart;

    const totalOpsAttempted = clients.reduce((sum, c) => sum + c.opsAttempted, 0);
    const totalOpsSucceeded = clients.reduce((sum, c) => sum + c.opsSucceeded, 0);

    const nonOkCloseCodes = [...closeCodes.entries()]
      .filter(([code]) => code !== 1000 && code !== 1001)
      .reduce((sum, [, count]) => sum + count, 0);

    const disconnectedEvents = [...disconnectedCountByClient.values()].reduce((sum, v) => sum + v, 0);

    // eslint-disable-next-line no-console
    console.log("\n--- summary ---");
    // eslint-disable-next-line no-console
    console.log(`clients: ${opts.clients}`);
    // eslint-disable-next-line no-console
    console.log(`docs: ${opts.docs}`);
    // eslint-disable-next-line no-console
    console.log(`workloadDuration: ${formatMs(workloadDurationMs)}`);
    // eslint-disable-next-line no-console
    console.log(`opsAttempted: ${totalOpsAttempted}`);
    // eslint-disable-next-line no-console
    console.log(`opsSucceeded: ${totalOpsSucceeded}`);
    // eslint-disable-next-line no-console
    console.log(
      `throughput: ${(totalOpsSucceeded / Math.max(1, workloadDurationMs / 1000)).toFixed(1)} ops/s`
    );
    // eslint-disable-next-line no-console
    console.log(`convergenceTime: ${formatMs(convergenceMs)}`);
    // eslint-disable-next-line no-console
    console.log(`disconnectEvents: ${disconnectedEvents}`);

    // eslint-disable-next-line no-console
    console.log("wsCloseCodes:");
    const sortedCodes = [...closeCodes.entries()].sort(([a], [b]) => a - b);
    for (const [code, count] of sortedCodes) {
      // eslint-disable-next-line no-console
      console.log(`  ${code}: ${count}`);
    }

    const metrics = await fetchMetricsSnapshot(server.httpUrl);
    if (metrics) {
      // eslint-disable-next-line no-console
      console.log("\nmetrics snapshot (/metrics):");
      for (const line of metrics) {
        // eslint-disable-next-line no-console
        console.log(`  ${line}`);
      }
    } else {
      // eslint-disable-next-line no-console
      console.log("\nmetrics snapshot: unavailable");
    }

    if (abortController.signal.aborted) {
      // eslint-disable-next-line no-console
      console.error("aborted");
      exitCode = 1;
    }
    if (nonOkCloseCodes > 0) {
      // eslint-disable-next-line no-console
      console.error(`unexpected websocket close codes observed: ${nonOkCloseCodes}`);
      exitCode = 1;
    }
    if (disconnectedEvents > 0) {
      // Treat any disconnect as a failure signal; tune this threshold once we
      // have more production data on expected reconnect behavior under load.
      // eslint-disable-next-line no-console
      console.error(`websocket disconnects observed: ${disconnectedEvents}`);
      exitCode = 1;
    }
  } finally {
    for (const timer of awarenessTimers) {
      try {
        clearInterval(timer);
      } catch {
        // ignore
      }
    }

    for (const client of clients) {
      try {
        client.provider.destroy();
      } catch {
        // ignore
      }
      try {
        client.doc.destroy();
      } catch {
        // ignore
      }
    }

    try {
      await server.stop();
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error(`failed to stop sync-server: ${String(err)}`);
      exitCode = 1;
    }

    process.off("SIGINT", onSigint);
    process.off("SIGTERM", onSigterm);

    try {
      await rm(dataDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  }

  if (exitCode !== 0) {
    process.exitCode = exitCode;
  } else {
    const totalMs = Date.now() - now;
    // eslint-disable-next-line no-console
    console.log(`\ncompleted in ${formatMs(totalMs)}`);
  }
}

process.on("unhandledRejection", (err) => {
  // eslint-disable-next-line no-console
  console.error("unhandledRejection", err);
  process.exitCode = 1;
});

process.on("uncaughtException", (err) => {
  // eslint-disable-next-line no-console
  console.error("uncaughtException", err);
  process.exitCode = 1;
});

await main();
