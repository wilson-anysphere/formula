import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";
import type { SyncServerConfig } from "../src/config.js";

function createConfig(dataDir: string, overrides: Partial<SyncServerConfig> = {}): SyncServerConfig {
  return {
    host: "127.0.0.1",
    port: 0,
    trustProxy: false,
    gc: true,
    shutdownGraceMs: 0,
    tls: null,
    metrics: { public: true },
    dataDir,
    disableDataDirLock: true,
    persistence: {
      backend: "file",
      compactAfterUpdates: 10,
      maxQueueDepthPerDoc: 0,
      maxQueueDepthTotal: 0,
      leveldbDocNameHashing: false,
      encryption: { mode: "off" },
    },
    auth: { mode: "opaque", token: "test-token" },
    enforceRangeRestrictions: false,
    introspection: null,
    internalAdminToken: null,
    retention: { ttlMs: 0, sweepIntervalMs: 0, tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000 },
    limits: {
      maxUrlBytes: 8192,
      maxTokenBytes: 4096,
      maxConnections: 100,
      maxConnectionsPerIp: 25,
      maxConnectionsPerDoc: 0,
      maxConnAttemptsPerWindow: 100,
      connAttemptWindowMs: 60_000,
      maxMessageBytes: 2 * 1024 * 1024,
      maxMessagesPerWindow: 2_000,
      messageWindowMs: 10_000,
      maxMessagesPerIpWindow: 0,
      ipMessageWindowMs: 0,
      maxAwarenessStateBytes: 64 * 1024,
      maxAwarenessEntries: 10,
      maxMessagesPerDocWindow: 10_000,
      docMessageWindowMs: 10_000,
      maxBranchingCommitsPerDoc: 0,
      maxVersionsPerDoc: 0,
    },
    logLevel: "silent",
    ...overrides,
  };
}

function waitForWebSocketOpen(ws: WebSocket): Promise<void> {
  return new Promise((resolve, reject) => {
    const onError = (err: unknown) => {
      ws.off("open", onOpen);
      reject(err);
    };
    const onOpen = () => {
      ws.off("error", onError);
      resolve();
    };
    ws.once("error", onError);
    ws.once("open", onOpen);
  });
}

function expectWebSocketUpgradeStatus(
  url: string,
  expectedStatusCode: number,
  expectedBodySubstring?: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(url);
    let finished = false;

    const finish = (cb: () => void) => {
      if (finished) return;
      finished = true;
      try {
        ws.terminate();
      } catch {
        // ignore
      }
      cb();
    };

    ws.on("open", () => {
      finish(() => reject(new Error("Expected WebSocket upgrade rejection")));
    });

    ws.on("unexpected-response", (_req, res) => {
      res.setEncoding("utf8");
      let body = "";
      res.on("data", (chunk) => {
        body += chunk;
      });
      res.on("end", () => {
        try {
          assert.equal(res.statusCode, expectedStatusCode);
          if (expectedBodySubstring) {
            assert.ok(body.includes(expectedBodySubstring));
          }
          finish(resolve);
        } catch (err) {
          finish(() => reject(err));
        }
      });
    });

    ws.on("error", (err) => {
      if (finished) return;
      reject(err);
    });
  });
}

test("rejects websocket upgrade when connection attempts exceed the per-IP rate limit", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-attempts-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const logger = createLogger("silent");
  const server = createSyncServer(
    createConfig(dataDir, {
      limits: {
        ...createConfig(dataDir).limits,
        maxConnAttemptsPerWindow: 1,
        connAttemptWindowMs: 60_000,
      },
    }),
    logger
  );
  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const wsUrl = `${server.getWsUrl()}/attempts-doc?token=test-token`;

  const ws1 = new WebSocket(wsUrl);
  t.after(() => ws1.terminate());
  await waitForWebSocketOpen(ws1);

  await expectWebSocketUpgradeStatus(wsUrl, 429, "Too Many Requests");
});

test("rejects websocket upgrade when max connections per IP is exceeded", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-max-per-ip-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const base = createConfig(dataDir);
  const logger = createLogger("silent");
  const server = createSyncServer(
    {
      ...base,
      limits: {
        ...base.limits,
        maxConnectionsPerIp: 1,
        // Ensure the attempt limiter does not interfere with this test.
        maxConnAttemptsPerWindow: 100,
        connAttemptWindowMs: 60_000,
      },
    },
    logger
  );
  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const wsUrl = `${server.getWsUrl()}/max-per-ip-doc?token=test-token`;

  const ws1 = new WebSocket(wsUrl);
  t.after(() => ws1.terminate());
  await waitForWebSocketOpen(ws1);

  await expectWebSocketUpgradeStatus(wsUrl, 429, "max_connections_per_ip_exceeded");
});

test("rejects websocket upgrade when max total connections is exceeded", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-max-total-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const base = createConfig(dataDir);
  const logger = createLogger("silent");
  const server = createSyncServer(
    {
      ...base,
      trustProxy: true,
      limits: {
        ...base.limits,
        maxConnections: 1,
        // Ensure per-IP limit does not interfere with total limit.
        maxConnectionsPerIp: 100,
      },
    },
    logger
  );
  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const url = `${server.getWsUrl()}/max-total-doc?token=test-token`;

  const ws1 = new WebSocket(url, { headers: { "x-forwarded-for": "1.2.3.4" } });
  t.after(() => ws1.terminate());
  await waitForWebSocketOpen(ws1);

  // Use a distinct x-forwarded-for so the server treats it as a separate IP.
  await expectWebSocketUpgradeStatus(url, 429, "max_connections_exceeded");
});

test("trustProxy falls back to remoteAddress when x-forwarded-for is empty", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-trust-proxy-empty-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const base = createConfig(dataDir);
  const logger = createLogger("silent");
  const server = createSyncServer(
    {
      ...base,
      trustProxy: true,
      limits: {
        ...base.limits,
        maxConnectionsPerIp: 1,
        // Disable attempt limiting so we only test per-IP connection tracking.
        maxConnAttemptsPerWindow: 0,
        connAttemptWindowMs: 0,
      },
    },
    logger
  );
  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const url = `${server.getWsUrl()}/trust-proxy-empty?token=test-token`;

  // A blank x-forwarded-for should be treated as invalid and fall back to the
  // remote address (127.0.0.1 in this test environment).
  const ws1 = new WebSocket(url, { headers: { "x-forwarded-for": " " } });
  t.after(() => ws1.terminate());
  await waitForWebSocketOpen(ws1);

  // If the first connection used an empty-string IP key, this second connection
  // would be allowed (different key). With the fallback it should be rejected.
  await expectWebSocketUpgradeStatus(url, 429, "max_connections_per_ip_exceeded");
});
