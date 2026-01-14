import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import type { SyncServerConfig } from "../src/config.js";
import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";
import { waitForCondition } from "./test-helpers.ts";

function createConfig(
  dataDir: string,
  overrides: Partial<SyncServerConfig> = {}
): SyncServerConfig {
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

function waitForWebSocketEnd(ws: WebSocket): Promise<void> {
  return new Promise((resolve) => {
    if (ws.readyState === WebSocket.CLOSED) return resolve();
    ws.once("close", () => resolve());
    ws.once("error", () => resolve());
  });
}

function getMetricValue(body: string, name: string, labels?: Record<string, string>): number | null {
  const labelFragment = labels
    ? `\\{[^}]*${Object.entries(labels)
        .map(([k, v]) => `${k}="${v}"`)
        .join("[^}]*")}[^}]*\\}`
    : "(?:\\{[^}]*\\})?";
  const match = body.match(
    // Allow default labels (e.g. `{service="sync-server"}`) while keeping the matcher simple.
    new RegExp(`^${name}${labelFragment}\\s+(-?\\d+(?:\\.\\d+)?(?:e[+-]?\\d+)?)$`, "m")
  );
  if (!match) return null;
  const value = Number(match[1]);
  return Number.isFinite(value) ? value : null;
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

test("stop() enters drain mode, rejects new upgrades, and terminates remaining clients after grace", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-shutdown-drain-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const graceMs = 500;

  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(dataDir, { shutdownGraceMs: graceMs }), logger);
  await server.start();
  const httpUrl = server.getHttpUrl();
  const wsUrl = server.getWsUrl();

  t.after(async () => {
    try {
      await server.stop();
    } catch {
      // ignore
    }
  });

  const ws = new WebSocket(`${wsUrl}/drain-doc?token=test-token`);
  t.after(() => {
    try {
      ws.terminate();
    } catch {
      // ignore
    }
  });
  await waitForWebSocketOpen(ws);
  const wsEnded = waitForWebSocketEnd(ws);

  const stopStartedAt = Date.now();
  const stopPromise = server.stop();

  // The server should not immediately terminate existing connections when grace > 0.
  await new Promise((r) => setTimeout(r, 50));
  assert.equal(ws.readyState, WebSocket.OPEN);

  // During drain mode, readiness should fail so load balancers stop routing.
  await waitForCondition(async () => {
    const res = await fetch(`${httpUrl}/readyz`);
    if (res.status !== 503) return false;
    const body = (await res.json()) as { reason?: unknown };
    return body.reason === "draining";
  }, 2_000);

  // New websocket upgrades should be rejected while draining.
  await expectWebSocketUpgradeStatus(
    `${wsUrl}/new-doc?token=test-token`,
    503,
    "draining"
  );

  // Metrics should expose drain state and count rejections with a bounded reason label.
  const metricsRes = await fetch(`${httpUrl}/metrics`);
  assert.equal(metricsRes.status, 200);
  const metricsText = await metricsRes.text();
  assert.equal(getMetricValue(metricsText, "sync_server_shutdown_draining_current"), 1);
  const rejectedDraining = getMetricValue(metricsText, "sync_server_ws_connections_rejected_total", {
    reason: "draining",
  });
  assert.ok(typeof rejectedDraining === "number" && rejectedDraining >= 1);

  // The stop should complete after the grace period (since we keep the socket open).
  const timeout = new Promise<never>((_, reject) => {
    const t = setTimeout(() => reject(new Error("Timed out waiting for server.stop()")), 5_000);
    t.unref();
  });
  await Promise.race([stopPromise, timeout]);

  const elapsed = Date.now() - stopStartedAt;
  assert.ok(elapsed >= graceMs - 50);

  await wsEnded;
});
