import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";
import type { SyncServerConfig } from "../src/config.js";

function createConfig(dataDir: string): SyncServerConfig {
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

function expectWebSocketUpgradeStatus(url: string, expectedStatusCode: number): Promise<void> {
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
      try {
        assert.equal(res.statusCode, expectedStatusCode);
        res.resume();
        finish(resolve);
      } catch (err) {
        finish(() => reject(err));
      }
    });

    ws.on("error", (err) => {
      if (finished) return;
      reject(err);
    });
  });
}

test("accepts document ids up to 1024 bytes and rejects larger ids", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-docname-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(dataDir), logger);
  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const okDocName = "a".repeat(1024);
  const okWs = new WebSocket(`${server.getWsUrl()}/${okDocName}?token=test-token`);
  t.after(() => okWs.terminate());
  await waitForWebSocketOpen(okWs);
  okWs.close();

  const tooLongDocName = "a".repeat(1025);
  await expectWebSocketUpgradeStatus(
    `${server.getWsUrl()}/${tooLongDocName}?token=test-token`,
    414
  );
});

test("accepts document ids containing :// sequences (no absolute-form confusion)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-docname-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(dataDir), logger);
  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const docName = "doc://name";
  const ws = new WebSocket(`${server.getWsUrl()}/${docName}?token=test-token`);
  t.after(() => ws.terminate());
  await waitForWebSocketOpen(ws);
  ws.close();
});
