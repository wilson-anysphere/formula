import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import type { SyncServerConfig } from "../src/config.js";
import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";
import { waitForProviderSync } from "./test-helpers.ts";

class DelayedLeveldbPersistence {
  private readonly docs = new Map<string, Y.Doc>();
  private readonly metas = new Map<string, Map<string, unknown>>();

  storeUpdateCalls = 0;

  private resolveFirstStoreUpdate: (() => void) | null = null;
  firstStoreUpdateStarted = new Promise<void>((resolve) => {
    this.resolveFirstStoreUpdate = resolve;
  });

  private readonly pendingResolves: Array<() => void> = [];

  releasePendingWrites(): void {
    const pending = this.pendingResolves.splice(0);
    for (const resolve of pending) resolve();
  }

  async getYDoc(docName: string): Promise<Y.Doc> {
    return this.docs.get(docName) ?? new Y.Doc();
  }

  async storeUpdate(docName: string, update: Uint8Array): Promise<void> {
    this.storeUpdateCalls += 1;
    if (this.storeUpdateCalls === 1) {
      this.resolveFirstStoreUpdate?.();
    }

    const doc = this.docs.get(docName) ?? new Y.Doc();
    this.docs.set(docName, doc);
    Y.applyUpdate(doc, update);

    await new Promise<void>((resolve) => {
      this.pendingResolves.push(resolve);
    });
  }

  async flushDocument(_docName: string): Promise<void> {
    // No-op for in-memory persistence.
  }

  async destroy(): Promise<void> {
    // No-op.
  }

  async clearDocument(docName: string): Promise<void> {
    this.docs.delete(docName);
    this.metas.delete(docName);
  }

  async getAllDocNames(): Promise<string[]> {
    return [...this.docs.keys()];
  }

  async setMeta(docName: string, metaKey: string, value: unknown): Promise<void> {
    const docMetas = this.metas.get(docName) ?? new Map<string, unknown>();
    this.metas.set(docName, docMetas);
    docMetas.set(metaKey, value);
  }

  async getMeta(docName: string, metaKey: string): Promise<unknown> {
    return this.metas.get(docName)?.get(metaKey);
  }
}

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
      backend: "leveldb",
      compactAfterUpdates: 10,
      maxQueueDepthPerDoc: 1,
      maxQueueDepthTotal: 0,
      leveldbDocNameHashing: false,
      encryption: { mode: "off" },
    },
    auth: { mode: "opaque", token: "test-token" },
    enforceRangeRestrictions: false,
    introspection: null,
    internalAdminToken: null,
    retention: {
      ttlMs: 0,
      sweepIntervalMs: 0,
      tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000,
    },
    limits: {
      maxUrlBytes: 8192,
      maxTokenBytes: 4096,
      maxConnections: 100,
      maxConnectionsPerIp: 25,
      maxConnectionsPerDoc: 0,
      maxConnAttemptsPerWindow: 100,
      connAttemptWindowMs: 60_000,
      maxMessagesPerWindow: 10_000,
      messageWindowMs: 10_000,
      maxMessagesPerIpWindow: 0,
      ipMessageWindowMs: 0,
      maxMessageBytes: 2 * 1024 * 1024,
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

function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error("Timed out")), timeoutMs);
    timeout.unref();
    promise.then(
      (value) => {
        clearTimeout(timeout);
        resolve(value);
      },
      (err) => {
        clearTimeout(timeout);
        reject(err);
      }
    );
  });
}

test("LevelDB persistence backpressure closes active sockets when the per-doc queue is full", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-leveldb-backpressure-"));

  const ldb = new DelayedLeveldbPersistence();
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(dataDir), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    // Ensure we can drain the queue so `server.stop()` doesn't hang.
    ldb.releasePendingWrites();
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "bp-doc";
  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.getWsUrl(), docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);

  // Keep a raw socket open to capture the close code deterministically (the
  // provider may auto-reconnect).
  const rawWs = new WebSocket(`${server.getWsUrl()}/${docName}?token=test-token`);
  rawWs.on("message", () => {
    // drain
  });
  await waitForWebSocketOpen(rawWs);
  t.after(() => {
    rawWs.terminate();
  });

  const close = new Promise<{ code: number; reason: Buffer }>((resolve) => {
    rawWs.once("close", (code, reason) => resolve({ code, reason: Buffer.from(reason) }));
  });

  // First update enqueues a write that never resolves (simulated LevelDB stall).
  doc.getText("t").insert(0, "a");
  await ldb.firstStoreUpdateStarted;
  assert.equal(ldb.storeUpdateCalls, 1);

  // Second update should trip maxQueueDepthPerDoc=1 and force-close the doc sockets.
  doc.getText("t").insert(1, "b");

  const { code, reason } = await withTimeout(close, 5_000);
  assert.equal(code, 1013);
  assert.match(reason.toString("utf8"), /persistence overloaded/);

  // Stop the provider from attempting to reconnect after the forced close.
  provider.destroy();

  // Give the server a tick; it should not have enqueued a second storeUpdate.
  await new Promise((r) => setImmediate(r));
  assert.equal(ldb.storeUpdateCalls, 1);

  // Allow cleanup.
  ldb.releasePendingWrites();
});
