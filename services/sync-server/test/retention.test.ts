import assert from "node:assert/strict";
import test from "node:test";

import WebSocket from "ws";
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";

import type { SyncServerConfig } from "../src/config.js";
import { createLogger } from "../src/logger.js";
import { LAST_SEEN_META_KEY } from "../src/retention.js";
import { createSyncServer } from "../src/server.js";

function waitForProviderSync(provider: {
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

class InMemoryLeveldbPersistence {
  private readonly docs = new Map<string, Y.Doc>();
  private readonly metas = new Map<string, Map<string, unknown>>();

  async getYDoc(docName: string): Promise<Y.Doc> {
    return this.docs.get(docName) ?? new Y.Doc();
  }

  async storeUpdate(docName: string, update: Uint8Array): Promise<void> {
    const doc = this.docs.get(docName) ?? new Y.Doc();
    this.docs.set(docName, doc);
    Y.applyUpdate(doc, update);
  }

  async flushDocument(_docName: string): Promise<void> {
    // No-op. This implementation stores merged state in-memory.
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

function seedUpdate(contents: string): Uint8Array {
  const doc = new Y.Doc();
  doc.getText("t").insert(0, contents);
  return Y.encodeStateAsUpdate(doc);
}

function createConfig(ttlMs: number): SyncServerConfig {
  return {
    host: "127.0.0.1",
    port: 0,
    trustProxy: false,
    gc: true,
    dataDir: ":memory:",
    disableDataDirLock: true,
    persistence: {
      backend: "leveldb",
      compactAfterUpdates: 10,
      leveldbDocNameHashing: false,
      encryption: { mode: "off" },
    },
    auth: { mode: "opaque", token: "test-token" },
    internalAdminToken: "admin-token",
    retention: { ttlMs, sweepIntervalMs: 0, tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000 },
    limits: {
      maxConnections: 100,
      maxConnectionsPerIp: 25,
      maxConnAttemptsPerWindow: 100,
      connAttemptWindowMs: 60_000,
      maxMessagesPerWindow: 2_000,
      messageWindowMs: 10_000,
    },
    logLevel: "silent",
  };
}

test("retention sweep purges inactive docs (leveldb)", async (t) => {
  const ldb = new InMemoryLeveldbPersistence();
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(1_000), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const now = Date.now();

  await ldb.storeUpdate("purge-me", seedUpdate("a"));
  await ldb.setMeta("purge-me", LAST_SEEN_META_KEY, now - 10_000);

  await ldb.storeUpdate("keep-me", seedUpdate("b"));
  await ldb.setMeta("keep-me", LAST_SEEN_META_KEY, now);

  await ldb.storeUpdate("no-meta", seedUpdate("c"));

  const res = await fetch(`${server.getHttpUrl()}/internal/retention/sweep`, {
    method: "POST",
    headers: {
      "x-internal-admin-token": "admin-token",
    },
  });

  assert.equal(res.status, 200);
  const body = (await res.json()) as any;
  assert.deepEqual(body, {
    ok: true,
    scanned: 3,
    purged: 1,
    skippedActive: 0,
    skippedNoMeta: 1,
    errors: [],
  });

  const remaining = (await ldb.getAllDocNames()).sort();
  assert.deepEqual(remaining, ["keep-me", "no-meta"]);
});

test("retention sweep skips docs with active websocket connections", async (t) => {
  const ldb = new InMemoryLeveldbPersistence();
  await ldb.storeUpdate("active-doc", seedUpdate("hi"));

  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(1_000), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.getWsUrl(), "active-doc", doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);

  // Overwrite lastSeenMs after the connection is active so the sweep would purge
  // the doc if active-connection detection regressed.
  await ldb.setMeta("active-doc", LAST_SEEN_META_KEY, Date.now() - 10_000);

  const res = await fetch(`${server.getHttpUrl()}/internal/retention/sweep`, {
    method: "POST",
    headers: {
      "x-internal-admin-token": "admin-token",
    },
  });

  assert.equal(res.status, 200);
  const body = (await res.json()) as any;
  assert.equal(body.ok, true);
  assert.equal(body.scanned, 1);
  assert.equal(body.purged, 0);
  assert.equal(body.skippedActive, 1);
  assert.equal(body.skippedNoMeta, 0);

  const remaining = await ldb.getAllDocNames();
  assert.deepEqual(remaining, ["active-doc"]);
});

test("retention sweep endpoint returns 404 when internal admin token is disabled", async (t) => {
  const ldb = new InMemoryLeveldbPersistence();
  const logger = createLogger("silent");
  const server = createSyncServer(
    {
      ...createConfig(1_000),
      internalAdminToken: null,
    },
    logger,
    { createLeveldbPersistence: () => ldb as any }
  );

  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const res = await fetch(`${server.getHttpUrl()}/internal/retention/sweep`, {
    method: "POST",
    headers: {
      "x-internal-admin-token": "admin-token",
    },
  });

  assert.equal(res.status, 404);
  assert.deepEqual(await res.json(), { error: "not_found" });
});

test("retention sweep endpoint returns 400 when retention TTL is disabled", async (t) => {
  const ldb = new InMemoryLeveldbPersistence();
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(0), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const res = await fetch(`${server.getHttpUrl()}/internal/retention/sweep`, {
    method: "POST",
    headers: {
      "x-internal-admin-token": "admin-token",
    },
  });

  assert.equal(res.status, 400);
  const body = (await res.json()) as any;
  assert.equal(body.error, "retention_disabled");
  assert.equal(typeof body.message, "string");
});
