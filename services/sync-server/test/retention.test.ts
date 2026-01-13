import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import type { SyncServerConfig } from "../src/config.js";
import { sha256Hex } from "../src/leveldb-docname.js";
import { createLogger } from "../src/logger.js";
import { LAST_SEEN_META_KEY } from "../src/retention.js";
import { createSyncServer } from "../src/server.js";

import { loadYLeveldbFromTarball } from "./y-leveldb-tarball.js";

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
  getMetaCalls = 0;
  beforeClearDocument?: (docName: string) => Promise<void> | void;
  beforeGetMeta?: (docName: string, metaKey: string) => Promise<void> | void;

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
    await this.beforeClearDocument?.(docName);
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
    this.getMetaCalls += 1;
    await this.beforeGetMeta?.(docName, metaKey);
    return this.metas.get(docName)?.get(metaKey);
  }
}

function seedUpdate(contents: string): Uint8Array {
  const doc = new Y.Doc();
  doc.getText("t").insert(0, contents);
  return Y.encodeStateAsUpdate(doc);
}

function createConfig(ttlMs: number, dataDir: string): SyncServerConfig {
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
      maxQueueDepthPerDoc: 0,
      maxQueueDepthTotal: 0,
      leveldbDocNameHashing: false,
      encryption: { mode: "off" },
    },
    auth: { mode: "opaque", token: "test-token" },
    enforceRangeRestrictions: false,
    introspection: null,
    internalAdminToken: "admin-token",
    retention: { ttlMs, sweepIntervalMs: 0, tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000 },
    limits: {
      maxUrlBytes: 8192,
      maxTokenBytes: 4096,
      maxConnections: 100,
      maxConnectionsPerIp: 25,
      maxConnectionsPerDoc: 0,
      maxConnAttemptsPerWindow: 100,
      connAttemptWindowMs: 60_000,
      maxMessagesPerWindow: 2_000,
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

function expectWebSocketUpgradeStatus(
  url: string,
  expectedStatusCode: number
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
      try {
        assert.equal(res.statusCode, expectedStatusCode);
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

test("retention sweep purges inactive docs (leveldb)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-retention-"));

  const ldb = new InMemoryLeveldbPersistence();
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(1_000, dataDir), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-retention-"));

  const ldb = new InMemoryLeveldbPersistence();
  await ldb.storeUpdate("active-doc", seedUpdate("hi"));

  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(1_000, dataDir), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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

test(
  "rejects websocket upgrades while retention is purging a legacy (raw docName) document (leveldb hashing)",
  async (t) => {
    const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-retention-"));

    const docName = "legacy-doc";

    const ldb = new InMemoryLeveldbPersistence();
    await ldb.storeUpdate(docName, seedUpdate("hi"));
    await ldb.setMeta(docName, LAST_SEEN_META_KEY, Date.now() - 10_000);

    let clearStartedResolve: (() => void) | null = null;
    const clearStarted = new Promise<void>((resolve) => {
      clearStartedResolve = resolve;
    });

    let clearContinueResolve: (() => void) | null = null;
    const clearContinue = new Promise<void>((resolve) => {
      clearContinueResolve = resolve;
    });

    ldb.beforeClearDocument = async (name: string) => {
      if (name !== docName) return;
      clearStartedResolve?.();
      await clearContinue;
    };

    const logger = createLogger("silent");
    const baseConfig = createConfig(1_000, dataDir);
    const server = createSyncServer(
      {
        ...baseConfig,
        persistence: {
          ...baseConfig.persistence,
          leveldbDocNameHashing: true,
        },
      },
      logger,
      { createLeveldbPersistence: () => ldb as any }
    );

    await server.start();
    t.after(async () => {
      await server.stop();
    });
    t.after(async () => {
      await rm(dataDir, { recursive: true, force: true });
    });

    const sweepPromise = fetch(`${server.getHttpUrl()}/internal/retention/sweep`, {
      method: "POST",
      headers: {
        "x-internal-admin-token": "admin-token",
      },
    });

    await clearStarted;

    await expectWebSocketUpgradeStatus(
      `${server.getWsUrl()}/${docName}?token=test-token`,
      503
    );

    clearContinueResolve?.();
    const res = await sweepPromise;
    assert.equal(res.status, 200);
  }
);

test("pong-based lastSeen refresh uses persisted docName when docName hashing is enabled", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-retention-"));

  const ldb = new InMemoryLeveldbPersistence();
  const logger = createLogger("silent");
  const baseConfig = createConfig(60_000, dataDir);
  const server = createSyncServer(
    {
      ...baseConfig,
      persistence: {
        ...baseConfig.persistence,
        leveldbDocNameHashing: true,
      },
    },
    logger,
    { createLeveldbPersistence: () => ldb as any }
  );

  await server.start();
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "active-doc";
  const persistedName = sha256Hex(docName);

  const ws = new WebSocket(`${server.getWsUrl()}/${docName}?token=test-token`);
  t.after(() => {
    ws.terminate();
  });

  await waitForWebSocketOpen(ws);

  // Wait for the initial `markSeen(persistedName)` to land.
  await new Promise((r) => setTimeout(r, 20));

  // Send a pong frame to trigger the server's `ws.on("pong")` handler.
  ws.pong();
  await new Promise((r) => setTimeout(r, 20));

  assert.equal(typeof (await ldb.getMeta(persistedName, LAST_SEEN_META_KEY)), "number");
  assert.equal(await ldb.getMeta(docName, LAST_SEEN_META_KEY), undefined);
});

test("retention sweep endpoint returns 404 when internal admin token is disabled", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-retention-"));

  const ldb = new InMemoryLeveldbPersistence();
  const logger = createLogger("silent");
  const baseConfig = createConfig(1_000, dataDir);
  const server = createSyncServer(
    {
      ...baseConfig,
      internalAdminToken: null,
    },
    logger,
    { createLeveldbPersistence: () => ldb as any }
  );

  await server.start();
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-retention-"));

  const ldb = new InMemoryLeveldbPersistence();
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(0, dataDir), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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

test("retention sweep purges docs using y-leveldb + level-mem (no native LevelDB)", async (t) => {
  const { LeveldbPersistence } = await loadYLeveldbFromTarball(t);

  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-leveldb-"));

  let ldb: any;
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(1_000, dataDir), logger, {
    createLeveldbPersistence: (location: string) => {
      ldb = new LeveldbPersistence(location);
      return ldb;
    },
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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

test("retention sweep endpoint de-duplicates concurrent requests", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-retention-"));

  const ldb = new InMemoryLeveldbPersistence();
  await ldb.storeUpdate("keep-me", seedUpdate("hello"));
  await ldb.setMeta("keep-me", LAST_SEEN_META_KEY, Date.now());

  // Slow down the sweep so we can overlap two requests.
  ldb.beforeGetMeta = async (_docName, metaKey) => {
    if (metaKey !== LAST_SEEN_META_KEY) return;
    await new Promise((r) => setTimeout(r, 200));
  };

  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(60_000, dataDir), logger, {
    createLeveldbPersistence: () => ldb as any,
  });

  await server.start();
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const url = `${server.getHttpUrl()}/internal/retention/sweep`;
  const requestInit: RequestInit = {
    method: "POST",
    headers: {
      "x-internal-admin-token": "admin-token",
    },
  };

  const [resA, resB] = await Promise.all([fetch(url, requestInit), fetch(url, requestInit)]);
  assert.equal(resA.status, 200);
  assert.equal(resB.status, 200);

  const [bodyA, bodyB] = await Promise.all([(await resA.json()) as any, (await resB.json()) as any]);
  assert.equal(bodyA.ok, true);
  assert.equal(bodyB.ok, true);

  // Two concurrent requests should share the same in-flight sweep rather than
  // scanning LevelDB twice.
  assert.equal(ldb.getMetaCalls, 1);
});
