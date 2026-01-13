import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import type { SyncServerConfig } from "../src/config.js";
import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";

import { startSyncServer, waitForCondition, waitForProviderSync } from "./test-helpers.ts";

const TEST_KEYRING_JSON = JSON.stringify({
  currentVersion: 1,
  keys: {
    // 32 bytes of deterministic test key material.
    "1": Buffer.alloc(32, 7).toString("base64"),
  },
});

function yjsFilePathForDoc(dataDir: string, docName: string): string {
  const id = createHash("sha256").update(docName).digest("hex");
  return path.join(dataDir, `${id}.yjs`);
}

async function waitForFileSize(filePath: string, minBytes: number): Promise<void> {
  await waitForCondition(async () => {
    try {
      const st = await stat(filePath);
      return st.size >= minBytes;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === "ENOENT") return false;
      throw err;
    }
  }, 10_000);
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

function seedUpdate(contents: string): Uint8Array {
  const doc = new Y.Doc();
  doc.getText("t").insert(0, contents);
  return Y.encodeStateAsUpdate(doc);
}

class SlowInMemoryLeveldbPersistence {
  private readonly docs = new Map<string, Y.Doc>();
  private readonly metas = new Map<string, Map<string, unknown>>();

  constructor(private readonly delayMs: number) {}

  async getYDoc(docName: string): Promise<Y.Doc> {
    if (this.delayMs > 0) {
      await new Promise((r) => setTimeout(r, this.delayMs));
    }
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

  async clearDocument(docName: string): Promise<void> {
    this.docs.delete(docName);
    this.metas.delete(docName);
  }

  async getAllDocNames(): Promise<string[]> {
    return [...this.docs.keys()];
  }

  async getMeta(docName: string, metaKey: string): Promise<unknown> {
    return this.metas.get(docName)?.get(metaKey);
  }

  async setMeta(docName: string, metaKey: string, value: unknown): Promise<void> {
    const docMetas = this.metas.get(docName) ?? new Map<string, unknown>();
    this.metas.set(docName, docMetas);
    docMetas.set(metaKey, value);
  }

  async destroy(): Promise<void> {
    // No-op.
  }
}

function createLeveldbConfig(dataDir: string): SyncServerConfig {
  return {
    host: "127.0.0.1",
    port: 0,
    trustProxy: false,
    gc: true,
    tls: null,
    metrics: { public: true },
    dataDir,
    disableDataDirLock: true,
    persistence: {
      backend: "leveldb",
      compactAfterUpdates: 10,
      leveldbDocNameHashing: false,
      encryption: { mode: "off" },
    },
    auth: { mode: "opaque", token: "test-token" },
    enforceRangeRestrictions: false,
    introspection: null,
    internalAdminToken: null,
    retention: { ttlMs: 0, sweepIntervalMs: 0, tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000 },
    limits: {
      maxConnections: 100,
      maxConnectionsPerIp: 25,
      maxConnectionsPerDoc: 0,
      maxConnAttemptsPerWindow: 100,
      connAttemptWindowMs: 60_000,
      maxMessagesPerWindow: 2_000,
      messageWindowMs: 10_000,
      maxMessageBytes: 32 * 1024 * 1024,
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

test("loads encrypted file persistence before initial client sync after restart", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-ready-encrypted-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: TEST_KEYRING_JSON,
      // Allow large updates so persistence load takes long enough to reliably
      // race the initial sync without a readiness gate.
      SYNC_SERVER_MAX_MESSAGE_BYTES: String(32 * 1024 * 1024),
    },
  });

  const port = server.port;
  const wsUrl = server.wsUrl;
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "ready-encrypted";
  const largeText = `ready:${"x".repeat(3_000_000)}`;

  const doc1 = new Y.Doc();
  const provider1 = new WebsocketProvider(wsUrl, docName, doc1, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    provider1.destroy();
    doc1.destroy();
  });

  await waitForProviderSync(provider1);
  doc1.getText("t").insert(0, largeText);

  await waitForFileSize(yjsFilePathForDoc(dataDir, docName), 1 * 1024 * 1024);

  provider1.destroy();
  doc1.destroy();

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: TEST_KEYRING_JSON,
      SYNC_SERVER_MAX_MESSAGE_BYTES: String(32 * 1024 * 1024),
    },
  });

  const doc2 = new Y.Doc();
  const provider2 = new WebsocketProvider(wsUrl, docName, doc2, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    provider2.destroy();
    doc2.destroy();
  });

  let textAtSync: string | null = null;
  const onSync = (isSynced: boolean) => {
    if (!isSynced) return;
    if (textAtSync !== null) return;
    textAtSync = doc2.getText("t").toString();
  };
  provider2.on("sync", onSync);

  await waitForProviderSync(provider2);
  provider2.off("sync", onSync);

  // If the provider synced before we attached the handler, capture immediately.
  if (textAtSync === null) {
    textAtSync = doc2.getText("t").toString();
  }

  assert.equal(textAtSync.length, largeText.length);
  assert.equal(textAtSync.slice(0, 32), largeText.slice(0, 32));
});

test("loads LevelDB persistence before initial client sync", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-ready-leveldb-"));

  const ldb = new SlowInMemoryLeveldbPersistence(100);
  const docName = "ready-leveldb";
  await ldb.storeUpdate(docName, seedUpdate("hello"));

  const logger = createLogger("silent");
  const server = createSyncServer(createLeveldbConfig(dataDir), logger, {
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
  const provider = new WebsocketProvider(server.getWsUrl(), docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  let textAtSync: string | null = null;
  const onSync = (isSynced: boolean) => {
    if (!isSynced) return;
    if (textAtSync !== null) return;
    textAtSync = doc.getText("t").toString();
  };
  provider.on("sync", onSync);

  await waitForProviderSync(provider);
  provider.off("sync", onSync);

  if (textAtSync === null) {
    textAtSync = doc.getText("t").toString();
  }

  assert.equal(textAtSync, "hello");
});

test("rejects websocket upgrade with 503 when encrypted file persistence can't be loaded", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-ready-encrypted-fail-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: TEST_KEYRING_JSON,
    },
  });
  const port = server.port;
  const docName = "ready-encrypted-fail";

  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.wsUrl, docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);
  doc.getText("t").insert(0, "hello");

  // Ensure the encrypted header exists on disk.
  await waitForFileSize(yjsFilePathForDoc(dataDir, docName), 12);

  provider.destroy();
  doc.destroy();

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      // Now start without encryption; the persisted encrypted file should fail closed.
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
      // Ensure no external keyring env vars leak into this test.
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: "",
      SYNC_SERVER_ENCRYPTION_KEYRING_PATH: "",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: "",
    },
  });

  await expectWebSocketUpgradeStatus(
    `${server.wsUrl}/${encodeURIComponent(docName)}?token=test-token`,
    503
  );
});
