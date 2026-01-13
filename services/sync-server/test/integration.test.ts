import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { hostname, tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import jwt from "jsonwebtoken";
import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import {
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

async function waitForProcessExit(
  child: ReturnType<typeof spawn>,
  timeoutMs: number
): Promise<{ code: number | null; signal: NodeJS.Signals | null }> {
  return await new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      try {
        child.kill("SIGKILL");
      } catch {
        // ignore
      }
      reject(new Error("Timed out waiting for process exit"));
    }, timeoutMs);
    timeout.unref();

    const onExit = (code: number | null, signal: NodeJS.Signals | null) => {
      clearTimeout(timeout);
      resolve({ code, signal });
    };

    child.once("exit", onExit);
    if (child.exitCode !== null || child.signalCode !== null) {
      child.off("exit", onExit);
      onExit(child.exitCode, child.signalCode);
    }
  });
}

const TEST_KEYRING_JSON = JSON.stringify({
  currentVersion: 1,
  keys: {
    // 32 bytes of deterministic test key material.
    "1": Buffer.alloc(32, 7).toString("base64"),
  },
});

function signJwtToken(params: {
  secret: string;
  docId: string;
  userId: string;
  role: "owner" | "admin" | "editor" | "commenter" | "viewer";
}): string {
  return jwt.sign(
    {
      sub: params.userId,
      docId: params.docId,
      role: params.role,
    },
    params.secret,
    {
      algorithm: "HS256",
      audience: "formula-sync",
      expiresIn: "10m",
    }
  );
}

function yjsFilePathForDoc(dataDir: string, docName: string): string {
  const id = createHash("sha256").update(docName).digest("hex");
  return path.join(dataDir, `${id}.yjs`);
}

async function waitForFile(filePath: string): Promise<void> {
  await waitForCondition(async () => {
    try {
      const bytes = await readFile(filePath);
      return bytes.byteLength > 0;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === "ENOENT") return false;
      throw err;
    }
  }, 10_000);
}

async function expectWsUpgradeStatus(
  url: string,
  expectedStatusCode: number
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
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

function parseEncryptedYjsKeyVersions(bytes: Buffer): number[] {
  if (bytes.byteLength < 12) return [];
  if (bytes.subarray(0, 8).toString("ascii") !== "FMLYJS01") return [];
  const flags = bytes.readUInt8(8);
  if ((flags & 0b1) !== 0b1) return [];

  const out: number[] = [];
  let offset = 12; // header bytes: magic(8) + flags(1) + reserved(3)
  while (offset + 4 <= bytes.byteLength) {
    const recordLen = bytes.readUInt32BE(offset);
    offset += 4;
    if (offset + recordLen > bytes.byteLength) break;
    if (recordLen < 4) break;
    out.push(bytes.readUInt32BE(offset));
    offset += recordLen;
  }
  return out;
}

function encodeVarUint(value: number): Uint8Array {
  const bytes: number[] = [];
  let v = value;
  while (v > 0x7f) {
    bytes.push(0x80 | (v % 0x80));
    v = Math.floor(v / 0x80);
  }
  bytes.push(v);
  return Uint8Array.from(bytes);
}

function concatUint8Arrays(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((sum, arr) => sum + arr.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const arr of arrays) {
    out.set(arr, offset);
    offset += arr.length;
  }
  return out;
}

function encodeVarString(value: string): Uint8Array {
  const encoded = new TextEncoder().encode(value);
  return concatUint8Arrays([encodeVarUint(encoded.length), encoded]);
}

function buildAwarenessMessage(entries: {
  clientID: number;
  clock: number;
  stateJSON: string;
}[]): Buffer {
  const updateChunks: Uint8Array[] = [encodeVarUint(entries.length)];
  for (const entry of entries) {
    updateChunks.push(encodeVarUint(entry.clientID));
    updateChunks.push(encodeVarUint(entry.clock));
    updateChunks.push(encodeVarString(entry.stateJSON));
  }
  const update = concatUint8Arrays(updateChunks);
  const fullMessage = concatUint8Arrays([
    encodeVarUint(1),
    encodeVarUint(update.length),
    update,
  ]);
  return Buffer.from(fullMessage);
}

test("syncs between two clients and persists encrypted state across restart", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: TEST_KEYRING_JSON,
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

  const docName = "test-doc";
  const secretText = "encryption-at-rest-secret: hello world 0123456789 abcdef";

  const doc1 = new Y.Doc();
  const doc2 = new Y.Doc();

  const provider1 = new WebsocketProvider(wsUrl, docName, doc1, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const provider2 = new WebsocketProvider(wsUrl, docName, doc2, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider1.destroy();
    provider2.destroy();
    doc1.destroy();
    doc2.destroy();
  });

  await waitForProviderSync(provider1);
  await waitForProviderSync(provider2);

  doc1.getText("t").insert(0, secretText);

  await waitForCondition(
    () => doc2.getText("t").toString() === secretText,
    10_000
  );
  assert.equal(doc2.getText("t").toString(), secretText);

  provider1.destroy();
  provider2.destroy();
  doc1.destroy();
  doc2.destroy();

  const persistedPath = yjsFilePathForDoc(dataDir, docName);
  await waitForFile(persistedPath);
  const persistedBytes = await readFile(persistedPath);
  assert.equal(persistedBytes.subarray(0, 8).toString("ascii"), "FMLYJS01");
  assert.equal(persistedBytes.readUInt8(8) & 0b1, 0b1);
  assert.equal(
    persistedBytes.includes(Buffer.from(secretText, "utf8")),
    false,
    "encrypted persistence should not contain plaintext UTF-8"
  );

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: TEST_KEYRING_JSON,
    },
  });

  const doc3 = new Y.Doc();
  const provider3 = new WebsocketProvider(wsUrl, docName, doc3, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider3.destroy();
    doc3.destroy();
  });

  await waitForProviderSync(provider3);
  await waitForCondition(
    () => doc3.getText("t").toString() === secretText,
    10_000
  );
  assert.equal(doc3.getText("t").toString(), secretText);
});

test(
  "syncs and persists encrypted state across restart with SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64",
  async (t) => {
    const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

    const keyBase64 = Buffer.alloc(32, 5).toString("base64");

    let server = await startSyncServer({
      dataDir,
      auth: { mode: "opaque", token: "test-token" },
      env: {
        // Ensure key_b64 implies keyring mode even when SYNC_SERVER_PERSISTENCE_ENCRYPTION is off.
        SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
        SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: keyBase64,
        // Ensure no external keyring env vars leak into this test.
        SYNC_SERVER_ENCRYPTION_KEYRING_JSON: "",
        SYNC_SERVER_ENCRYPTION_KEYRING_PATH: "",
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

    const docName = "test-doc-key-b64";
    const secretText =
      "encryption-at-rest-secret(key_b64): hello world 0123456789 abcdef";

    const doc1 = new Y.Doc();
    const doc2 = new Y.Doc();

    const provider1 = new WebsocketProvider(wsUrl, docName, doc1, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: "test-token" },
    });
    const provider2 = new WebsocketProvider(wsUrl, docName, doc2, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: "test-token" },
    });

    t.after(() => {
      provider1.destroy();
      provider2.destroy();
      doc1.destroy();
      doc2.destroy();
    });

    await waitForProviderSync(provider1);
    await waitForProviderSync(provider2);

    doc1.getText("t").insert(0, secretText);

    await waitForCondition(
      () => doc2.getText("t").toString() === secretText,
      10_000
    );

    provider1.destroy();
    provider2.destroy();
    doc1.destroy();
    doc2.destroy();

    const persistedPath = yjsFilePathForDoc(dataDir, docName);
    await waitForFile(persistedPath);
    const persistedBytes = await readFile(persistedPath);
    assert.equal(persistedBytes.subarray(0, 8).toString("ascii"), "FMLYJS01");
    assert.equal(persistedBytes.readUInt8(8) & 0b1, 0b1);
    assert.equal(
      persistedBytes.includes(Buffer.from(secretText, "utf8")),
      false,
      "encrypted persistence should not contain plaintext UTF-8"
    );

    await server.stop();
    server = await startSyncServer({
      port,
      dataDir,
      auth: { mode: "opaque", token: "test-token" },
      env: {
        SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
        SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: keyBase64,
        SYNC_SERVER_ENCRYPTION_KEYRING_JSON: "",
        SYNC_SERVER_ENCRYPTION_KEYRING_PATH: "",
      },
    });

    const doc3 = new Y.Doc();
    const provider3 = new WebsocketProvider(wsUrl, docName, doc3, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: "test-token" },
    });

    t.after(() => {
      provider3.destroy();
      doc3.destroy();
    });

    await waitForProviderSync(provider3);
    await waitForCondition(
      () => doc3.getText("t").toString() === secretText,
      10_000
    );
    assert.equal(doc3.getText("t").toString(), secretText);
  }
);

test("purges persisted documents via internal admin API", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      // Ensure any externally-provided internal token doesn't leak into tests.
      INTERNAL_ADMIN_TOKEN: "",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "test-admin",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
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

  const docName = "purge-doc";

  const doc1 = new Y.Doc();
  const doc2 = new Y.Doc();

  const provider1 = new WebsocketProvider(wsUrl, docName, doc1, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const provider2 = new WebsocketProvider(wsUrl, docName, doc2, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider1.destroy();
    provider2.destroy();
    doc1.destroy();
    doc2.destroy();
  });

  await waitForProviderSync(provider1);
  await waitForProviderSync(provider2);

  doc1.getText("t").insert(0, "hello");

  await waitForCondition(() => doc2.getText("t").toString() === "hello", 10_000);
  assert.equal(doc2.getText("t").toString(), "hello");

  provider1.destroy();
  provider2.destroy();
  doc1.destroy();
  doc2.destroy();

  // Give the server a moment to persist state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 250));

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      INTERNAL_ADMIN_TOKEN: "",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "test-admin",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });

  const doc3 = new Y.Doc();
  const provider3 = new WebsocketProvider(wsUrl, docName, doc3, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider3.destroy();
    doc3.destroy();
  });

  await waitForProviderSync(provider3);
  await waitForCondition(() => doc3.getText("t").toString() === "hello", 10_000);
  assert.equal(doc3.getText("t").toString(), "hello");

  const purgePath = `${server.httpUrl}/internal/docs/${encodeURIComponent(docName)}`;

  const forbidden = await fetch(purgePath, {
    method: "DELETE",
    headers: {
      "x-internal-admin-token": "wrong-token",
    },
  });
  assert.equal(forbidden.status, 403);
  assert.deepEqual(await forbidden.json(), { error: "forbidden" });

  // Keep the doc active on the server while purging, but avoid reconnection logic
  // from the Yjs provider by holding a raw websocket open.
  const rawWs = new WebSocket(`${wsUrl}/${docName}?token=test-token`);
  await new Promise<void>((resolve, reject) => {
    rawWs.once("open", () => resolve());
    rawWs.once("error", (err) => reject(err));
  });
  provider3.destroy();

  const purged = await fetch(purgePath, {
    method: "DELETE",
    headers: {
      "x-internal-admin-token": "test-admin",
    },
  });
  assert.equal(purged.status, 200);
  assert.deepEqual(await purged.json(), { ok: true });

  await new Promise<void>((resolve) => {
    if (rawWs.readyState === WebSocket.CLOSED) return resolve();
    rawWs.once("close", () => resolve());
  });

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      INTERNAL_ADMIN_TOKEN: "",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "test-admin",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });

  await expectWsUpgradeStatus(
    `${wsUrl}/${encodeURIComponent(docName)}?token=test-token`,
    410
  );

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      INTERNAL_ADMIN_TOKEN: "",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });

  const disabled = await fetch(
    `${server.httpUrl}/internal/docs/${encodeURIComponent(docName)}`,
    {
      method: "DELETE",
      headers: {
        "x-internal-admin-token": "test-admin",
      },
    }
  );
  assert.equal(disabled.status, 404);
  assert.deepEqual(await disabled.json(), { error: "not_found" });
});

test("migrates legacy plaintext persistence files to encrypted format", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: { SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off" },
  });
  const port = server.port;
  const wsUrl = server.wsUrl;
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "migration-doc";
  const secretText = "migration-secret: abcdefghijklmnopqrstuvwxyz 0123456789";

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
  doc1.getText("t").insert(0, secretText);

  provider1.destroy();
  doc1.destroy();
  await server.stop();

  const persistedPath = yjsFilePathForDoc(dataDir, docName);
  await waitForFile(persistedPath);
  const plaintextBytes = await readFile(persistedPath);
  assert.notEqual(
    plaintextBytes.subarray(0, 8).toString("ascii"),
    "FMLYJS01",
    "expected legacy plaintext file without header"
  );

  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: TEST_KEYRING_JSON,
    },
  });

  const migratedBytes = await readFile(persistedPath);
  assert.equal(migratedBytes.subarray(0, 8).toString("ascii"), "FMLYJS01");
  assert.equal(migratedBytes.readUInt8(8) & 0b1, 0b1);
  assert.equal(
    migratedBytes.includes(Buffer.from(secretText, "utf8")),
    false,
    "migrated encrypted persistence should not contain plaintext UTF-8"
  );

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

  await waitForProviderSync(provider2);
  await waitForCondition(
    () => doc2.getText("t").toString() === secretText,
    10_000
  );
  assert.equal(doc2.getText("t").toString(), secretText);
});

test("supports KeyRing rotation for encrypted file persistence", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const docName = "rotation-doc";
  const secretText = "rotation-secret: hello world 0123456789 abcdef";

  const keyV1 = Buffer.alloc(32, 7).toString("base64");
  const keyV2 = Buffer.alloc(32, 8).toString("base64");

  const keyRingV1 = JSON.stringify({
    currentVersion: 1,
    keys: { "1": keyV1 },
  });
  const keyRingV2 = JSON.stringify({
    currentVersion: 2,
    keys: { "1": keyV1, "2": keyV2 },
  });

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: keyRingV1,
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

  {
    const doc = new Y.Doc();
    const provider = new WebsocketProvider(wsUrl, docName, doc, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: "test-token" },
    });

    t.after(() => {
      provider.destroy();
      doc.destroy();
    });

    await waitForProviderSync(provider);
    doc.getText("t").insert(0, secretText);

    provider.destroy();
    doc.destroy();

    const persistedPath = yjsFilePathForDoc(dataDir, docName);
    await waitForCondition(async () => {
      try {
        const bytes = await readFile(persistedPath);
        return parseEncryptedYjsKeyVersions(bytes).includes(1);
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code === "ENOENT") return false;
        throw err;
      }
    }, 10_000);

    const bytesV1 = await readFile(persistedPath);
    assert.equal(bytesV1.subarray(0, 8).toString("ascii"), "FMLYJS01");
    assert.equal(bytesV1.readUInt8(8) & 0b1, 0b1);
    assert.equal(bytesV1.includes(Buffer.from(secretText, "utf8")), false);

    const versionsV1 = parseEncryptedYjsKeyVersions(bytesV1);
    assert.ok(
      versionsV1.includes(1),
      `expected persisted records to include keyVersion=1 (got ${versionsV1.join(",")})`
    );
  }

  await server.stop();

  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: keyRingV2,
    },
  });

  {
    const doc = new Y.Doc();
    const provider = new WebsocketProvider(wsUrl, docName, doc, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: "test-token" },
    });

    t.after(() => {
      provider.destroy();
      doc.destroy();
    });

    await waitForProviderSync(provider);
    await waitForCondition(
      () => doc.getText("t").toString() === secretText,
      10_000
    );
    assert.equal(doc.getText("t").toString(), secretText);

    // Trigger a write under the rotated keyring (currentVersion=2).
    doc.getText("t").insert(secretText.length, "!");

    provider.destroy();
    doc.destroy();

    const persistedPath = yjsFilePathForDoc(dataDir, docName);
    await waitForCondition(async () => {
      try {
        const bytes = await readFile(persistedPath);
        return parseEncryptedYjsKeyVersions(bytes).includes(2);
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code === "ENOENT") return false;
        throw err;
      }
    }, 10_000);
    const bytesV2 = await readFile(persistedPath);
    assert.equal(bytesV2.subarray(0, 8).toString("ascii"), "FMLYJS01");
    assert.equal(bytesV2.readUInt8(8) & 0b1, 0b1);
    assert.equal(bytesV2.includes(Buffer.from(secretText, "utf8")), false);

    const versionsV2 = parseEncryptedYjsKeyVersions(bytesV2);
    assert.ok(
      versionsV2.includes(2),
      `expected persisted records to include keyVersion=2 (got ${versionsV2.join(",")})`
    );
  }
});

test("loads KeyRing from SYNC_SERVER_ENCRYPTION_KEYRING_PATH", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const keyRingPath = path.join(dataDir, "keyring.json");
  await writeFile(keyRingPath, TEST_KEYRING_JSON);

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      SYNC_SERVER_ENCRYPTION_KEYRING_JSON: undefined,
      SYNC_SERVER_ENCRYPTION_KEYRING_PATH: keyRingPath,
    },
  });
  const wsUrl = server.wsUrl;
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "keyring-path-doc";
  const secretText = "keyring-path-secret: hello world 0123456789 abcdef";

  const doc = new Y.Doc();
  const provider = new WebsocketProvider(wsUrl, docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);
  doc.getText("t").insert(0, secretText);

  provider.destroy();
  doc.destroy();

  const persistedPath = yjsFilePathForDoc(dataDir, docName);
  await waitForFile(persistedPath);
  const persistedBytes = await readFile(persistedPath);
  assert.equal(persistedBytes.subarray(0, 8).toString("ascii"), "FMLYJS01");
  assert.equal(persistedBytes.readUInt8(8) & 0b1, 0b1);
  assert.equal(persistedBytes.includes(Buffer.from(secretText, "utf8")), false);
});

test(
  "refuses to start in production when encryption is enabled but keyring is missing",
  async (t) => {
    const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-keyring-"));
    t.after(async () => {
      await rm(dataDir, { recursive: true, force: true });
    });

    const serviceDir = path.resolve(
      path.dirname(fileURLToPath(import.meta.url)),
      ".."
    );
    const entry = path.join(serviceDir, "src", "index.ts");
    const nodeWithTsx = path.join(serviceDir, "scripts", "node-with-tsx.mjs");

    let stdout = "";
    let stderr = "";
    const child = spawn(process.execPath, [nodeWithTsx, entry], {
      cwd: serviceDir,
      env: {
        ...process.env,
        NODE_ENV: "production",
        LOG_LEVEL: "silent",
        SYNC_SERVER_HOST: "127.0.0.1",
        SYNC_SERVER_PORT: "0",
        SYNC_SERVER_DATA_DIR: dataDir,
        SYNC_SERVER_AUTH_TOKEN: "test-token",
        SYNC_SERVER_JWT_SECRET: "",
        SYNC_SERVER_JWT_AUDIENCE: "",
        SYNC_SERVER_JWT_ISSUER: "",
        SYNC_SERVER_PERSISTENCE_BACKEND: "file",
        SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "10",
        SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
        SYNC_SERVER_ENCRYPTION_KEYRING_JSON: undefined,
        SYNC_SERVER_ENCRYPTION_KEYRING_PATH: undefined,
        // Ensure unrelated encryption flags don't accidentally fail config parsing.
        SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: undefined,
      },
      stdio: ["ignore", "pipe", "pipe"],
    });

    child.stdout?.on("data", (d) => {
      stdout += d.toString();
      stdout = stdout.slice(-10_000);
    });
    child.stderr?.on("data", (d) => {
      stderr += d.toString();
      stderr = stderr.slice(-10_000);
    });

    const exit = await waitForProcessExit(child, 10_000);
    assert.notEqual(exit.code, 0);

    const combinedLogs = `${stdout}\n${stderr}`.toLowerCase();
    assert.ok(
      combinedLogs.includes("sync_server_persistence_encryption=keyring") &&
        combinedLogs.includes("sync_server_encryption_keyring_json"),
      `expected missing keyring error in logs.\nstdout:\n${stdout}\nstderr:\n${stderr}`
    );
  }
);

test("enforces viewer role as read-only for Yjs updates", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const secret = "test-secret";
  const docName = "permissions-doc";

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret, audience: "formula-sync" },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const wsUrl = server.wsUrl;

  const editorToken = signJwtToken({
    secret,
    docId: docName,
    userId: "editor-user",
    role: "editor",
  });
  const viewerToken = signJwtToken({
    secret,
    docId: docName,
    userId: "viewer-user",
    role: "viewer",
  });

  const editorDoc = new Y.Doc();
  const viewerDoc = new Y.Doc();

  const editorProvider = new WebsocketProvider(wsUrl, docName, editorDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  const viewerProvider = new WebsocketProvider(wsUrl, docName, viewerDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: viewerToken },
  });

  t.after(() => {
    editorProvider.destroy();
    viewerProvider.destroy();
    editorDoc.destroy();
    viewerDoc.destroy();
  });

  await waitForProviderSync(editorProvider);
  await waitForProviderSync(viewerProvider);

  editorDoc.getText("t").insert(0, "hello");
  await waitForCondition(
    () => viewerDoc.getText("t").toString() === "hello",
    10_000
  );

  // Viewer tries to write; server must drop the update.
  viewerDoc.getText("t").insert(5, "evil");

  // Give the server a moment to (not) broadcast the viewer update.
  await new Promise((r) => setTimeout(r, 250));
  assert.equal(editorDoc.getText("t").toString(), "hello");

  // A fresh editor connection should observe the server state unchanged.
  const observerDoc = new Y.Doc();
  const observerProvider = new WebsocketProvider(wsUrl, docName, observerDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  t.after(() => {
    observerProvider.destroy();
    observerDoc.destroy();
  });

  await waitForProviderSync(observerProvider);
  assert.equal(observerDoc.getText("t").toString(), "hello");
});

test("allows commenter role to write to the comments root but rejects other Yjs updates", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const secret = "test-secret";
  const docName = "commenter-comments-doc";

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret, audience: "formula-sync" },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const wsUrl = server.wsUrl;

  const editorToken = signJwtToken({
    secret,
    docId: docName,
    userId: "editor-user",
    role: "editor",
  });
  const commenterToken = signJwtToken({
    secret,
    docId: docName,
    userId: "commenter-user",
    role: "commenter",
  });

  const editorDoc = new Y.Doc();
  const commenterDoc = new Y.Doc();

  const editorProvider = new WebsocketProvider(wsUrl, docName, editorDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  const commenterProvider = new WebsocketProvider(wsUrl, docName, commenterDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: commenterToken },
  });

  t.after(() => {
    editorProvider.destroy();
    commenterProvider.destroy();
    editorDoc.destroy();
    commenterDoc.destroy();
  });

  await waitForProviderSync(editorProvider);
  await waitForProviderSync(commenterProvider);

  // Commenter writes to the `comments` root; this should sync.
  commenterDoc.getMap("comments").set("c1", "hello");

  await waitForCondition(
    () => editorDoc.getMap("comments").get("c1") === "hello",
    10_000
  );

  // Commenter attempts to write to any other root; server must reject.
  commenterDoc.getText("t").insert(0, "evil");

  // Give the server a moment to (not) broadcast the update.
  await new Promise((r) => setTimeout(r, 250));

  // The editor should not observe a new root or content.
  assert.equal(editorDoc.share.has("t"), false);
});

test("sanitizes awareness identity and blocks clientID spoofing", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const secret = "test-secret";
  const docName = "awareness-doc";

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret, audience: "formula-sync" },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const wsUrl = server.wsUrl;

  const tokenA = signJwtToken({
    secret,
    docId: docName,
    userId: "user-a",
    role: "editor",
  });
  const tokenB = signJwtToken({
    secret,
    docId: docName,
    userId: "user-b",
    role: "editor",
  });

  const docA = new Y.Doc();
  const docB = new Y.Doc();

  const providerA = new WebsocketProvider(wsUrl, docName, docA, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: tokenA },
  });
  const providerB = new WebsocketProvider(wsUrl, docName, docB, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: tokenB },
  });

  t.after(() => {
    providerA.destroy();
    providerB.destroy();
    docA.destroy();
    docB.destroy();
  });

  await waitForProviderSync(providerA);
  await waitForProviderSync(providerB);

  const clientIdA = docA.clientID;

  // Spoof the identity fields; the server must rewrite them to match the JWT sub.
  (providerA as any).awareness.setLocalState({
    presence: { id: "spoofed", name: "Alice" },
    userId: "spoofed",
    user: { id: "spoofed" },
    id: "spoofed",
  });

  await waitForCondition(() => {
    const states = (providerB as any).awareness.getStates() as Map<
      number,
      any
    >;
    const state = states.get(clientIdA);
    return Boolean(state && state.presence && state.presence.id === "user-a");
  }, 10_000);

  {
    const states = (providerB as any).awareness.getStates() as Map<number, any>;
    const state = states.get(clientIdA);
    assert.equal(state.presence.id, "user-a");
    assert.equal(state.userId, "user-a");
    assert.equal(state.user.id, "user-a");
    assert.equal(state.id, "user-a");
  }

  // A malicious raw socket that has claimed its own clientID must not be able to
  // remove another client's awareness state.
  const attackerToken = signJwtToken({
    secret,
    docId: docName,
    userId: "attacker",
    role: "editor",
  });
  const attackerClientId = 1_234_567_890;
  const attackerWs = new WebSocket(
    `${wsUrl}/${docName}?token=${encodeURIComponent(attackerToken)}`
  );

  t.after(() => {
    attackerWs.terminate();
  });

  await new Promise<void>((resolve, reject) => {
    attackerWs.once("open", () => resolve());
    attackerWs.once("error", reject);
  });

  attackerWs.send(
    buildAwarenessMessage([
      {
        clientID: attackerClientId,
        clock: 1,
        stateJSON: JSON.stringify({ presence: { id: "spoofed" } }),
      },
    ])
  );

  // Attempt to remove user A by spoofing their clientID; should be ignored.
  attackerWs.send(
    buildAwarenessMessage([
      { clientID: clientIdA, clock: 999, stateJSON: "null" },
    ])
  );

  await new Promise((r) => setTimeout(r, 250));
  {
    const states = (providerB as any).awareness.getStates() as Map<number, any>;
    const state = states.get(clientIdA);
    assert.equal(state?.presence?.id, "user-a");
  }

  // Attempt to claim A's clientID from a fresh connection; server must close the
  // websocket with 1008 (policy violation).
  const collidingToken = signJwtToken({
    secret,
    docId: docName,
    userId: "collider",
    role: "editor",
  });
  const collidingWs = new WebSocket(
    `${wsUrl}/${docName}?token=${encodeURIComponent(collidingToken)}`
  );
  t.after(() => {
    collidingWs.terminate();
  });

  await new Promise<void>((resolve, reject) => {
    collidingWs.once("open", () => resolve());
    collidingWs.once("error", reject);
  });

  const close = new Promise<{ code: number; reason: Buffer }>((resolve) => {
    collidingWs.once("close", (code, reason) => resolve({ code, reason }));
  });

  collidingWs.send(
    buildAwarenessMessage([
      {
        clientID: clientIdA,
        clock: 1,
        stateJSON: JSON.stringify({ presence: { id: "spoofed" } }),
      },
    ])
  );

  const { code } = await close;
  assert.equal(code, 1008);

  // Defensive: ensure the security filter also applies to *text* websocket frames.
  // ws emits those with `isBinary=false`, but y-websocket will still parse the raw bytes.
  const textTokenA = signJwtToken({
    secret,
    docId: docName,
    userId: "text-a",
    role: "editor",
  });
  const textTokenB = signJwtToken({
    secret,
    docId: docName,
    userId: "text-b",
    role: "editor",
  });

  const textClientIdA = 42;
  const textClientIdB = 43;

  const textWsA = new WebSocket(
    `${wsUrl}/${docName}?token=${encodeURIComponent(textTokenA)}`
  );
  const textWsB = new WebSocket(
    `${wsUrl}/${docName}?token=${encodeURIComponent(textTokenB)}`
  );

  t.after(() => {
    textWsA.terminate();
    textWsB.terminate();
  });

  await Promise.all([
    new Promise<void>((resolve, reject) => {
      textWsA.once("open", () => resolve());
      textWsA.once("error", reject);
    }),
    new Promise<void>((resolve, reject) => {
      textWsB.once("open", () => resolve());
      textWsB.once("error", reject);
    }),
  ]);

  // Claim clientIDs with text frames and spoof the identity fields.
  textWsA.send(
    buildAwarenessMessage([
      {
        clientID: textClientIdA,
        clock: 1,
        stateJSON: JSON.stringify({ presence: { id: "spoofed" } }),
      },
    ]),
    { binary: false }
  );
  textWsB.send(
    buildAwarenessMessage([
      {
        clientID: textClientIdB,
        clock: 1,
        stateJSON: JSON.stringify({ presence: { id: "spoofed" } }),
      },
    ]),
    { binary: false }
  );

  await waitForCondition(() => {
    const states = (providerB as any).awareness.getStates() as Map<number, any>;
    const aState = states.get(textClientIdA);
    const bState = states.get(textClientIdB);
    return (
      aState?.presence?.id === "text-a" && bState?.presence?.id === "text-b"
    );
  }, 10_000);

  // Attempt to remove A using B (text frame spoof); must be ignored.
  textWsB.send(
    buildAwarenessMessage([
      { clientID: textClientIdA, clock: 2, stateJSON: "null" },
    ]),
    { binary: false }
  );

  await new Promise((r) => setTimeout(r, 250));
  {
    const states = (providerB as any).awareness.getStates() as Map<number, any>;
    const aState = states.get(textClientIdA);
    assert.equal(aState?.presence?.id, "text-a");
  }

  // Collision check via text frames (policy violation close 1008).
  const textColliderToken = signJwtToken({
    secret,
    docId: docName,
    userId: "text-collider",
    role: "editor",
  });
  const textColliderWs = new WebSocket(
    `${wsUrl}/${docName}?token=${encodeURIComponent(textColliderToken)}`
  );
  t.after(() => {
    textColliderWs.terminate();
  });

  await new Promise<void>((resolve, reject) => {
    textColliderWs.once("open", () => resolve());
    textColliderWs.once("error", reject);
  });

  const textCollisionClose = new Promise<number>((resolve) => {
    textColliderWs.once("close", (code) => resolve(code));
  });

  textColliderWs.send(
    buildAwarenessMessage([
      {
        clientID: textClientIdA,
        clock: 1,
        stateJSON: JSON.stringify({ presence: { id: "spoofed" } }),
      },
    ]),
    { binary: false }
  );

  assert.equal(await textCollisionClose, 1008);
});

test("refuses to start a second server using the same data directory", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-lock-"));

  const serviceDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    ".."
  );
  const entry = path.join(serviceDir, "src", "index.ts");
  const nodeWithTsx = path.join(serviceDir, "scripts", "node-with-tsx.mjs");

  const server1 = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });
  t.after(async () => {
    await server1.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  let stdout2 = "";
  let stderr2 = "";
  const server2 = spawn(process.execPath, [nodeWithTsx, entry], {
    cwd: serviceDir,
    env: {
      ...process.env,
      NODE_ENV: "test",
      LOG_LEVEL: "silent",
      SYNC_SERVER_HOST: "127.0.0.1",
      SYNC_SERVER_PORT: "0",
      SYNC_SERVER_DATA_DIR: dataDir,
      SYNC_SERVER_AUTH_TOKEN: "test-token",
      SYNC_SERVER_JWT_SECRET: "",
      SYNC_SERVER_JWT_AUDIENCE: "",
      SYNC_SERVER_JWT_ISSUER: "",
      SYNC_SERVER_PERSISTENCE_BACKEND: "file",
      SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "10",
      SYNC_SERVER_PERSISTENCE_MAX_QUEUE_DEPTH_PER_DOC: "",
      SYNC_SERVER_PERSISTENCE_MAX_QUEUE_DEPTH_TOTAL: "",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  server2.stdout?.on("data", (d) => {
    stdout2 += d.toString();
    stdout2 = stdout2.slice(-10_000);
  });
  server2.stderr?.on("data", (d) => {
    stderr2 += d.toString();
    stderr2 = stderr2.slice(-10_000);
  });

  const exit2 = await waitForProcessExit(server2, 10_000);
  assert.notEqual(exit2.code, 0);

  const combinedLogs = `${stdout2}\n${stderr2}`.toLowerCase();
  assert.ok(
    combinedLogs.includes(".sync-server.lock") ||
      combinedLogs.includes("data directory lock") ||
      combinedLogs.includes("acquire") ||
      combinedLogs.includes("lock"),
    `expected lock acquisition error in logs.\nstdout:\n${stdout2}\nstderr:\n${stderr2}`
  );
});

test("does not treat lock from another host as stale", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-lock-host-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const lockPath = path.join(dataDir, ".sync-server.lock");
  await writeFile(
    lockPath,
    `${JSON.stringify({
      pid: process.pid + 10_000_000,
      startedAtMs: Date.now(),
      host: "definitely-not-this-host",
    })}\n`
  );

  const serviceDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    ".."
  );
  const entry = path.join(serviceDir, "src", "index.ts");
  const nodeWithTsx = path.join(serviceDir, "scripts", "node-with-tsx.mjs");

  let stdout = "";
  let stderr = "";
  const child = spawn(process.execPath, [nodeWithTsx, entry], {
    cwd: serviceDir,
    env: {
      ...process.env,
      NODE_ENV: "test",
      LOG_LEVEL: "silent",
      SYNC_SERVER_HOST: "127.0.0.1",
      SYNC_SERVER_PORT: "0",
      SYNC_SERVER_DATA_DIR: dataDir,
      SYNC_SERVER_AUTH_TOKEN: "test-token",
      SYNC_SERVER_JWT_SECRET: "",
      SYNC_SERVER_JWT_AUDIENCE: "",
      SYNC_SERVER_JWT_ISSUER: "",
      SYNC_SERVER_PERSISTENCE_BACKEND: "file",
      SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "10",
      SYNC_SERVER_PERSISTENCE_MAX_QUEUE_DEPTH_PER_DOC: "",
      SYNC_SERVER_PERSISTENCE_MAX_QUEUE_DEPTH_TOTAL: "",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  child.stdout?.on("data", (d) => {
    stdout += d.toString();
    stdout = stdout.slice(-10_000);
  });
  child.stderr?.on("data", (d) => {
    stderr += d.toString();
    stderr = stderr.slice(-10_000);
  });

  const exit = await waitForProcessExit(child, 10_000);
  assert.notEqual(exit.code, 0);

  const combinedLogs = `${stdout}\n${stderr}`.toLowerCase();
  assert.ok(
    combinedLogs.includes(".sync-server.lock") ||
      combinedLogs.includes("data directory lock") ||
      combinedLogs.includes("lock was created on a different host") ||
      combinedLogs.includes("lock"),
    `expected lock acquisition error in logs.\nstdout:\n${stdout}\nstderr:\n${stderr}`
  );
});

test("removes stale lock from the same host and continues startup", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-lock-stale-"));

  const stalePid = process.pid + 10_000_000;
  const lockPath = path.join(dataDir, ".sync-server.lock");
  await writeFile(
    lockPath,
    `${JSON.stringify({
      pid: stalePid,
      startedAtMs: Date.now(),
      host: hostname(),
    })}\n`
  );

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const activeLock = JSON.parse(await readFile(lockPath, "utf8")) as {
    pid?: unknown;
  };
  assert.equal(typeof activeLock.pid, "number");
  assert.notEqual(activeLock.pid, stalePid);
});

test("/readyz returns 503 when data dir locking is disabled", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-readyz-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_DISABLE_DATA_DIR_LOCK: "true",
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const res = await fetch(`${server.httpUrl}/readyz`);
  assert.equal(res.status, 503);

  const body = (await res.json()) as { reason?: unknown };
  assert.equal(body.reason, "data_dir_lock_disabled");
});

test("removes data directory lock file on graceful shutdown", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-lock-cleanup-"));
  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const lockPath = path.join(dataDir, ".sync-server.lock");
  const lockContents = await readFile(lockPath, "utf8");
  assert.ok(lockContents.includes('"pid"'), "expected lock file to contain metadata");

  await server.stop();

  await waitForCondition(async () => {
    try {
      await readFile(lockPath, "utf8");
      return false;
    } catch (err) {
      return (err as NodeJS.ErrnoException).code === "ENOENT";
    }
  }, 5_000);
});
