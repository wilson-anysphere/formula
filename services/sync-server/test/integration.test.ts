import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import type { AddressInfo } from "node:net";
import net from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import jwt from "jsonwebtoken";
import WebSocket from "ws";
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";

async function waitForCondition(
  condition: () => boolean | Promise<boolean>,
  timeoutMs: number,
  intervalMs: number = 25
): Promise<void> {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (await condition()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

async function getAvailablePort(): Promise<number> {
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

async function waitForServerReady(baseUrl: string): Promise<void> {
  await waitForCondition(async () => {
    try {
      const res = await fetch(`${baseUrl}/healthz`);
      return res.ok;
    } catch {
      return false;
    }
  }, 10_000);
}

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

test("syncs between two clients and persists across restart", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const httpUrl = `http://127.0.0.1:${port}`;
  const wsUrl = `ws://127.0.0.1:${port}`;

  const serviceDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    ".."
  );
  const entry = path.join(serviceDir, "src", "index.ts");

  let serverProcess: ReturnType<typeof spawn> | null = null;
  let stdout = "";
  let stderr = "";

  const startServer = async () => {
    const child = spawn(process.execPath, ["--import", "tsx", entry], {
      cwd: serviceDir,
      env: {
        ...process.env,
        NODE_ENV: "test",
        LOG_LEVEL: "silent",
        SYNC_SERVER_HOST: "127.0.0.1",
        SYNC_SERVER_PORT: String(port),
        SYNC_SERVER_DATA_DIR: dataDir,
        SYNC_SERVER_AUTH_TOKEN: "test-token",
        SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "10"
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

    serverProcess = child;

    try {
      await waitForServerReady(httpUrl);
    } catch (err) {
      child.kill("SIGTERM");
      throw new Error(
        `Server failed to start: ${String(err)}\nstdout:\n${stdout}\nstderr:\n${stderr}`
      );
    }
  };

  const stopServer = async () => {
    const child = serverProcess;
    if (!child) return;
    serverProcess = null;
    child.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      child.once("exit", () => resolve());
    });
  };

  t.after(async () => {
    await stopServer();
  });

  await startServer();

  const docName = "test-doc";

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

  await stopServer();
  await startServer();

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
});

test("enforces read-only roles (viewer/commenter) for Yjs updates", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const httpUrl = `http://127.0.0.1:${port}`;
  const wsUrl = `ws://127.0.0.1:${port}`;

  const secret = "test-secret";
  const docName = "permissions-doc";

  const serviceDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    ".."
  );
  const entry = path.join(serviceDir, "src", "index.ts");

  const serverProcess = spawn(process.execPath, ["--import", "tsx", entry], {
    cwd: serviceDir,
    env: {
      ...process.env,
      NODE_ENV: "test",
      LOG_LEVEL: "silent",
      SYNC_SERVER_HOST: "127.0.0.1",
      SYNC_SERVER_PORT: String(port),
      SYNC_SERVER_DATA_DIR: dataDir,
      SYNC_SERVER_JWT_SECRET: secret,
      SYNC_SERVER_JWT_AUDIENCE: "formula-sync",
      SYNC_SERVER_PERSISTENCE_BACKEND: "file",
    },
    stdio: ["ignore", "ignore", "ignore"],
  });
  const stopServer = async () => {
    if (serverProcess.exitCode !== null) return;
    serverProcess.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      serverProcess.once("exit", () => resolve());
    });
  };
  t.after(async () => {
    await stopServer();
  });

  await waitForServerReady(httpUrl);

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

test("sanitizes awareness identity and blocks clientID spoofing", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const httpUrl = `http://127.0.0.1:${port}`;
  const wsUrl = `ws://127.0.0.1:${port}`;

  const secret = "test-secret";
  const docName = "awareness-doc";

  const serviceDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    ".."
  );
  const entry = path.join(serviceDir, "src", "index.ts");

  const serverProcess = spawn(process.execPath, ["--import", "tsx", entry], {
    cwd: serviceDir,
    env: {
      ...process.env,
      NODE_ENV: "test",
      LOG_LEVEL: "silent",
      SYNC_SERVER_HOST: "127.0.0.1",
      SYNC_SERVER_PORT: String(port),
      SYNC_SERVER_DATA_DIR: dataDir,
      SYNC_SERVER_JWT_SECRET: secret,
      SYNC_SERVER_JWT_AUDIENCE: "formula-sync",
      SYNC_SERVER_PERSISTENCE_BACKEND: "file",
    },
    stdio: ["ignore", "ignore", "ignore"],
  });
  const stopServer = async () => {
    if (serverProcess.exitCode !== null) return;
    serverProcess.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      serverProcess.once("exit", () => resolve());
    });
  };
  t.after(async () => {
    await stopServer();
  });

  await waitForServerReady(httpUrl);

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
});
