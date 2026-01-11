import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";

import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

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
  const wsUrl = `ws://127.0.0.1:${port}`;

  let server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });
  t.after(async () => {
    await server.stop();
  });

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

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
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
});

test("enforces read-only roles (viewer/commenter) for Yjs updates", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();

  const secret = "test-secret";
  const docName = "permissions-doc";

  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret, audience: "formula-sync" },
  });
  t.after(async () => {
    await server.stop();
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

test("sanitizes awareness identity and blocks clientID spoofing", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();

  const secret = "test-secret";
  const docName = "awareness-doc";

  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret, audience: "formula-sync" },
  });
  t.after(async () => {
    await server.stop();
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
