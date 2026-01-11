import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

const JWT_SECRET = "test-secret";
const JWT_AUDIENCE = "formula-sync";

function signJwt(payload: Record<string, unknown>): string {
  return jwt.sign(payload, JWT_SECRET, {
    algorithm: "HS256",
    audience: JWT_AUDIENCE,
  });
}

async function expectConditionToStayFalse(
  condition: () => boolean,
  timeoutMs: number
): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    assert.equal(condition(), false);
    await new Promise((r) => setTimeout(r, 25));
  }
}

function writeVarUint(dst: number[], value: number) {
  let num = value >>> 0;
  while (num >= 0x80) {
    dst.push((num & 0x7f) | 0x80);
    num >>>= 7;
  }
  dst.push(num);
}

function writeVarUint8Array(dst: number[], data: Uint8Array) {
  writeVarUint(dst, data.length);
  for (const byte of data) dst.push(byte);
}

function writeVarString(dst: number[], value: string) {
  const bytes = new TextEncoder().encode(value);
  writeVarUint8Array(dst, bytes);
}

function encodeAwarenessUpdate(entries: {
  clientId: number;
  clock: number;
  state: unknown;
}[]): Uint8Array {
  const out: number[] = [];
  writeVarUint(out, entries.length);
  for (const entry of entries) {
    writeVarUint(out, entry.clientId);
    writeVarUint(out, entry.clock);
    writeVarString(out, JSON.stringify(entry.state));
  }
  return Uint8Array.from(out);
}

function encodeWebsocketAwarenessMessage(awarenessUpdate: Uint8Array): Uint8Array {
  const messageAwareness = 1;
  const out: number[] = [];
  writeVarUint(out, messageAwareness);
  writeVarUint8Array(out, awarenessUpdate);
  return Uint8Array.from(out);
}

test("rejects JWT when docId mismatches path", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
  });
  t.after(async () => {
    await server.stop();
  });

  const token = signJwt({ sub: "u1", docId: "doc-a", orgId: "o1", role: "editor" });
  const url = `${server.wsUrl}/doc-b?token=${encodeURIComponent(token)}`;

  const statusCode = await new Promise<number>((resolve, reject) => {
    let settled = false;
    const ws = new WebSocket(url);

    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      fn();
    };

    ws.once("open", () => {
      finish(() => {
        ws.close();
        reject(new Error("Expected connection to be rejected"));
      });
    });

    ws.once("unexpected-response", (_req, res) => {
      finish(() => {
        res.resume();
        ws.terminate();
        resolve(res.statusCode ?? 0);
      });
    });

    ws.once("error", (err) => {
      finish(() => {
        const match = String(err).match(/\b(\d{3})\b/);
        if (match) resolve(Number(match[1]));
        else reject(err);
      });
    });
  });

  assert.equal(statusCode, 403);
});

test("viewer can sync but cannot write", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const editorToken = signJwt({
    sub: "editor",
    docId: docName,
    orgId: "o1",
    role: "editor",
  });
  const viewerToken = signJwt({
    sub: "viewer",
    docId: docName,
    orgId: "o1",
    role: "viewer",
  });

  const docEditor = new Y.Doc();
  const docViewer = new Y.Doc();

  const providerEditor = new WebsocketProvider(server.wsUrl, docName, docEditor, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  const providerViewer = new WebsocketProvider(server.wsUrl, docName, docViewer, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: viewerToken },
  });

  t.after(() => {
    providerEditor.destroy();
    providerViewer.destroy();
    docEditor.destroy();
    docViewer.destroy();
  });

  await waitForProviderSync(providerEditor);
  await waitForProviderSync(providerViewer);

  docEditor.getText("t").insert(0, "hello");

  await waitForCondition(() => docViewer.getText("t").toString() === "hello", 10_000);
  assert.equal(docViewer.getText("t").toString(), "hello");

  docViewer.getText("t").insert(docViewer.getText("t").length, "evil");
  await expectConditionToStayFalse(
    () => docEditor.getText("t").toString().includes("evil"),
    1_000
  );

  providerViewer.destroy();
  docViewer.destroy();

  const docEditor2 = new Y.Doc();
  const providerEditor2 = new WebsocketProvider(server.wsUrl, docName, docEditor2, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });

  t.after(() => {
    providerEditor2.destroy();
    docEditor2.destroy();
  });

  await waitForProviderSync(providerEditor2);
  await waitForCondition(() => docEditor2.getText("t").toString() === "hello", 10_000);
  assert.equal(docEditor2.getText("t").toString(), "hello");
});

test("awareness sanitizes presence.id and blocks clientID spoof", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;

  const attackerToken = signJwt({
    sub: "attacker",
    docId: docName,
    orgId: "o1",
    role: "editor",
  });
  const victimToken = signJwt({
    sub: "victim",
    docId: docName,
    orgId: "o1",
    role: "editor",
  });

  const docAttacker = new Y.Doc();
  const docVictim = new Y.Doc();

  const attackerProvider = new WebsocketProvider(server.wsUrl, docName, docAttacker, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: attackerToken },
  });
  const victimProvider = new WebsocketProvider(server.wsUrl, docName, docVictim, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: victimToken },
  });

  t.after(() => {
    attackerProvider.destroy();
    victimProvider.destroy();
    docAttacker.destroy();
    docVictim.destroy();
  });

  await waitForProviderSync(attackerProvider);
  await waitForProviderSync(victimProvider);

  victimProvider.awareness.setLocalStateField("presence", {
    v: 1,
    id: "victim",
    name: "V",
    color: "#0f0",
    sheet: "Sheet1",
    cursor: null,
    selections: [],
    lastActive: Date.now(),
  });

  attackerProvider.awareness.setLocalStateField("presence", {
    v: 1,
    id: "victim",
    name: "Evil",
    color: "#f00",
    sheet: "Sheet1",
    cursor: null,
    selections: [],
    lastActive: Date.now(),
  });

  await waitForCondition(() => {
    const state = victimProvider.awareness.getStates().get(docAttacker.clientID) as any;
    return state?.presence?.id === "attacker";
  }, 10_000);

  const receivedAttackerState = victimProvider.awareness
    .getStates()
    .get(docAttacker.clientID) as any;
  assert.equal(receivedAttackerState?.presence?.id, "attacker");

  await waitForCondition(
    () => attackerProvider.awareness.getStates().has(docVictim.clientID),
    10_000
  );

  const victimClientIdFromAttacker = [...attackerProvider.awareness.getStates().keys()].find(
    (id) => id !== docAttacker.clientID
  );
  assert.equal(victimClientIdFromAttacker, docVictim.clientID);

  const spoofRemove = encodeWebsocketAwarenessMessage(
    encodeAwarenessUpdate([
      {
        clientId: docVictim.clientID,
        clock: 9999,
        state: null,
      },
    ])
  );

  assert.ok(attackerProvider.ws, "Expected attacker WebsocketProvider to have a ws");
  attackerProvider.ws.send(Buffer.from(spoofRemove));

  await new Promise((r) => setTimeout(r, 250));

  assert.ok(victimProvider.awareness.getStates().has(docVictim.clientID));
  assert.ok(attackerProvider.awareness.getStates().has(docVictim.clientID));
});
