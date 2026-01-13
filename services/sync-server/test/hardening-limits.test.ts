import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import {
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

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

test("closes websocket on oversized binary message", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-limits-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_MESSAGE_BYTES: "64",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const ws = new WebSocket(`${server.wsUrl}/limits-doc?token=test-token`);
  t.after(() => ws.terminate());

  await new Promise<void>((resolve, reject) => {
    ws.once("open", () => resolve());
    ws.once("error", (err) => reject(err));
  });

  const close = new Promise<number>((resolve) => {
    ws.once("close", (code) => resolve(code));
  });

  ws.send(Buffer.alloc(65, 1));

  assert.equal(await close, 1009);

  // ws may reject oversize frames before emitting a "message" event. Ensure we
  // still observe an increment in the Prometheus counter.
  await waitForCondition(
    async () => {
      const res = await fetch(`${server.httpUrl}/metrics`);
      if (!res.ok) return false;
      const body = await res.text();
      const match = body.match(
        /^sync_server_ws_messages_too_large_total(?:\{[^}]*\})?\s+([0-9.]+)$/m
      );
      if (!match) return false;
      return Number(match[1]) >= 1;
    },
    5_000,
    100
  );
});

test("drops oversized awareness JSON without closing connection", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-awareness-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_MESSAGE_BYTES: "4096",
      SYNC_SERVER_MAX_AWARENESS_STATE_BYTES: "256",
      SYNC_SERVER_MAX_AWARENESS_ENTRIES: "10",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "awareness-size-doc";
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  const providerA = new WebsocketProvider(server.wsUrl, docName, docA, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const providerB = new WebsocketProvider(server.wsUrl, docName, docB, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
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

  (providerA as any).awareness.setLocalState({ status: "small" });
  await waitForCondition(() => {
    const states = (providerB as any).awareness.getStates() as Map<number, any>;
    return states.get(clientIdA)?.status === "small";
  }, 10_000);

  const huge = "x".repeat(1024);
  (providerA as any).awareness.setLocalState({ status: "huge", huge });

  // If the server forwarded the update, providerB would have observed "huge".
  await new Promise((r) => setTimeout(r, 500));
  {
    const states = (providerB as any).awareness.getStates() as Map<number, any>;
    assert.equal(states.get(clientIdA)?.status, "small");
  }

  // A subsequent small update should still propagate, proving the socket stayed open.
  (providerA as any).awareness.setLocalState({ status: "after" });
  await waitForCondition(() => {
    const states = (providerB as any).awareness.getStates() as Map<number, any>;
    return states.get(clientIdA)?.status === "after";
  }, 10_000);
});

test("rate limits messages per document", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-doc-rate-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_MESSAGES_PER_DOC_WINDOW: "5",
      SYNC_SERVER_DOC_MESSAGE_WINDOW_MS: "60000",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const ws = new WebSocket(`${server.wsUrl}/doc-rate?token=test-token`);
  t.after(() => ws.terminate());

  await new Promise<void>((resolve, reject) => {
    ws.once("open", () => resolve());
    ws.once("error", (err) => reject(err));
  });

  const close = new Promise<number>((resolve) => {
    ws.once("close", (code) => resolve(code));
  });

  for (let i = 0; i < 10; i += 1) {
    ws.send(
      buildAwarenessMessage([
        { clientID: 123, clock: i + 1, stateJSON: JSON.stringify({ ok: true }) },
      ])
    );
  }

  assert.equal(await close, 1013);
});

test("rate limits messages per IP across connections", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-ip-rate-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_MESSAGES_PER_IP_WINDOW: "5",
      SYNC_SERVER_IP_MESSAGE_WINDOW_MS: "60000",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const wsA = new WebSocket(`${server.wsUrl}/ip-rate-a?token=test-token`);
  const wsB = new WebSocket(`${server.wsUrl}/ip-rate-b?token=test-token`);
  t.after(() => wsA.terminate());
  t.after(() => wsB.terminate());

  await Promise.all([
    new Promise<void>((resolve, reject) => {
      wsA.once("open", () => resolve());
      wsA.once("error", (err) => reject(err));
    }),
    new Promise<void>((resolve, reject) => {
      wsB.once("open", () => resolve());
      wsB.once("error", (err) => reject(err));
    }),
  ]);

  const closeA = new Promise<{ code: number; reason: string }>((resolve) => {
    wsA.once("close", (code, reason) => resolve({ code, reason: reason.toString() }));
  });
  const closeB = new Promise<{ code: number; reason: string }>((resolve) => {
    wsB.once("close", (code, reason) => resolve({ code, reason: reason.toString() }));
  });
  const firstClose = Promise.race([closeA, closeB]);

  // Send 6 total messages across two connections from the same IP. With the per-IP
  // limit set to 5, one of the connections should be closed.
  for (let i = 0; i < 3; i += 1) {
    wsA.send(
      buildAwarenessMessage([
        { clientID: 123, clock: i + 1, stateJSON: JSON.stringify({ ok: true }) },
      ])
    );
  }
  for (let i = 0; i < 3; i += 1) {
    wsB.send(
      buildAwarenessMessage([
        { clientID: 456, clock: i + 1, stateJSON: JSON.stringify({ ok: true }) },
      ])
    );
  }

  const timeout = new Promise<never>((_, reject) => {
    const timer = setTimeout(
      () => reject(new Error("Timed out waiting for per-IP rate limit close")),
      5_000
    );
    timer.unref();
  });

  const { code, reason } = await Promise.race([firstClose, timeout]);
  assert.equal(code, 1013);
  assert.equal(reason, "Rate limit exceeded");
});

test("disables ws maxPayload when SYNC_SERVER_MAX_MESSAGE_BYTES is 0", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-limit-disabled-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_MESSAGE_BYTES: "0",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.wsUrl, "limit-disabled-doc", doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  // If `ws` maxPayload were set to 0, even the initial sync messages would be
  // rejected before reaching our application-level guards.
  await waitForProviderSync(provider);
});
