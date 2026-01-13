import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { Y } from "../src/yjs.js";
import { startSyncServer } from "./test-helpers.ts";

function encodeVarUint(value: number): Uint8Array {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error("Invalid varUint value");
  }
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

function buildSyncUpdateMessage(update: Uint8Array): Buffer {
  // y-websocket sync message encoding:
  // - outer message type 0 (sync)
  // - inner message type 2 (update)
  // - varUint length + update bytes
  const message = concatUint8Arrays([
    encodeVarUint(0),
    encodeVarUint(2),
    encodeVarUint(update.length),
    update,
  ]);
  return Buffer.from(message);
}

async function waitForWebSocketOpen(ws: WebSocket, timeoutMs: number = 5_000): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const timeout = setTimeout(() => {
      ws.off("open", onOpen);
      ws.off("error", onError);
      reject(new Error("Timed out waiting for websocket open"));
    }, timeoutMs);
    timeout.unref();

    const onError = (err: unknown) => {
      clearTimeout(timeout);
      ws.off("open", onOpen);
      reject(err);
    };
    const onOpen = () => {
      clearTimeout(timeout);
      ws.off("error", onError);
      resolve();
    };
    ws.once("error", onError);
    ws.once("open", onOpen);
  });
}

async function waitForWebSocketClose(
  ws: WebSocket,
  timeoutMs: number = 5_000
): Promise<{ code: number; reason: Buffer }> {
  return await new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      ws.off("close", onClose);
      ws.off("error", onError);
      reject(new Error("Timed out waiting for websocket close"));
    }, timeoutMs);
    timeout.unref();

    const onError = (err: unknown) => {
      clearTimeout(timeout);
      ws.off("close", onClose);
      reject(err);
    };
    const onClose = (code: number, reason: Buffer) => {
      clearTimeout(timeout);
      ws.off("error", onError);
      resolve({ code, reason });
    };
    ws.once("error", onError);
    ws.once("close", onClose);
  });
}

async function assertWebSocketNotClosedWithin(
  ws: WebSocket,
  windowMs: number
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const timeout = setTimeout(() => {
      ws.off("close", onClose);
      ws.off("error", onError);
      resolve();
    }, windowMs);
    timeout.unref();

    const onError = (err: unknown) => {
      clearTimeout(timeout);
      ws.off("close", onClose);
      reject(err);
    };
    const onClose = (code: number, reason: Buffer) => {
      clearTimeout(timeout);
      ws.off("error", onError);
      reject(
        new Error(
          `WebSocket closed unexpectedly (code=${code}, reason=${reason.toString("utf8")})`
        )
      );
    };
    ws.once("error", onError);
    ws.once("close", onClose);
  });
}

test("reserved root guard names/prefixes are configurable via env vars", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-reserved-roots-env-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
      SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED: "true",
      // Replace defaults with a custom root name and prefix so operators can opt out
      // of guarding built-in history roots like `versions`.
      SYNC_SERVER_RESERVED_ROOT_NAMES: "customRoot",
      SYNC_SERVER_RESERVED_ROOT_PREFIXES: "customPrefix:",
    },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = "reserved-root-guard-env-doc";

  // Mutating the default reserved roots should NOT be rejected when they are removed
  // from SYNC_SERVER_RESERVED_ROOT_NAMES (guard is still enabled).
  {
    const ws = new WebSocket(`${server.wsUrl}/${docName}?token=test-token`);
    t.after(() => ws.terminate());
    await waitForWebSocketOpen(ws);

    const doc = new Y.Doc();
    doc.getMap("versions").set("v1", "allowed");
    const update = Y.encodeStateAsUpdate(doc);

    ws.send(buildSyncUpdateMessage(update));
    await assertWebSocketNotClosedWithin(ws, 250);

    ws.terminate();
  }

  // Mutating a configured reserved root name should be rejected with 1008.
  {
    const ws = new WebSocket(`${server.wsUrl}/${docName}?token=test-token`);
    t.after(() => ws.terminate());
    await waitForWebSocketOpen(ws);

    const doc = new Y.Doc();
    doc.getMap("customRoot").set("x", 1);
    const update = Y.encodeStateAsUpdate(doc);

    ws.send(buildSyncUpdateMessage(update));
    const { code } = await waitForWebSocketClose(ws);
    assert.equal(code, 1008);
  }

  // Mutating a configured reserved root prefix should be rejected with 1008.
  {
    const ws = new WebSocket(`${server.wsUrl}/${docName}?token=test-token`);
    t.after(() => ws.terminate());
    await waitForWebSocketOpen(ws);

    const doc = new Y.Doc();
    doc.getMap("customPrefix:main").set("x", 1);
    const update = Y.encodeStateAsUpdate(doc);

    ws.send(buildSyncUpdateMessage(update));
    const { code } = await waitForWebSocketClose(ws);
    assert.equal(code, 1008);
  }
});

