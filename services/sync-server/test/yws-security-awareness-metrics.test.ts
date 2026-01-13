import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import test from "node:test";

import type WebSocket from "ws";

import { installYwsSecurity } from "../src/ywsSecurity.js";
import { Y } from "../src/yjs.js";

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

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

function readVarUint(buf: Uint8Array, offset: number): { value: number; offset: number } {
  let value = 0;
  let multiplier = 1;
  while (true) {
    if (offset >= buf.length) {
      throw new Error("Unexpected end of buffer while reading varUint");
    }
    const byte = buf[offset++];
    value += (byte & 0x7f) * multiplier;
    if (byte < 0x80) break;
    multiplier *= 0x80;
    if (!Number.isSafeInteger(value)) {
      throw new Error("varUint exceeds safe integer range");
    }
  }
  return { value, offset };
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
  const bytes = textEncoder.encode(value);
  return concatUint8Arrays([encodeVarUint(bytes.length), bytes]);
}

type AwarenessEntry = { clientID: number; clock: number; stateJSON: string };

function encodeAwarenessUpdate(entries: AwarenessEntry[]): Uint8Array {
  const chunks: Uint8Array[] = [encodeVarUint(entries.length)];
  for (const entry of entries) {
    chunks.push(encodeVarUint(entry.clientID));
    chunks.push(encodeVarUint(entry.clock));
    chunks.push(encodeVarString(entry.stateJSON));
  }
  return concatUint8Arrays(chunks);
}

function encodeAwarenessMessage(entries: AwarenessEntry[]): Uint8Array {
  const update = encodeAwarenessUpdate(entries);
  return concatUint8Arrays([encodeVarUint(1), encodeVarUint(update.length), update]);
}

function readVarString(buf: Uint8Array, offset: number): { value: string; offset: number } {
  const lenRes = readVarUint(buf, offset);
  const length = lenRes.value;
  const start = lenRes.offset;
  const end = start + length;
  if (end > buf.length) {
    throw new Error("Unexpected end of buffer while reading varString");
  }
  return { value: textDecoder.decode(buf.subarray(start, end)), offset: end };
}

function decodeAwarenessUpdate(update: Uint8Array): AwarenessEntry[] {
  let offset = 0;
  const countRes = readVarUint(update, offset);
  const count = countRes.value;
  offset = countRes.offset;

  const entries: AwarenessEntry[] = [];
  for (let i = 0; i < count; i += 1) {
    const clientRes = readVarUint(update, offset);
    const clientID = clientRes.value;
    offset = clientRes.offset;

    const clockRes = readVarUint(update, offset);
    const clock = clockRes.value;
    offset = clockRes.offset;

    const stateRes = readVarString(update, offset);
    const stateJSON = stateRes.value;
    offset = stateRes.offset;

    entries.push({ clientID, clock, stateJSON });
  }
  return entries;
}

function decodeAwarenessMessage(message: Uint8Array): AwarenessEntry[] {
  const outerRes = readVarUint(message, 0);
  assert.equal(outerRes.value, 1);
  const lenRes = readVarUint(message, outerRes.offset);
  const payloadLen = lenRes.value;
  const payloadStart = lenRes.offset;
  const payloadEnd = payloadStart + payloadLen;
  assert.ok(payloadEnd <= message.length);
  return decodeAwarenessUpdate(message.subarray(payloadStart, payloadEnd));
}

class FakeWebSocket extends EventEmitter {
  public closeCalls: Array<{ code: number; reason: string }> = [];

  close(code: number, reason?: string): void {
    this.closeCalls.push({ code, reason: reason ?? "" });
  }
}

test("awareness spoof attempts: increments metric and filters other clientIDs", () => {
  const docName = "doc-awareness-spoof-metrics";
  const userId = "user-1";

  const ws = new FakeWebSocket() as unknown as WebSocket;

  const warnCalls: Array<{ obj: Record<string, unknown>; event: string }> = [];
  const logger = {
    warn(obj: Record<string, unknown>, event: string) {
      warnCalls.push({ obj, event });
    },
  } as any;

  let spoofAttempts = 0;
  const metrics = {
    wsReservedRootQuotaViolationsTotal: { inc() {} },
    wsAwarenessSpoofAttemptsTotal: {
      inc() {
        spoofAttempts += 1;
      },
    },
    wsAwarenessClientIdCollisionsTotal: { inc() {} },
  } as any;

  installYwsSecurity(ws, {
    docName,
    auth: {
      userId,
      tokenType: "jwt",
      docId: docName,
      orgId: null,
      role: "editor",
    },
    logger,
    ydoc: new Y.Doc(),
    metrics,
    limits: {
      maxMessageBytes: 10_000_000,
      maxAwarenessStateBytes: 10_000_000,
      maxAwarenessEntries: 1_000,
    },
  });

  let delivered: unknown = null;
  ws.on("message", (data) => {
    delivered = data;
  });

  const message = encodeAwarenessMessage([
    {
      clientID: 1,
      clock: 0,
      stateJSON: JSON.stringify({ presence: { id: "spoofed" }, userId: "spoofed" }),
    },
    {
      clientID: 2,
      clock: 0,
      stateJSON: JSON.stringify({ presence: { id: "other" }, userId: "other" }),
    },
  ]);

  ws.emit("message", message, true);

  assert.equal(spoofAttempts, 1);
  assert.equal(warnCalls.length, 1);
  assert.equal(warnCalls[0].event, "awareness_spoof_attempt_filtered");

  assert.ok(delivered instanceof Uint8Array || Buffer.isBuffer(delivered));
  const buf = Buffer.isBuffer(delivered) ? delivered : Buffer.from(delivered as Uint8Array);
  const entries = decodeAwarenessMessage(
    new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength)
  );

  assert.equal(entries.length, 1);
  assert.equal(entries[0].clientID, 1);
});

test("awareness clientID collision: closes second socket and increments metric", () => {
  const docName = "doc-awareness-client-id-collision-metrics";
  const clientID = 42;

  const warnCalls: Array<{ obj: Record<string, unknown>; event: string }> = [];
  const logger = {
    warn(obj: Record<string, unknown>, event: string) {
      warnCalls.push({ obj, event });
    },
  } as any;

  let collisions = 0;
  const metrics = {
    wsReservedRootQuotaViolationsTotal: { inc() {} },
    wsAwarenessSpoofAttemptsTotal: { inc() {} },
    wsAwarenessClientIdCollisionsTotal: {
      inc() {
        collisions += 1;
      },
    },
  } as any;

  const ws1 = new FakeWebSocket() as unknown as WebSocket;
  installYwsSecurity(ws1, {
    docName,
    auth: {
      userId: "user-a",
      tokenType: "jwt",
      docId: docName,
      orgId: null,
      role: "editor",
    },
    logger,
    ydoc: new Y.Doc(),
    metrics,
    limits: {
      maxMessageBytes: 10_000_000,
      maxAwarenessStateBytes: 10_000_000,
      maxAwarenessEntries: 1_000,
    },
  });
  let delivered1 = 0;
  ws1.on("message", () => {
    delivered1 += 1;
  });

  const ws2 = new FakeWebSocket() as unknown as WebSocket;
  installYwsSecurity(ws2, {
    docName,
    auth: {
      userId: "user-b",
      tokenType: "jwt",
      docId: docName,
      orgId: null,
      role: "editor",
    },
    logger,
    ydoc: new Y.Doc(),
    metrics,
    limits: {
      maxMessageBytes: 10_000_000,
      maxAwarenessStateBytes: 10_000_000,
      maxAwarenessEntries: 1_000,
    },
  });
  let delivered2 = 0;
  ws2.on("message", () => {
    delivered2 += 1;
  });

  const message = encodeAwarenessMessage([
    {
      clientID,
      clock: 0,
      stateJSON: "{}",
    },
  ]);

  ws1.emit("message", message, true);
  assert.equal(delivered1, 1);
  assert.equal((ws1 as any).closeCalls.length, 0);

  ws2.emit("message", message, true);
  assert.equal(delivered2, 0);
  assert.equal((ws2 as any).closeCalls.length, 1);
  assert.deepEqual((ws2 as any).closeCalls[0], {
    code: 1008,
    reason: "awareness clientID collision",
  });
  assert.equal(collisions, 1);

  assert.ok(warnCalls.some((call) => call.event === "awareness_client_id_collision"));
});

