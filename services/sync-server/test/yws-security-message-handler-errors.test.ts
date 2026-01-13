import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import test from "node:test";

import type WebSocket from "ws";

import { installYwsSecurity } from "../src/ywsSecurity.js";
import { Y } from "../src/yjs.js";

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

class FakeWebSocket extends EventEmitter {
  public closeCalls: Array<{ code: number; reason: string }> = [];

  close(code: number, reason?: string): void {
    this.closeCalls.push({ code, reason: reason ?? "" });
  }
}

test("ws message handler: handler throw is caught, socket closed, and metric incremented", () => {
  const ws = new FakeWebSocket() as unknown as WebSocket;

  const warnCalls: Array<{ obj: Record<string, unknown>; event: string }> = [];
  const logger = {
    warn(obj: Record<string, unknown>, event: string) {
      warnCalls.push({ obj, event });
    },
  } as any;

  const handlerErrorStages: string[] = [];
  const metrics = {
    wsReservedRootQuotaViolationsTotal: { inc: () => {} },
    wsMessageHandlerErrorsTotal: {
      inc: (labels?: { stage?: string }) => {
        if (labels?.stage) handlerErrorStages.push(labels.stage);
      },
    },
  } as any;

  installYwsSecurity(ws, {
    docName: "doc-message-handler-throw",
    auth: undefined,
    logger,
    ydoc: new Y.Doc(),
    metrics,
    limits: {
      maxMessageBytes: 10_000_000,
      maxAwarenessStateBytes: 10_000_000,
      maxAwarenessEntries: 1_000,
    },
  });

  ws.on("message", () => {
    throw new Error("boom");
  });

  assert.doesNotThrow(() => {
    ws.emit("message", Buffer.from([0, 0]), true);
  });

  assert.equal((ws as any).closeCalls.length, 1);
  assert.deepEqual((ws as any).closeCalls[0], { code: 1011, reason: "Internal error" });

  assert.deepEqual(handlerErrorStages, ["handler"]);
  assert.equal(warnCalls.length, 1);
  assert.equal(warnCalls[0].event, "ws_message_handler_error");
  assert.deepEqual(warnCalls[0].obj, { stage: "handler" });
});

test("ws message handler: guard throw is caught, socket closed, and metric incremented", () => {
  const ws = new FakeWebSocket() as unknown as WebSocket;

  const warnCalls: Array<{ obj: Record<string, unknown>; event: string }> = [];
  const logger = {
    warn(obj: Record<string, unknown>, event: string) {
      warnCalls.push({ obj, event });
    },
  } as any;

  const handlerErrorStages: string[] = [];
  const metrics = {
    wsReservedRootQuotaViolationsTotal: { inc: () => {} },
    wsMessageHandlerErrorsTotal: {
      inc: (labels?: { stage?: string }) => {
        if (labels?.stage) handlerErrorStages.push(labels.stage);
      },
    },
  } as any;

  installYwsSecurity(ws, {
    docName: "doc-message-guard-throw",
    auth: {
      userId: "user",
      tokenType: "jwt",
      docId: "doc-message-guard-throw",
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
    __testHooks: {
      inspectUpdateForReservedRootGuard: () => {
        throw new Error("boom");
      },
    },
  });

  let delivered = 0;
  ws.on("message", () => {
    delivered += 1;
  });

  // Build a sync "Update" message (outer=0, inner=2) with an empty update payload.
  const message = concatUint8Arrays([
    encodeVarUint(0), // sync outer message
    encodeVarUint(2), // Update
    encodeVarUint(0), // empty update
  ]);

  assert.doesNotThrow(() => {
    ws.emit("message", message, true);
  });

  assert.equal(delivered, 0);
  assert.equal((ws as any).closeCalls.length, 1);
  assert.deepEqual((ws as any).closeCalls[0], { code: 1011, reason: "Internal error" });

  assert.deepEqual(handlerErrorStages, ["guard"]);
  assert.equal(warnCalls.length, 1);
  assert.equal(warnCalls[0].event, "ws_message_handler_error");
  assert.deepEqual(warnCalls[0].obj, { stage: "guard" });
});
