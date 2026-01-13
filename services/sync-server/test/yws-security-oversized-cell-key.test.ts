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

test("rangeRestrictions: logs oversized cell key length without logging the full key", () => {
  const docName = "doc-oversized-cell-key";
  const userId = "attacker";

  const oversizedKey = `Sheet1:${"0".repeat(1025)}`;
  assert.ok(oversizedKey.length > 1024);

  const attackerDoc = new Y.Doc();
  attackerDoc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "evil");
    attackerDoc.getMap("cells").set(oversizedKey, cell);
  });
  const update = Y.encodeStateAsUpdate(attackerDoc);

  const message = concatUint8Arrays([
    encodeVarUint(0), // sync outer message
    encodeVarUint(2), // update
    encodeVarUint(update.length),
    update,
  ]);

  const ws = new FakeWebSocket() as unknown as WebSocket;
  const warnCalls: Array<{ obj: Record<string, unknown>; event: string }> = [];
  const logger = {
    warn(obj: Record<string, unknown>, event: string) {
      warnCalls.push({ obj, event });
    },
  } as any;

  installYwsSecurity(ws, {
    docName,
    auth: {
      userId,
      tokenType: "jwt",
      docId: docName,
      orgId: null,
      role: "editor",
      rangeRestrictions: [
        {
          range: { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
          editAllowlist: [userId],
        },
      ],
    },
    logger,
    ydoc: new Y.Doc(),
    limits: {
      maxMessageBytes: 10_000_000,
      maxAwarenessStateBytes: 10_000_000,
      maxAwarenessEntries: 1_000,
    },
    enforceRangeRestrictions: true,
  });

  let delivered = 0;
  ws.on("message", () => {
    delivered += 1;
  });

  ws.emit("message", message, true);

  assert.equal(delivered, 0);
  assert.equal((ws as any).closeCalls.length, 1);
  assert.deepEqual((ws as any).closeCalls[0], {
    code: 1008,
    reason: "unparseable cell key",
  });

  assert.equal(warnCalls.length, 1);
  assert.equal(warnCalls[0].event, "range_restriction_oversized_cell_key");
  assert.equal(warnCalls[0].obj.docName, docName);
  assert.equal(warnCalls[0].obj.userId, userId);
  assert.equal(warnCalls[0].obj.role, "editor");
  assert.equal(warnCalls[0].obj.cellKeyLength, oversizedKey.length);
  assert.equal(Object.prototype.hasOwnProperty.call(warnCalls[0].obj, "cellKey"), false);
});

test("rangeRestrictions: logs unparseable cell key length without logging the full key", () => {
  const docName = "doc-unparseable-cell-key";
  const userId = "attacker";

  // Long but within MAX_CELL_KEY_CHARS; invalid because it lacks a second `:` or `,`.
  const unparseableKey = `Sheet1:${"0".repeat(1024 - "Sheet1:".length)}`;
  assert.equal(unparseableKey.length, 1024);

  const attackerDoc = new Y.Doc();
  attackerDoc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "evil");
    attackerDoc.getMap("cells").set(unparseableKey, cell);
  });
  const update = Y.encodeStateAsUpdate(attackerDoc);

  const message = concatUint8Arrays([
    encodeVarUint(0), // sync outer message
    encodeVarUint(2), // update
    encodeVarUint(update.length),
    update,
  ]);

  const ws = new FakeWebSocket() as unknown as WebSocket;
  const warnCalls: Array<{ obj: Record<string, unknown>; event: string }> = [];
  const logger = {
    warn(obj: Record<string, unknown>, event: string) {
      warnCalls.push({ obj, event });
    },
  } as any;

  installYwsSecurity(ws, {
    docName,
    auth: {
      userId,
      tokenType: "jwt",
      docId: docName,
      orgId: null,
      role: "editor",
      rangeRestrictions: [
        {
          range: { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
          editAllowlist: [userId],
        },
      ],
    },
    logger,
    ydoc: new Y.Doc(),
    limits: {
      maxMessageBytes: 10_000_000,
      maxAwarenessStateBytes: 10_000_000,
      maxAwarenessEntries: 1_000,
    },
    enforceRangeRestrictions: true,
  });

  let delivered = 0;
  ws.on("message", () => {
    delivered += 1;
  });

  ws.emit("message", message, true);

  assert.equal(delivered, 0);
  assert.equal((ws as any).closeCalls.length, 1);
  assert.deepEqual((ws as any).closeCalls[0], {
    code: 1008,
    reason: "unparseable cell key",
  });

  assert.equal(warnCalls.length, 1);
  assert.equal(warnCalls[0].event, "range_restriction_unparseable_cell");
  assert.equal(warnCalls[0].obj.docName, docName);
  assert.equal(warnCalls[0].obj.userId, userId);
  assert.equal(warnCalls[0].obj.role, "editor");
  assert.equal(warnCalls[0].obj.cellKeyLength, unparseableKey.length);
  assert.equal(Object.prototype.hasOwnProperty.call(warnCalls[0].obj, "cellKey"), false);
  assert.equal(JSON.stringify(warnCalls[0].obj).includes(unparseableKey), false);
});
