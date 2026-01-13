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

for (const innerType of [1, 2]) {
  for (const role of ["editor", "viewer"] as const) {
    test(
      `reservedRootGuard: rejects reserved root mutation (innerType=${innerType}, role=${role})`,
      () => {
        const docName = `doc-reserved-roots-${innerType}-${role}`;
        const userId = "attacker";

        const attackerDoc = new Y.Doc();
        attackerDoc.getMap("versions").set("v1", new Y.Map());
        const update = Y.encodeStateAsUpdate(attackerDoc);

        const message = concatUint8Arrays([
          encodeVarUint(0), // sync outer message
          encodeVarUint(innerType), // SyncStep2 (1) or Update (2)
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
            role,
          },
          logger,
          ydoc: new Y.Doc(),
          limits: {
            maxMessageBytes: 10_000_000,
            maxAwarenessStateBytes: 10_000_000,
            maxAwarenessEntries: 1_000,
          },
        });

        let delivered = 0;
        ws.on("message", () => {
          delivered += 1;
        });

        ws.emit("message", message, true);
        ws.emit("message", message, true);

        assert.equal(delivered, 0);
        assert.ok((ws as any).closeCalls.length >= 1);
        assert.deepEqual((ws as any).closeCalls[0], {
          code: 1008,
          reason: "reserved root mutation",
        });

        assert.equal(warnCalls.length, 1);
        assert.equal(warnCalls[0].event, "reserved_root_mutation_rejected");
        assert.equal(warnCalls[0].obj.docName, docName);
        assert.equal(warnCalls[0].obj.userId, userId);
        assert.equal(warnCalls[0].obj.role, role);
        assert.equal(warnCalls[0].obj.root, "versions");
        assert.deepEqual(warnCalls[0].obj.keyPath, ["v1"]);
        assert.equal(Object.prototype.hasOwnProperty.call(warnCalls[0].obj, "update"), false);
        assert.equal(Object.prototype.hasOwnProperty.call(warnCalls[0].obj, "updateBytes"), false);
      }
    );

    test(
      `reservedRootGuard: rejects reserved root mutation even when root must be derived from store (innerType=${innerType}, role=${role})`,
      () => {
        const docName = `doc-reserved-roots-store-${innerType}-${role}`;
        const userId = "attacker";

        const serverDoc = new Y.Doc();
        serverDoc.getMap("versions").set("v1", "one");

        const attackerDoc = new Y.Doc();
        Y.applyUpdate(attackerDoc, Y.encodeStateAsUpdate(serverDoc));
        attackerDoc.getMap("versions").set("v1", "two");

        const update = Y.encodeStateAsUpdate(attackerDoc, Y.encodeStateVector(serverDoc));

        const message = concatUint8Arrays([
          encodeVarUint(0), // sync outer message
          encodeVarUint(innerType), // SyncStep2 (1) or Update (2)
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
            role,
          },
          logger,
          ydoc: serverDoc,
          limits: {
            maxMessageBytes: 10_000_000,
            maxAwarenessStateBytes: 10_000_000,
            maxAwarenessEntries: 1_000,
          },
        });

        let delivered = 0;
        ws.on("message", () => {
          delivered += 1;
        });

        ws.emit("message", message, true);
        ws.emit("message", message, true);

        assert.equal(delivered, 0);
        assert.ok((ws as any).closeCalls.length >= 1);
        assert.deepEqual((ws as any).closeCalls[0], {
          code: 1008,
          reason: "reserved root mutation",
        });

        assert.equal(warnCalls.length, 1);
        assert.equal(warnCalls[0].event, "reserved_root_mutation_rejected");
        assert.equal(warnCalls[0].obj.docName, docName);
        assert.equal(warnCalls[0].obj.userId, userId);
        assert.equal(warnCalls[0].obj.role, role);
        assert.equal(warnCalls[0].obj.root, "versions");
        assert.deepEqual(warnCalls[0].obj.keyPath, ["v1"]);
      }
    );

    test(
      `reservedRootGuard: rejects nested reserved root mutation (innerType=${innerType}, role=${role})`,
      () => {
        const docName = `doc-reserved-roots-nested-${innerType}-${role}`;
        const userId = "attacker";

        const serverDoc = new Y.Doc();
        const record = new Y.Map<any>();
        serverDoc.getMap("versions").set("v1", record);

        const attackerDoc = new Y.Doc();
        Y.applyUpdate(attackerDoc, Y.encodeStateAsUpdate(serverDoc));
        const recordClient = attackerDoc.getMap("versions").get("v1") as any;
        assert.ok(recordClient && typeof recordClient.set === "function");
        recordClient.set("checkpointLocked", true);

        const update = Y.encodeStateAsUpdate(attackerDoc, Y.encodeStateVector(serverDoc));

        const message = concatUint8Arrays([
          encodeVarUint(0), // sync outer message
          encodeVarUint(innerType), // SyncStep2 (1) or Update (2)
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
            role,
          },
          logger,
          ydoc: serverDoc,
          limits: {
            maxMessageBytes: 10_000_000,
            maxAwarenessStateBytes: 10_000_000,
            maxAwarenessEntries: 1_000,
          },
        });

        let delivered = 0;
        ws.on("message", () => {
          delivered += 1;
        });

        ws.emit("message", message, true);

        assert.equal(delivered, 0);
        assert.ok((ws as any).closeCalls.length >= 1);
        assert.deepEqual((ws as any).closeCalls[0], {
          code: 1008,
          reason: "reserved root mutation",
        });

        assert.equal(warnCalls.length, 1);
        assert.equal(warnCalls[0].event, "reserved_root_mutation_rejected");
        assert.equal(warnCalls[0].obj.docName, docName);
        assert.equal(warnCalls[0].obj.userId, userId);
        assert.equal(warnCalls[0].obj.role, role);
        assert.equal(warnCalls[0].obj.root, "versions");
        assert.deepEqual(warnCalls[0].obj.keyPath, ["v1", "checkpointLocked"]);
      }
    );

    test(
      `reservedRootGuard: rejects nested reserved root mutation derived from store (innerType=${innerType}, role=${role})`,
      () => {
        const docName = `doc-reserved-roots-nested-store-${innerType}-${role}`;
        const userId = "attacker";

        const serverDoc = new Y.Doc();
        serverDoc.transact(() => {
          const record = new Y.Map<any>();
          record.set("checkpointLocked", false);
          serverDoc.getMap("versions").set("v1", record);
        });

        const attackerDoc = new Y.Doc();
        Y.applyUpdate(attackerDoc, Y.encodeStateAsUpdate(serverDoc));
        const recordClient = attackerDoc.getMap("versions").get("v1") as any;
        assert.ok(recordClient && typeof recordClient.set === "function");
        recordClient.set("checkpointLocked", true);

        const update = Y.encodeStateAsUpdate(attackerDoc, Y.encodeStateVector(serverDoc));

        // Ensure this scenario actually exercises the "parent omitted, copy from origin"
        // encoding where the root/keyPath must be derived from the server store.
        const decoded = Y.decodeUpdate(update) as any;
        const sawParentOmitted =
          Array.isArray(decoded?.structs) &&
          decoded.structs.some(
            (s: any) =>
              s &&
              typeof s === "object" &&
              "id" in s &&
              "parent" in s &&
              (s as any).parent == null
          );
        assert.equal(sawParentOmitted, true);

        const message = concatUint8Arrays([
          encodeVarUint(0), // sync outer message
          encodeVarUint(innerType), // SyncStep2 (1) or Update (2)
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
            role,
          },
          logger,
          ydoc: serverDoc,
          limits: {
            maxMessageBytes: 10_000_000,
            maxAwarenessStateBytes: 10_000_000,
            maxAwarenessEntries: 1_000,
          },
        });

        let delivered = 0;
        ws.on("message", () => {
          delivered += 1;
        });

        ws.emit("message", message, true);

        assert.equal(delivered, 0);
        assert.ok((ws as any).closeCalls.length >= 1);
        assert.deepEqual((ws as any).closeCalls[0], {
          code: 1008,
          reason: "reserved root mutation",
        });

        assert.equal(warnCalls.length, 1);
        assert.equal(warnCalls[0].event, "reserved_root_mutation_rejected");
        assert.equal(warnCalls[0].obj.docName, docName);
        assert.equal(warnCalls[0].obj.userId, userId);
        assert.equal(warnCalls[0].obj.role, role);
        assert.equal(warnCalls[0].obj.root, "versions");
        assert.deepEqual(warnCalls[0].obj.keyPath, ["v1", "checkpointLocked"]);
      }
    );

    test(
      `reservedRootGuard: rejects delete-only reserved root mutation (innerType=${innerType}, role=${role})`,
      () => {
        const docName = `doc-reserved-roots-delete-${innerType}-${role}`;
        const userId = "attacker";

        const serverDoc = new Y.Doc();
        serverDoc.getMap("versionsMeta").set("order", "abc");

        const attackerDoc = new Y.Doc();
        Y.applyUpdate(attackerDoc, Y.encodeStateAsUpdate(serverDoc));
        attackerDoc.getMap("versionsMeta").delete("order");

        const update = Y.encodeStateAsUpdate(attackerDoc, Y.encodeStateVector(serverDoc));

        const message = concatUint8Arrays([
          encodeVarUint(0), // sync outer message
          encodeVarUint(innerType), // SyncStep2 (1) or Update (2)
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
            role,
          },
          logger,
          ydoc: serverDoc,
          limits: {
            maxMessageBytes: 10_000_000,
            maxAwarenessStateBytes: 10_000_000,
            maxAwarenessEntries: 1_000,
          },
        });

        let delivered = 0;
        ws.on("message", () => {
          delivered += 1;
        });

        ws.emit("message", message, true);

        assert.equal(delivered, 0);
        assert.ok((ws as any).closeCalls.length >= 1);
        assert.deepEqual((ws as any).closeCalls[0], {
          code: 1008,
          reason: "reserved root mutation",
        });

        assert.equal(warnCalls.length, 1);
        assert.equal(warnCalls[0].event, "reserved_root_mutation_rejected");
        assert.equal(warnCalls[0].obj.docName, docName);
        assert.equal(warnCalls[0].obj.userId, userId);
        assert.equal(warnCalls[0].obj.role, role);
        assert.equal(warnCalls[0].obj.root, "versionsMeta");
        assert.deepEqual(warnCalls[0].obj.keyPath, ["order"]);
      }
    );

    test(
      `reservedRootGuard: rejects reserved root prefix mutation derived from store (innerType=${innerType}, role=${role})`,
      () => {
        const docName = `doc-reserved-roots-branching-store-${innerType}-${role}`;
        const userId = "attacker";

        const serverDoc = new Y.Doc();
        serverDoc.getMap("branching:main").set("x", 1);

        const attackerDoc = new Y.Doc();
        Y.applyUpdate(attackerDoc, Y.encodeStateAsUpdate(serverDoc));
        attackerDoc.getMap("branching:main").set("x", 2);

        const update = Y.encodeStateAsUpdate(attackerDoc, Y.encodeStateVector(serverDoc));

        const message = concatUint8Arrays([
          encodeVarUint(0), // sync outer message
          encodeVarUint(innerType), // SyncStep2 (1) or Update (2)
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
            role,
          },
          logger,
          ydoc: serverDoc,
          limits: {
            maxMessageBytes: 10_000_000,
            maxAwarenessStateBytes: 10_000_000,
            maxAwarenessEntries: 1_000,
          },
        });

        let delivered = 0;
        ws.on("message", () => {
          delivered += 1;
        });

        ws.emit("message", message, true);
        ws.emit("message", message, true);

        assert.equal(delivered, 0);
        assert.ok((ws as any).closeCalls.length >= 1);
        assert.deepEqual((ws as any).closeCalls[0], {
          code: 1008,
          reason: "reserved root mutation",
        });

        assert.equal(warnCalls.length, 1);
        assert.equal(warnCalls[0].event, "reserved_root_mutation_rejected");
        assert.equal(warnCalls[0].obj.docName, docName);
        assert.equal(warnCalls[0].obj.userId, userId);
        assert.equal(warnCalls[0].obj.role, role);
        assert.equal(warnCalls[0].obj.root, "branching:main");
        assert.deepEqual(warnCalls[0].obj.keyPath, ["x"]);
      }
    );
  }
}

test("reservedRootGuard: truncates large keyPath segments in logs", () => {
  const docName = "doc-reserved-roots-truncation";
  const userId = "attacker";
  const role = "editor";

  const hugeKey = `v-${"x".repeat(2_000)}`;
  const attackerDoc = new Y.Doc();
  attackerDoc.getMap("versions").set(hugeKey, new Y.Map());
  const update = Y.encodeStateAsUpdate(attackerDoc);

  const message = concatUint8Arrays([
    encodeVarUint(0), // sync outer message
    encodeVarUint(2), // Update
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
      role,
    },
    logger,
    ydoc: new Y.Doc(),
    limits: {
      maxMessageBytes: 10_000_000,
      maxAwarenessStateBytes: 10_000_000,
      maxAwarenessEntries: 1_000,
    },
  });

  // installYwsSecurity guards messages by wrapping websocket "message" listeners.
  // Attach a listener to ensure the guard runs when we emit below.
  let delivered = 0;
  ws.on("message", () => {
    delivered += 1;
  });

  ws.emit("message", message, true);

  assert.equal(delivered, 0);
  assert.ok((ws as any).closeCalls.length >= 1);
  assert.equal((ws as any).closeCalls[0].code, 1008);
  assert.equal((ws as any).closeCalls[0].reason, "reserved root mutation");

  assert.equal(warnCalls.length, 1);
  assert.equal(warnCalls[0].event, "reserved_root_mutation_rejected");
  assert.equal(warnCalls[0].obj.root, "versions");
  assert.ok(Array.isArray(warnCalls[0].obj.keyPath));
  const loggedKey = (warnCalls[0].obj.keyPath as string[])[0];
  assert.equal(typeof loggedKey, "string");
  assert.ok(loggedKey.length < hugeKey.length);
  assert.ok(loggedKey.endsWith("..."));
});
