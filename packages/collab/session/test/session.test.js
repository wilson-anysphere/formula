import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { InMemoryAwarenessHub } from "@formula/collab-presence";

import { createCollabSession } from "../src/index.ts";

const REMOTE_ORIGIN = Symbol("remote");

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function connectDocs(docA, docB) {
  const forwardA = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docB, update, REMOTE_ORIGIN);
  };
  const forwardB = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docA, update, REMOTE_ORIGIN);
  };

  docA.on("update", forwardA);
  docB.on("update", forwardB);

  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

test("CollabSession permissions gate reads/writes + mask unreadable values (API rangeRestrictions shape)", async () => {
  const session = createCollabSession({ doc: new Y.Doc() });

  session.setPermissions({
    role: "editor",
    userId: "u-editor",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        readAllowlist: ["u-owner"],
        editAllowlist: [],
      },
    ],
  });

  assert.equal(session.canReadCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);
  assert.equal(
    session.maskValueIfUnreadable({ sheetId: "Sheet1", row: 0, col: 0, value: "secret" }),
    "###"
  );

  const wrote = await session.safeSetCellValue("Sheet1:0:0", "secret");
  assert.equal(wrote, false);
  assert.equal(session.cells.has("Sheet1:0:0"), false);
});

test("CollabSession.setPermissions validates rangeRestrictions eagerly", () => {
  const session = createCollabSession({ doc: new Y.Doc() });

  assert.throws(
    () => {
      session.setPermissions({
        role: "editor",
        userId: "u-editor",
        rangeRestrictions: [
          {
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 0,
            endCol: 0,
            // Invalid: should be an array.
            readAllowlist: "u-owner",
          },
        ],
      });
    },
    (err) => {
      assert.ok(err instanceof Error);
      assert.match(err.message, /rangeRestrictions\[0\] invalid:/);
      assert.match(err.message, /restriction\.readAllowlist must be an array when provided/);
      return true;
    },
  );
});

test("CollabSession.setPermissions rejects non-array rangeRestrictions", () => {
  const session = createCollabSession({ doc: new Y.Doc() });

  assert.throws(
    () => {
      session.setPermissions({
        role: "editor",
        userId: "u-editor",
        // Misconfigured: should be an array.
        rangeRestrictions: null,
      });
    },
    (err) => {
      assert.ok(err instanceof Error);
      assert.match(err.message, /rangeRestrictions must be an array/);
      return true;
    },
  );
});

test("CollabSession.setPermissions supports legacy { range: ... } restriction shape", () => {
  const session = createCollabSession({ doc: new Y.Doc() });

  session.setPermissions({
    role: "editor",
    userId: "u-editor",
    rangeRestrictions: [
      {
        range: {
          sheetName: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 0,
          endCol: 0,
        },
        readAllowlist: [],
        editAllowlist: [],
      },
    ],
  });

  const stored = session.getPermissions()?.rangeRestrictions ?? null;
  assert.ok(Array.isArray(stored));
  assert.equal(stored.length, 1);
  const restriction = stored[0];
  assert.ok(restriction && typeof restriction === "object");
  assert.ok(restriction.range && typeof restriction.range === "object");
  assert.equal(restriction.range.sheetId, "Sheet1");
});

test("CollabSession.setPermissions error message includes the failing restriction index", () => {
  const session = createCollabSession({ doc: new Y.Doc() });

  assert.throws(
    () => {
      session.setPermissions({
        role: "editor",
        userId: "u-editor",
        rangeRestrictions: [
          {
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 0,
            endCol: 0,
            readAllowlist: [],
            editAllowlist: [],
          },
          {
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 0,
            endCol: 0,
            // Invalid: should be an array.
            editAllowlist: "u-owner",
          },
        ],
      });
    },
    (err) => {
      assert.ok(err instanceof Error);
      assert.match(err.message, /rangeRestrictions\[1\] invalid:/);
      assert.match(err.message, /restriction\.editAllowlist must be an array when provided/);
      return true;
    },
  );
});

test("CollabSession.setPermissions stores normalized rangeRestrictions (sheetName → range.sheetId)", () => {
  const session = createCollabSession({ doc: new Y.Doc() });

  session.setPermissions({
    role: "editor",
    userId: "u-editor",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        readAllowlist: [],
        editAllowlist: [],
      },
    ],
  });

  const stored = session.getPermissions()?.rangeRestrictions ?? null;
  assert.ok(Array.isArray(stored));
  assert.equal(stored.length, 1);
  const restriction = stored[0];
  assert.ok(restriction && typeof restriction === "object");
  assert.ok(restriction.range && typeof restriction.range === "object");
  assert.equal(restriction.range.sheetId, "Sheet1");
  // Ensure we didn't store the flattened input shape.
  assert.equal(restriction.startRow, undefined);
  assert.equal(restriction.sheetName, undefined);
});

test("CollabSession safeSetCell* rejects invalid cell keys (no Yjs mutation)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(session.safeSetCellValue("bad-key", "hacked"), /Invalid cellKey/);
  await assert.rejects(session.safeSetCellFormula("bad-key", "=HACK()"), /Invalid cellKey/);
  await assert.rejects(session.safeSetCellValue("", "hacked"), /Invalid cellKey/);
  await assert.rejects(session.safeSetCellFormula("", "=HACK()"), /Invalid cellKey/);
  await assert.rejects(
    // @ts-expect-error intentionally invalid type
    session.safeSetCellValue(null, "hacked"),
    /Invalid cellKey/,
  );

  assert.deepEqual(session.cells.toJSON(), {});

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession integration: sync + presence (in-memory)", async () => {
  const doc1 = new Y.Doc();
  const doc2 = new Y.Doc();
  const disconnect = connectDocs(doc1, doc2);

  const hub = new InMemoryAwarenessHub();
  const awareness1 = hub.createAwareness(1);
  const awareness2 = hub.createAwareness(2);

  const session1 = createCollabSession({
    doc: doc1,
    awareness: awareness1,
    presence: {
      user: { id: "u1", name: "User 1", color: "#ff0000" },
      activeSheet: "Sheet1",
      throttleMs: 0,
    },
  });

  const session2 = createCollabSession({
    doc: doc2,
    awareness: awareness2,
    presence: {
      user: { id: "u2", name: "User 2", color: "#00ff00" },
      activeSheet: "Sheet1",
      throttleMs: 0,
    },
  });

  await session1.setCellFormula("Sheet1:0:0", "=1");
  assert.equal((await session2.getCell("Sheet1:0:0"))?.formula, "=1");

  session1.presence?.setCursor({ row: 0, col: 0 });
  const remote = session2.presence?.getRemotePresences() ?? [];
  assert.equal(remote.length, 1);
  assert.equal(remote[0].id, "u1");
  assert.deepEqual(remote[0].cursor, { row: 0, col: 0 });

  session1.destroy();
  session2.destroy();
  disconnect();
  doc1.destroy();
  doc2.destroy();
});

test("CollabSession schema.defaultSheetId is used when normalizing cell keys without a sheet id", async () => {
  const session = createCollabSession({
    doc: new Y.Doc(),
    schema: { defaultSheetId: "Main", defaultSheetName: "Main" },
  });

  // `r{row}c{col}` keys omit sheet id; they should resolve to schema.defaultSheetId.
  const wrote = await session.safeSetCellValue("r0c0", 123);
  assert.equal(wrote, true);
  assert.equal(session.cells.has("Main:0:0"), true);
});

test("CollabSession schema init waits for provider sync even if provider.synced is unset", async () => {
  const doc = new Y.Doc();
  /** @type {Map<string, Set<(...args: any[]) => void>>} */
  const events = new Map();

  const provider = {
    on(event, cb) {
      let listeners = events.get(event);
      if (!listeners) {
        listeners = new Set();
        events.set(event, listeners);
      }
      listeners.add(cb);
    },
    off(event, cb) {
      const listeners = events.get(event);
      if (!listeners) return;
      listeners.delete(cb);
      if (listeners.size === 0) events.delete(event);
    },
    destroy() {},
  };

  const session = createCollabSession({
    doc,
    provider,
  });

  assert.equal(session.sheets.length, 0);
  const syncListeners = events.get("sync");
  assert.equal(syncListeners instanceof Set, true);
  assert.equal(syncListeners?.size ? syncListeners.size > 0 : false, true);
  for (const cb of syncListeners ?? []) cb(true);
  assert.equal(session.sheets.length, 1);
  assert.equal(session.sheets.get(0)?.get("id"), "Sheet1");

  session.destroy();
  doc.destroy();
});

test("CollabSession schema.autoInit=false does not create workbook roots eagerly", () => {
  const session = createCollabSession({ doc: new Y.Doc(), schema: { autoInit: false } });
  assert.equal(session.sheets.length, 0);
  assert.equal(session.cells.size, 0);
  assert.equal(session.namedRanges.size, 0);
  assert.equal(session.metadata.size, 0);
  session.destroy();
});
