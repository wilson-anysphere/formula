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
