import assert from "node:assert/strict";
import test from "node:test";
import * as Y from "yjs";

import { emptyDocumentState } from "../state.js";
import { YjsBranchStore } from "./YjsBranchStore.js";

function makeInitialState() {
  const state = emptyDocumentState();
  state.sheets.order = ["Sheet1"];
  state.sheets.metaById = {
    Sheet1: { id: "Sheet1", name: "Sheet1", view: { frozenRows: 0, frozenCols: 0 } },
  };
  state.cells = { Sheet1: {} };
  return state;
}

test("YjsBranchStore: sanitizes oversized drawing ids in inline commit patches", async () => {
  const doc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc: doc, payloadEncoding: "json" });
  const actor = { userId: "u1", role: "owner" };
  const docId = "doc-drawings-patch";

  await store.ensureDocument(docId, actor, makeInitialState());
  const main = await store.getBranch(docId, "main");
  assert.ok(main);

  const oversized = "x".repeat(5000);
  const evilId = "evil-patch";

  doc.transact(() => {
    const commit = new Y.Map();
    commit.set("id", evilId);
    commit.set("docId", docId);
    commit.set("parentCommitId", main.headCommitId);
    commit.set("mergeParentCommitId", null);
    commit.set("createdBy", actor.userId);
    commit.set("createdAt", 123);
    commit.set("message", "evil");
    commit.set("patch", {
      schemaVersion: 1,
      sheets: {
        metaById: {
          Sheet1: {
            id: "Sheet1",
            name: "Sheet1",
            view: {
              frozenRows: 0,
              frozenCols: 0,
              drawings: [{ id: oversized }, { id: " ok " }],
            },
          },
        },
      },
    });
    doc.getMap("branching:commits").set(evilId, commit);
  });

  const loaded = await store.getCommit(evilId);
  assert.ok(loaded);
  assert.deepEqual(
    loaded.patch.sheets.metaById.Sheet1.view.drawings.map((d) => d.id),
    ["ok"],
  );

  const state = await store.getDocumentStateAtCommit(evilId);
  assert.deepEqual(
    state.sheets.metaById.Sheet1.view.drawings.map((d) => d.id),
    ["ok"],
  );
});

test("YjsBranchStore: sanitizes oversized drawing ids in inline commit snapshots", async () => {
  const doc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc: doc, payloadEncoding: "json" });
  const actor = { userId: "u1", role: "owner" };
  const docId = "doc-drawings-snapshot";

  await store.ensureDocument(docId, actor, makeInitialState());
  const main = await store.getBranch(docId, "main");
  assert.ok(main);

  const oversized = "x".repeat(5000);
  const evilId = "evil-snapshot";

  const snapshot = makeInitialState();
  snapshot.sheets.metaById.Sheet1.view.drawings = [{ id: oversized }, { id: " ok " }];

  doc.transact(() => {
    const commit = new Y.Map();
    commit.set("id", evilId);
    commit.set("docId", docId);
    commit.set("parentCommitId", main.headCommitId);
    commit.set("mergeParentCommitId", null);
    commit.set("createdBy", actor.userId);
    commit.set("createdAt", 123);
    commit.set("message", "evil");
    commit.set("snapshot", snapshot);
    doc.getMap("branching:commits").set(evilId, commit);
  });

  const state = await store.getDocumentStateAtCommit(evilId);
  assert.deepEqual(
    state.sheets.metaById.Sheet1.view.drawings.map((d) => d.id),
    ["ok"],
  );
});

