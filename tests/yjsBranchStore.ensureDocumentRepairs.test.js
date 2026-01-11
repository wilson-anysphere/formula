import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore.ensureDocument repairs missing main branch + invalid currentBranchName", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  // Corrupt the doc: delete main branch and set currentBranchName to a missing branch.
  ydoc.transact(() => {
    ydoc.getMap("branching:branches").delete("main");
    ydoc.getMap("branching:meta").set("currentBranchName", "ghost");
  });

  await store.ensureDocument(docId, actor, { sheets: {} });

  assert.ok(ydoc.getMap("branching:branches").has("main"));
  assert.equal(ydoc.getMap("branching:meta").get("currentBranchName"), "main");
  assert.equal(await store.getCurrentBranchName(docId), "main");
  assert.equal(await store.hasDocument(docId), true);
  assert.ok(await store.getBranch(docId, "main"));
});
