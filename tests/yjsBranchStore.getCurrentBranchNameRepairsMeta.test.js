import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore.getCurrentBranchName repairs dangling meta pointers", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  ydoc.getMap("branching:meta").set("currentBranchName", "ghost");

  assert.equal(await store.getCurrentBranchName(docId), "main");
  assert.equal(ydoc.getMap("branching:meta").get("currentBranchName"), "main");
});

