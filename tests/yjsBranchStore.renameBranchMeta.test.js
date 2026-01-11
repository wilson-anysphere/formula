import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore: renameBranch updates currentBranchName meta atomically", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  const main = await store.getBranch(docId, "main");
  assert.ok(main);

  await store.createBranch({
    docId,
    name: "feature",
    createdBy: actor.userId,
    createdAt: Date.now(),
    description: null,
    headCommitId: main.headCommitId,
  });

  await store.setCurrentBranchName(docId, "feature");
  assert.equal(ydoc.getMap("branching:meta").get("currentBranchName"), "feature");

  await store.renameBranch(docId, "feature", "feat2");

  assert.equal(ydoc.getMap("branching:meta").get("currentBranchName"), "feat2");
  assert.ok(await store.getBranch(docId, "feat2"));
  assert.equal(await store.getBranch(docId, "feature"), null);
});

