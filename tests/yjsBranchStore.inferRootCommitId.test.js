import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore.ensureDocument infers missing rootCommitId from existing history", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  const meta = ydoc.getMap("branching:meta");
  const commits = ydoc.getMap("branching:commits");
  const originalRoot = meta.get("rootCommitId");
  assert.ok(typeof originalRoot === "string" && originalRoot.length > 0);

  const commitsBefore = commits.size;

  // Simulate a legacy/corrupted doc missing the rootCommitId metadata.
  ydoc.transact(() => {
    meta.delete("rootCommitId");
  });

  // hasDocument should still report the document exists by inferring the root commit.
  assert.equal(await store.hasDocument(docId), true);

  await store.ensureDocument(docId, actor, { sheets: {} });

  assert.equal(meta.get("rootCommitId"), originalRoot);
  assert.equal(commits.size, commitsBefore);
});

