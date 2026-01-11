import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryBranchStore } from "../packages/versioning/branches/src/store/InMemoryBranchStore.js";

class CountingInMemoryBranchStore extends InMemoryBranchStore {
  applyPatchCalls = 0;

  _applyPatch(state, patch) {
    this.applyPatchCalls += 1;
    return super._applyPatch(state, patch);
  }
}

async function seedLinearHistory({ store, docId, commits, snapshotEveryNCommits }) {
  const actor = { userId: "u1", role: "owner" };
  await store.ensureDocument(docId, actor, { sheets: { Sheet1: { A1: { value: 0 } } } });

  const branch = await store.getBranch(docId, "main");
  assert.ok(branch);
  let head = branch.headCommitId;

  for (let i = 1; i <= commits; i += 1) {
    const patch = { sheets: { Sheet1: { A1: { value: i } } } };
    const commit = await store.createCommit({
      docId,
      parentCommitId: head,
      mergeParentCommitId: null,
      createdBy: actor.userId,
      createdAt: Date.now() + i,
      message: `c${i}`,
      patch,
      nextState: { sheets: { Sheet1: { A1: { value: i } } } },
    });
    head = commit.id;
    await store.updateBranchHead(docId, "main", head);
  }

  // Ensure we end on a non-snapshot commit so state reconstruction has to replay some patches.
  assert.notEqual(
    commits % snapshotEveryNCommits,
    0,
    "test setup expects commits to end on non-snapshot interval"
  );

  return { actor, headCommitId: head };
}

test("snapshotting bounds patch replay (InMemoryBranchStore, long history)", async () => {
  const snapshotEveryNCommits = 50;
  const commits = 249;
  const store = new CountingInMemoryBranchStore({ snapshotEveryNCommits });

  const { headCommitId } = await seedLinearHistory({
    store,
    docId: "doc1",
    commits,
    snapshotEveryNCommits,
  });

  store.applyPatchCalls = 0;
  const state = await store.getDocumentStateAtCommit(headCommitId);

  assert.deepEqual(state.cells.Sheet1, { A1: { value: commits } });
  assert.ok(
    store.applyPatchCalls <= snapshotEveryNCommits + 1,
    `expected <=${snapshotEveryNCommits + 1} applyPatch calls, got ${store.applyPatchCalls}`
  );
});
