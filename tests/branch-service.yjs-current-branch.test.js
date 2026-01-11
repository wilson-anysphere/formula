import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { emptyDocumentState } from "../packages/versioning/branches/src/state.js";
import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("BranchService (YjsBranchStore): current branch name is shared via Yjs meta", async () => {
  const ydoc = new Y.Doc();
  const docId = "doc1";

  const storeA = new YjsBranchStore({ ydoc });
  const storeB = new YjsBranchStore({ ydoc });

  const owner = { userId: "u-owner", role: "owner" };
  const editor = { userId: "u-editor", role: "editor" };

  const serviceA = new BranchService({ docId, store: storeA });
  const serviceB = new BranchService({ docId, store: storeB });

  await serviceA.init(owner, { sheets: {} });
  await serviceA.createBranch(owner, { name: "feature" });
  await serviceA.checkoutBranch(owner, { name: "feature" });
  assert.equal(await serviceB.getCurrentBranchName(), "feature");

  // Editor cannot checkout branches, but should still commit to the globally checked-out branch.
  await assert.rejects(serviceB.checkoutBranch(editor, { name: "main" }), {
    message: "checkoutBranch requires owner/admin permissions (role=editor)",
  });

  const nextState = emptyDocumentState();
  nextState.cells.Sheet1 = { A1: { value: 1 } };
  nextState.sheets.order = ["Sheet1"];
  nextState.sheets.metaById.Sheet1 = { id: "Sheet1", name: "Sheet1" };

  await serviceB.commit(editor, {
    nextState,
    message: "editor commit",
  });

  const feature = await storeA.getBranch(docId, "feature");
  assert.ok(feature);
  const head = await storeA.getCommit(feature.headCommitId);
  assert.equal(head?.message, "editor commit");
});
