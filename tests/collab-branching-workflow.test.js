import assert from "node:assert/strict";
import test from "node:test";
import * as Y from "yjs";

import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { CollabBranchingWorkflow } from "../packages/collab/branching/index.js";
import { BranchService, YjsBranchStore } from "../packages/versioning/branches/src/index.js";

test("CollabBranchingWorkflow: keeps global currentBranchName consistent for rename/delete", async () => {
  const session = createCollabSession({ doc: new Y.Doc() });
  const docId = "doc1";
  const owner = { userId: "u1", role: "owner" };

  const store = new YjsBranchStore({ ydoc: session.doc });
  const branchService = new BranchService({ docId, store });
  const workflow = new CollabBranchingWorkflow({ session, branchService });

  await branchService.init(owner, { sheets: {} });

  const meta = session.doc.getMap("branching:meta");
  assert.equal(meta.get("currentBranchName"), "main");

  await workflow.createBranch(owner, { name: "feature" });
  await workflow.checkoutBranch(owner, { name: "feature" });
  assert.equal(meta.get("currentBranchName"), "feature");

  await workflow.renameBranch(owner, { oldName: "feature", newName: "feat2" });
  assert.equal(meta.get("currentBranchName"), "feat2");

  await assert.rejects(workflow.deleteBranch(owner, { name: "feat2" }), {
    message: "Cannot delete the currently checked-out branch",
  });

  await workflow.checkoutBranch(owner, { name: "main" });
  assert.equal(meta.get("currentBranchName"), "main");

  await workflow.deleteBranch(owner, { name: "feat2" });

  const branches = await workflow.listBranches();
  assert.deepEqual(
    branches.map((b) => b.name).sort(),
    ["main"],
  );
});

