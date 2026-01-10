import test from "node:test";
import assert from "node:assert/strict";

import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { InMemoryBranchStore } from "../packages/versioning/branches/src/store/InMemoryBranchStore.js";

test("integration: create branch, diverge, merge back", async () => {
  const actor = { userId: "u1", role: "owner" };
  const store = new InMemoryBranchStore();
  const service = new BranchService({ docId: "doc1", store });

  await service.init(actor, { sheets: { Sheet1: { A1: { value: 1 } } } });

  await service.createBranch(actor, { name: "scenario" });
  await service.checkoutBranch(actor, { name: "scenario" });
  await service.commit(actor, {
    nextState: { sheets: { Sheet1: { A1: { value: 10 }, B1: { value: 99 } } } },
    message: "Scenario tweaks"
  });

  await service.checkoutBranch(actor, { name: "main" });
  await service.commit(actor, {
    nextState: { sheets: { Sheet1: { A1: { value: 5 }, C1: { value: 7 } } } },
    message: "Mainline edit"
  });

  const preview = await service.previewMerge(actor, { sourceBranch: "scenario" });
  assert.equal(preview.conflicts.length, 1, "A1 differs, should conflict");

  const merge = await service.merge(actor, {
    sourceBranch: "scenario",
    resolutions: [{ conflictIndex: 0, choice: "theirs" }]
  });

  assert.deepEqual(merge.state.sheets.Sheet1, {
    A1: { value: 10 },
    B1: { value: 99 },
    C1: { value: 7 }
  });
});
