import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { SQLiteBranchStore } from "../packages/versioning/branches/src/store/SQLiteBranchStore.js";

test("SQLiteBranchStore persists branches and commits", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "branch-store-"));
  const storePath = path.join(tmpDir, "branches.sqlite");

  const actor = { userId: "u1", role: "owner" };

  const store = new SQLiteBranchStore({ filePath: storePath });
  const service = new BranchService({ docId: "doc1", store });

  await service.init(actor, { sheets: { Sheet1: { A1: { value: 1 } } } });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 2 } } } } });
  await service.createBranch(actor, { name: "scenario" });
  await service.checkoutBranch(actor, { name: "scenario" });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 3 } } } } });
  store.close();

  const store2 = new SQLiteBranchStore({ filePath: storePath });
  const service2 = new BranchService({ docId: "doc1", store: store2 });

  // Should no-op because the document already exists in this store.
  await service2.init(actor, { sheets: {} });

  const branches = await service2.listBranches();
  assert.ok(branches.some((b) => b.name === "scenario"));

  const scenario = branches.find((b) => b.name === "scenario");
  assert.ok(scenario);
  const state = await store2.getDocumentStateAtCommit(scenario.headCommitId);
  assert.deepEqual(state.sheets.Sheet1, { A1: { value: 3 } });
  store2.close();
});

