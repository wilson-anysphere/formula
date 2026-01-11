import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";

import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { SQLiteBranchStore } from "../packages/versioning/branches/src/store/SQLiteBranchStore.js";

test("branch store back-compat: legacy cells-only patches remain readable + can merge + upgrade metadata", async (t) => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "branch-store-backcompat-"));
  const storePath = path.join(tmpDir, "branches.sqlite");

  t.after(async () => {
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  const docId = `doc-${randomUUID()}`;
  const actor = { userId: "u1", role: "owner" };

  // --- Seed a legacy history directly in SQLite ---
  const seedStore = new SQLiteBranchStore({ filePath: storePath, snapshotEveryNCommits: 3 });
  const db = await seedStore._open();

  const rootCommitId = `root-${randomUUID()}`;
  const mainBranchId = `branch-${randomUUID()}`;
  const now = Date.now();

  // Legacy v0 patch shape: { sheets: Record<sheetId, CellPatch> } where `sheets`
  // means *cells* and there is no schemaVersion.
  const legacyRootPatch = { sheets: { Sheet1: { A1: { value: 1 } } } };

  db.run("BEGIN");
  try {
    const insertCommit = db.prepare(
      `INSERT INTO commits
        (id, doc_id, parent_commit_id, merge_parent_commit_id, created_by, created_at, message, patch_json, snapshot_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    insertCommit.run([
      rootCommitId,
      docId,
      null,
      null,
      actor.userId,
      now,
      "root",
      JSON.stringify(legacyRootPatch),
      null, // Force patch application to exercise legacy patch compatibility.
    ]);
    insertCommit.free();

    const insertBranch = db.prepare(
      `INSERT INTO branches
        (id, doc_id, name, created_by, created_at, description, head_commit_id)
        VALUES (?, ?, ?, ?, ?, ?, ?)`
    );
    insertBranch.run([mainBranchId, docId, "main", actor.userId, now, null, rootCommitId]);
    insertBranch.free();

    db.run("COMMIT");
  } catch (e) {
    db.run("ROLLBACK");
    throw e;
  }

  await seedStore._queuePersist();
  seedStore.close();

  // --- Reopen the store and ensure the legacy history loads correctly ---
  const store1 = new SQLiteBranchStore({ filePath: storePath, snapshotEveryNCommits: 3 });
  const service1 = new BranchService({ docId, store: store1 });
  await service1.init(actor, { sheets: {} });

  const base = await service1.checkoutBranch(actor, { name: "main" });
  assert.equal(base.schemaVersion, 1);
  assert.equal(base.cells.Sheet1.A1.value, 1);
  assert.deepEqual(base.metadata, {});

  // --- Create branches/commits on top of the legacy root and merge ---
  await service1.createBranch(actor, { name: "scenario" });
  await service1.checkoutBranch(actor, { name: "scenario" });
  await service1.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 10 } } } } });

  await service1.checkoutBranch(actor, { name: "main" });
  await service1.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 5 } } } } });

  const preview = await service1.previewMerge(actor, { sourceBranch: "scenario" });
  assert.equal(preview.conflicts.length, 1);
  assert.equal(preview.conflicts[0].type, "cell");

  const merged = await service1.merge(actor, {
    sourceBranch: "scenario",
    resolutions: [{ conflictIndex: 0, choice: "theirs" }],
  });
  assert.equal(merged.state.cells.Sheet1.A1.value, 10);

  // --- Upgrade: commit a schemaVersion=1 state with workbook metadata ---
  const upgradedNext = structuredClone(await service1.getCurrentState());
  upgradedNext.metadata.upgraded = true;
  upgradedNext.namedRanges.NR1 = { sheetId: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } };
  upgradedNext.comments.c1 = { id: "c1", cellRef: "A1", content: "hello", resolved: false, replies: [] };

  await service1.commit(actor, { nextState: upgradedNext, message: "upgrade" });
  store1.close();

  // --- Reopen again and ensure the upgraded state persists ---
  const store2 = new SQLiteBranchStore({ filePath: storePath, snapshotEveryNCommits: 3 });
  const main = await store2.getBranch(docId, "main");
  assert.ok(main);
  const state2 = await store2.getDocumentStateAtCommit(main.headCommitId);
  assert.equal(state2.schemaVersion, 1);
  assert.equal(state2.cells.Sheet1.A1.value, 10);
  assert.equal(state2.metadata.upgraded, true);
  assert.ok(state2.namedRanges.NR1);
  assert.equal(state2.comments.c1?.content, "hello");
  store2.close();
});

