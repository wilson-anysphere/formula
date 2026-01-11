import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { KeyRing } from "../packages/security/crypto/keyring.js";
import { isEncryptedFileBytes } from "../packages/security/crypto/encryptedFile.js";
import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { SQLiteBranchStore } from "../packages/versioning/branches/src/store/SQLiteBranchStore.js";

test("SQLiteBranchStore persists branches and commits (with snapshots)", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "branch-store-"));
  const storePath = path.join(tmpDir, "branches.sqlite");

  const actor = { userId: "u1", role: "owner" };

  const store = new SQLiteBranchStore({ filePath: storePath, snapshotEveryNCommits: 3 });
  const service = new BranchService({ docId: "doc1", store });

  await service.init(actor, { sheets: { Sheet1: { A1: { value: 1 } } } });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 2 } } } } });
  await service.createBranch(actor, { name: "scenario" });
  await service.checkoutBranch(actor, { name: "scenario" });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 3 } } } } });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 4 } } } } });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 5 } } } } });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 6 } } } } });
  store.close();

  const store2 = new SQLiteBranchStore({ filePath: storePath, snapshotEveryNCommits: 3 });
  const service2 = new BranchService({ docId: "doc1", store: store2 });

  // Should no-op because the document already exists in this store.
  await service2.init(actor, { sheets: {} });

  const branches = await service2.listBranches();
  assert.ok(branches.some((b) => b.name === "scenario"));

  const scenario = branches.find((b) => b.name === "scenario");
  assert.ok(scenario);
  const state = await store2.getDocumentStateAtCommit(scenario.headCommitId);
  assert.deepEqual(state.sheets.Sheet1, { A1: { value: 6 } });
  store2.close();
});

test("SQLiteBranchStore encryption: keyring required to reopen encrypted store", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "branch-store-encryption-"));
  const storePath = path.join(tmpDir, "branches.sqlite");

  const actor = { userId: "u1", role: "owner" };
  const keyRing = KeyRing.create();

  const store = new SQLiteBranchStore({
    filePath: storePath,
    encryption: { mode: "keyring", keyRing, aadContext: { scope: "test" } }
  });
  const service = new BranchService({ docId: "doc1", store });

  await service.init(actor, { sheets: { Sheet1: { A1: { value: 1 } } } });
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 2 } } } } });
  store.close();

  const encryptedBytes = await fs.readFile(storePath);
  assert.ok(isEncryptedFileBytes(encryptedBytes), "expected encrypted file format on disk");
  assert.ok(
    !encryptedBytes.toString("ascii").includes("SQLite format 3"),
    "expected ciphertext-only SQLite store"
  );

  const reopened = new SQLiteBranchStore({
    filePath: storePath,
    encryption: { mode: "keyring", keyRing, aadContext: { scope: "test" } }
  });
  const branches = await reopened.listBranches("doc1");
  assert.ok(branches.some((b) => b.name === "main"));
  const main = branches.find((b) => b.name === "main");
  assert.ok(main);
  const state = await reopened.getDocumentStateAtCommit(main.headCommitId);
  assert.deepEqual(state.sheets.Sheet1, { A1: { value: 2 } });
  reopened.close();

  const missingKeyRing = new SQLiteBranchStore({ filePath: storePath });
  await assert.rejects(missingKeyRing.listBranches("doc1"), {
    message: "Encrypted SQLiteBranchStore requires encryption.mode='keyring'",
  });
});
