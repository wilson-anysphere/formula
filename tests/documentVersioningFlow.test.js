import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { SQLiteVersionStore } from "../packages/versioning/src/store/sqliteVersionStore.js";
import { VersionManager } from "../packages/versioning/src/versioning/versionManager.js";
import { diffDocumentVersionAgainstCurrent } from "../packages/versioning/src/document/versionHistory.js";

test("DocumentController: checkpoint -> edit -> compare -> restore (history preserved)", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "doc-versioning-"));
  const storePath = path.join(tmpDir, "versions.sqlite");

  const doc = new DocumentController();
  const store = new SQLiteVersionStore({ filePath: storePath });
  const vm = new VersionManager({ doc, store, autoStart: false, user: { userId: "u1", userName: "User" } });

  doc.setCellValue("Sheet1", "A1", 1);
  const checkpoint = await vm.createCheckpoint({ name: "Approved", locked: true });

  doc.setCellValue("Sheet1", "A1", 2);
  await vm.createSnapshot({ description: "edit" });

  const diff = await diffDocumentVersionAgainstCurrent({
    versionManager: vm,
    versionId: checkpoint.id,
    sheetId: "Sheet1",
  });
  assert.equal(diff.modified.length, 1);
  assert.deepEqual(diff.modified[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified[0].oldValue, 1);
  assert.equal(diff.modified[0].newValue, 2);

  await vm.restoreVersion(checkpoint.id);
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);

  const versions = await vm.listVersions();
  assert.equal(versions.length, 3);
  assert.ok(versions.some((v) => v.kind === "restore"));

  // Persistence: reload store and ensure history is still present.
  const doc2 = new DocumentController();
  const vm2 = new VersionManager({
    doc: doc2,
    store: new SQLiteVersionStore({ filePath: storePath }),
    autoStart: false,
  });
  const versions2 = await vm2.listVersions();
  assert.equal(versions2.length, 3);
});

