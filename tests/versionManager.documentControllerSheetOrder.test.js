import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { FileVersionStore } from "../packages/versioning/src/store/fileVersionStore.js";
import { VersionManager } from "../packages/versioning/src/versioning/versionManager.js";

test("VersionManager.restoreVersion preserves DocumentController sheet order + metadata", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "versioning-"));
  const storePath = path.join(tmpDir, "versions.json");

  const doc = new DocumentController();

  // Create sheets in a non-sorted order and then reorder them. This mimics the UI tab order
  // being distinct from sheet id alphabetical order.
  doc.getCell("S2", "A1");
  doc.getCell("S1", "A1");
  doc.getCell("S3", "A1");
  doc.reorderSheets(["S3", "S1", "S2"]);

  // Also include metadata to ensure we restore the "real" workbook navigation state.
  doc.renameSheet("S1", "Income");
  doc.setSheetVisibility("S2", "hidden");

  const checkpointOrder = doc.getSheetIds();
  assert.deepEqual(checkpointOrder, ["S3", "S1", "S2"]);
  assert.equal(doc.getSheetMeta("S1")?.name, "Income");
  assert.equal(doc.getSheetMeta("S2")?.visibility, "hidden");

  const store = new FileVersionStore({ filePath: storePath });
  const vm = new VersionManager({ doc, store, autoStart: false, user: { userId: "u1", userName: "User" } });

  const checkpoint = await vm.createCheckpoint({ name: "baseline", locked: true });

  // Mutate both order and metadata so restore is observable.
  doc.reorderSheets(["S1", "S2", "S3"]);
  doc.renameSheet("S1", "Budget");
  doc.setSheetVisibility("S2", "visible");

  await vm.createSnapshot({ description: "after edits" });

  assert.deepEqual(doc.getSheetIds(), ["S1", "S2", "S3"]);
  assert.equal(doc.getSheetMeta("S1")?.name, "Budget");
  assert.equal(doc.getSheetMeta("S2")?.visibility, "visible");

  await vm.restoreVersion(checkpoint.id);

  assert.deepEqual(doc.getSheetIds(), checkpointOrder);
  assert.equal(doc.getSheetMeta("S1")?.name, "Income");
  assert.equal(doc.getSheetMeta("S2")?.visibility, "hidden");
});

